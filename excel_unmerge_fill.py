#!/usr/bin/env python3
"""拆分 Excel 合并行并填充原值。Python 3.8+，只使用标准库。"""

import argparse
import copy
import heapq
import posixpath
import re
import shutil
import subprocess
import sys
from contextlib import ExitStack
from io import StringIO
from pathlib import Path
from tempfile import TemporaryFile
from xml.dom import Node, XML_NAMESPACE, XMLNS_NAMESPACE, expatbuilder, minidom
from zipfile import ZipFile


# 以解压后的工作表 XML 大小判断，避免高压缩率文件进入完整 DOM。
# 小表继续走既有路径；超过 8 MiB 时按行落盘，限制单元格树的驻留量。
STREAM_THRESHOLD = 8 * 1024 * 1024


def children(element, name):
    return [n for n in element.childNodes
            if n.nodeType == Node.ELEMENT_NODE and n.localName == name]


def coordinate(ref):
    match = re.fullmatch(r"\$?([A-Z]+)\$?([1-9][0-9]*)", ref.upper())
    if not match:
        raise ValueError("无效的单元格地址：" + ref)
    col = 0
    for letter in match[1]:
        col = col * 26 + ord(letter) - 64
    row = int(match[2])
    if col > 16384 or row > 1048576:
        raise ValueError("单元格地址超出 Excel 范围：" + ref)
    return row, col


def address(row, col):
    letters = ""
    while col:
        col, rest = divmod(col - 1, 26)
        letters = chr(65 + rest) + letters
    return letters + str(row)


def bounds(ref):
    parts = ref.split(":")
    if len(parts) > 2:
        raise ValueError("无效的区域：" + ref)
    first, last = coordinate(parts[0]), coordinate(parts[-1])
    if first[0] > last[0] or first[1] > last[1]:
        raise ValueError("区域起止位置无效：" + ref)
    return first[0], first[1], last[0], last[1]


def reject_overlapping_merges(items):
    """按行扫描活动区域；仅拒绝本次拆分涉及的相交合并。"""
    active, selected, endings = {}, {}, []
    for index, (node, box, chosen) in sorted(enumerate(items), key=lambda item: item[1][1][0]):
        r1, c1, r2, c2 = box
        while endings and endings[0][0] < r1:
            _, expired = heapq.heappop(endings)
            active.pop(expired, None)
            selected.pop(expired, None)
        candidates = active.values() if chosen else selected.values()
        for other, (_, left, _, right), _ in candidates:
            if c1 <= right and left <= c2:
                raise ValueError("合并区域 " + other.getAttribute("ref") + " 与 "
                                 + node.getAttribute("ref") + " 相互重叠，填充值依赖处理顺序，请先在 Excel 中核对。")
        active[index] = (node, box, chosen)
        if chosen:
            selected[index] = (node, box, chosen)
        heapq.heappush(endings, (r2, index))


def namespace_bindings(element):
    """返回元素当前作用域中的声明，靠近元素的声明优先。"""
    bindings = {}
    while element is not None:
        if element.nodeType == Node.ELEMENT_NODE:
            for attribute in element.attributes.values():
                if attribute.namespaceURI == XMLNS_NAMESPACE:
                    bindings.setdefault(attribute.name, attribute.value)
        element = element.parentNode
    bindings.setdefault("xmlns", "")
    return bindings


def new_element(doc, name, parent):
    root = doc.documentElement
    qualified = root.prefix + ":" + name if root.prefix else name
    element = doc.createElementNS(root.namespaceURI, qualified)
    declaration = "xmlns:" + root.prefix if root.prefix else "xmlns"
    if namespace_bindings(parent).get(declaration) != (root.namespaceURI or ""):
        element.setAttributeNS(XMLNS_NAMESPACE, declaration, root.namespaceURI or "")
    return element


def xml_space(element):
    """xml:space 可从单元格或行继承；未声明时按 default 处理。"""
    while element is not None:
        if element.nodeType == Node.ELEMENT_NODE and element.hasAttributeNS(XML_NAMESPACE, "space"):
            return element.getAttributeNS(XML_NAMESPACE, "space")
        element = element.parentNode
    return "default"


def clone_payload(child, target):
    """把源前缀和空白处理作用域补到复制子树上。"""
    cloned = child.cloneNode(True)
    destination = namespace_bindings(target)
    for declaration, namespace in namespace_bindings(child).items():
        if not cloned.hasAttribute(declaration) and destination.get(declaration) != namespace:
            cloned.setAttributeNS(XMLNS_NAMESPACE, declaration, namespace)
    source_space = xml_space(child)
    if source_space != xml_space(target):
        cloned.setAttributeNS(XML_NAMESPACE, "xml:space", source_space)
    return cloned


def payload(cell):
    if cell is None:
        return False
    if children(cell, "f"):
        return True
    # 空的 <v/> / <is><t/></is> 是空白；"0"、" " 都是原有内容。
    values = children(cell, "v") + list(cell.getElementsByTagNameNS(cell.namespaceURI, "t"))
    return any(n.nodeValue for value in values for n in value.childNodes
               if n.nodeType in (Node.TEXT_NODE, Node.CDATA_SECTION_NODE))


def has_metadata(cell):
    """这些值依赖额外关联，不能按普通 v/is 复制或视作空白覆盖。"""
    return cell is not None and (
        any(cell.hasAttribute(attr) for attr in ("vm", "cm"))
        or bool(children(cell, "extLst")))


def reorder_children(parent, ordered):
    """一次重建 minidom 子节点序列，避免逐项 removeChild 的平方开销。"""
    parent.childNodes[:] = ordered
    previous = None
    for child in ordered:
        child.parentNode = parent
        child.previousSibling = previous
        if previous is not None:
            previous.nextSibling = child
        previous = child
    if previous is not None:
        previous.nextSibling = None


def serialize_xml(document, stream=None, insertions=None):
    """序列化文档或节点；可把落盘的行插回骨架，不累积整表字符串。"""
    def escape(value, attribute=False):
        value = (value.replace("&", "&amp;").replace("<", "&lt;")
                 .replace(">", "&gt;").replace("\r", "&#13;"))
        if attribute:
            value = (value.replace('"', "&quot;").replace("\n", "&#10;")
                     .replace("\t", "&#9;"))
        return value

    output = StringIO() if stream is None else None
    write = output.write if output is not None else lambda text: stream.write(text.encode("utf-8"))
    insertions = insertions or {}
    if document.nodeType == Node.DOCUMENT_NODE:
        declaration = '<?xml version="' + (document.version or "1.0") + '" encoding="utf-8"'
        if document.standalone is not None:
            declaration += ' standalone="' + ("yes" if document.standalone else "no") + '"'
        write(declaration + "?>")
        nodes = document.childNodes
    else:
        nodes = [document]
    # 显式栈避免工作表扩展 XML 的层级增加 Python 递归深度。
    pending = list(reversed(nodes))
    while pending:
        node = pending.pop()
        if callable(node):
            node()
        elif isinstance(node, str):
            write(node)
        elif node.nodeType == Node.ELEMENT_NODE:
            write("<" + node.tagName)
            for attribute in node.attributes.values():
                write(' ' + attribute.name + '="' + escape(attribute.value, True) + '"')
            before, after = insertions.get(node, (None, None))
            if node.childNodes or before or after:
                write(">")
                pending.append("</" + node.tagName + ">")
                if after:
                    pending.append(after)
                pending.extend(reversed(node.childNodes))
                if before:
                    pending.append(before)
            else:
                write("/>")
        elif node.nodeType == Node.TEXT_NODE:
            write(escape(node.data))
        elif node.nodeType == Node.CDATA_SECTION_NODE:
            data = node.data.replace("]]>", "]]]]><![CDATA[>")
            write("<![CDATA[" + data.replace("\r", "]]>&#13;<![CDATA[") + "]]>")
        else:
            # 注释、处理指令和文档类型不含需重新转义的属性/文本值。
            write(node.toxml())
    if output is not None:
        return output.getvalue().encode("utf-8")


def transform_sheet(raw, all_merges=False):
    """返回 (XML bytes, 拆分区域数, 新填充单元格数)。不修改普通空白。"""
    doc = minidom.parseString(raw)
    try:
        root = doc.documentElement
        merge_groups = children(root, "mergeCells")
        if not merge_groups:
            return raw, 0, 0
        merges = merge_groups[0]
        selected, merge_boxes = [], []
        for merge in children(merges, "mergeCell"):
            box = bounds(merge.getAttribute("ref"))
            chosen = box[2] > box[0] or (all_merges and box[3] > box[1])
            merge_boxes.append((merge, box, chosen))
            if chosen:
                selected.append((merge, box))
        if not selected:
            return raw, 0, 0
        reject_overlapping_merges(merge_boxes)
        data = children(root, "sheetData")[0]
        row_nodes = {int(row.getAttribute("r")): row for row in children(data, "row")}
        cells, cells_by_row = {}, {}
        for row_node in row_nodes.values():
            for cell in children(row_node, "c"):
                ref = cell.getAttribute("r")
                row, col = coordinate(ref)
                # Excel 接受小写和绝对引用；只规范化索引，保留原格的 r 属性。
                canonical = ref.replace("$", "").upper()
                if canonical in cells:
                    raise ValueError("工作表含重复单元格地址 " + canonical + "，请先在 Excel 中核对。")
                cells[canonical] = cell
                cells_by_row.setdefault(row, {})[col] = cell
        changed_rows = set()
        filled = 0
        for merge, (r1, c1, r2, c2) in selected:
            anchor_ref = address(r1, c1)
            anchor = cells.get(anchor_ref)
            if anchor is not None and children(anchor, "f"):
                raise ValueError("合并区域 " + merge.getAttribute("ref")
                                 + " 含公式，请先在 Excel 中复制并粘贴为值后重试。")
            if has_metadata(anchor):
                raise ValueError("合并区域 " + merge.getAttribute("ref")
                                 + " 含单元格图片或扩展元数据，暂不支持拆分填充。"
                                 + "请先在副本中人工核对并转换为普通单元格值后重试。")
            # 先检查被合并掩盖的内容，避免静默覆盖原有数据。
            # 只访问实际存在的单元格，宽而稀疏的合并块不枚举整片空白。
            for row in range(r1, r2 + 1):
                for col, cell in cells_by_row.get(row, {}).items():
                    if (c1 <= col <= c2 and (row, col) != (r1, c1)
                            and (payload(cell) or has_metadata(cell))):
                        raise ValueError("合并区域内的 " + cell.getAttribute("r")
                                         + " 仍有独立内容或元数据，请先核对后再处理。")
            anchor_has_value = payload(anchor)
            for row in range(r1, r2 + 1):
                if row not in row_nodes:
                    row_node = new_element(doc, "row", data)
                    row_node.setAttribute("r", str(row))
                    row_nodes[row] = row_node
                    data.appendChild(row_node)
                row_node = row_nodes[row]
                changed_rows.add(row)
                # 默认只拆纵向；矩形区域每一行仍保留横向合并。
                columns = range(c1, c2 + 1) if all_merges else (c1,)
                for col in columns:
                    ref = address(row, col)
                    if ref == anchor_ref:
                        continue
                    old = cells.get(ref)
                    cell = new_element(doc, "c", row_node)
                    cell.setAttribute("r", ref)
                    if old is not None:
                        row_node.replaceChild(cell, old)
                    else:
                        row_node.appendChild(cell)
                    if anchor is not None:
                        for attr in ("s", "t"):
                            if anchor.hasAttribute(attr):
                                cell.setAttribute(attr, anchor.getAttribute(attr))
                        for child in anchor.childNodes:
                            if (child.nodeType == Node.ELEMENT_NODE
                                    and child.localName in ("v", "is")):
                                cell.appendChild(clone_payload(child, cell))
                    cells[ref] = cell
                    cells_by_row.setdefault(row, {})[col] = cell
                    if anchor_has_value:
                        filled += 1
                if not all_merges and c2 > c1:
                    horizontal = new_element(doc, "mergeCell", merges)
                    horizontal.setAttribute("ref", address(row, c1) + ":" + address(row, c2))
                    merges.appendChild(horizontal)
            merges.removeChild(merge)

        # OOXML 要求行、单元格按坐标递增排列；新增行保留其他行属性。
        for row in sorted(changed_rows):
            row_node = row_nodes[row]
            if row_node.hasAttribute("spans"):
                row_node.removeAttribute("spans")
            ordered_cells = sorted(children(row_node, "c"),
                                   key=lambda c: coordinate(c.getAttribute("r"))[1])
            other_nodes = [n for n in row_node.childNodes
                           if n.nodeType != Node.ELEMENT_NODE or n.localName != "c"]
            reorder_children(row_node, ordered_cells + other_nodes)
        other_nodes = [n for n in data.childNodes
                       if n.nodeType != Node.ELEMENT_NODE or n.localName != "row"]
        reorder_children(data, [row_nodes[row] for row in sorted(row_nodes)] + other_nodes)
        remaining = children(merges, "mergeCell")
        if remaining:
            merges.setAttribute("count", str(len(remaining)))
        else:
            root.removeChild(merges)
        dimensions = children(root, "dimension")
        if dimensions:
            r1, c1, r2, c2 = bounds(dimensions[0].getAttribute("ref"))
            for _, (a, b, c, d) in selected:
                r1, c1, r2, c2 = min(r1, a), min(c1, b), max(r2, c), max(c2, d)
            # 部分导出文件错误地只声明 A1；流式读取器会据此截断普通数据。
            # 保留原范围，并覆盖全部实体格和仍保留的合并区域，不只看拆分目标。
            for row, columns in cells_by_row.items():
                r1, r2 = min(r1, row), max(r2, row)
                c1, c2 = min(c1, min(columns)), max(c2, max(columns))
            for merge in remaining:
                a, b, c, d = bounds(merge.getAttribute("ref"))
                r1, c1, r2, c2 = min(r1, a), min(c1, b), max(r2, c), max(c2, d)
            dimensions[0].setAttribute("ref", address(r1, c1) + ":" + address(r2, c2))
        return serialize_xml(doc), len(selected), filled
    finally:
        doc.unlink()


class RowSpoolBuilder(expatbuilder.ExpatBuilderNS):
    """使用与 minidom 相同的解析器，每个完整行落盘后立即释放其 DOM。"""
    def __init__(self, spool):
        super().__init__()
        self.spool = spool
        self.rows = {}
        self.extent = None
        self.row_error = None

    def end_element_handler(self, name):
        node = self.curNode
        super().end_element_handler(name)
        parent = node.parentNode
        if (node.localName != "row" or parent is None or parent.localName != "sheetData"
                or parent.parentNode is not self.document.documentElement):
            return
        try:
            if self.row_error is None:
                self.store_row(node)
        except ValueError as error:
            # mergeCells 通常在 sheetData 后；无匹配时与旧路径一样不校验行/格坐标。
            self.row_error = error
        finally:
            parent.removeChild(node)
            node.unlink()

    def store_row(self, node):
        number = int(node.getAttribute("r"))
        if number in self.rows:
            raise ValueError("工作表含重复行号 " + str(number) + "，请先在 Excel 中核对。")
        seen = set()
        for cell in children(node, "c"):
            row, col = coordinate(cell.getAttribute("r"))
            if row != number:
                raise ValueError("单元格 " + cell.getAttribute("r") + " 与所在行编号不一致，请先核对。")
            if col in seen:
                raise ValueError("工作表含重复单元格地址 " + address(row, col) + "，请先在 Excel 中核对。")
            seen.add(col)
            if self.extent is None:
                self.extent = (row, col, row, col)
            else:
                a, b, c, d = self.extent
                self.extent = min(a, row), min(b, col), max(c, row), max(d, col)
        raw = serialize_xml(node)
        self.rows[number] = (self.spool.tell(), len(raw))
        self.spool.write(raw)


def opening_xml(element):
    clone = element.cloneNode(False)
    try:
        return serialize_xml(clone)[:-2] + b">"
    finally:
        clone.unlink()


def detached_anchor(anchor):
    """锚点离开原行后仍携带其有效 namespace 和 xml:space 作用域。"""
    if anchor is None:
        return None
    cloned = anchor.cloneNode(True)
    for declaration, namespace in namespace_bindings(anchor).items():
        if not cloned.hasAttribute(declaration):
            cloned.setAttributeNS(XMLNS_NAMESPACE, declaration, namespace)
    cloned.setAttributeNS(XML_NAMESPACE, "xml:space", xml_space(anchor))
    return cloned


def transform_sheet_stream(source, output, all_merges=False):
    """两遍行处理，输出只写入临时流；返回 (拆分区域数, 新填充格数)。"""
    doc = None
    plans = []
    with TemporaryFile(mode="w+b") as original_rows, TemporaryFile(mode="w+b") as result_rows:
        builder = RowSpoolBuilder(original_rows)
        try:
            doc = builder.parseFile(source)
            root = doc.documentElement
            groups = children(root, "mergeCells")
            if not groups:
                return 0, 0
            merges = groups[0]
            merge_boxes = []
            for merge in children(merges, "mergeCell"):
                box = bounds(merge.getAttribute("ref"))
                chosen = box[2] > box[0] or (all_merges and box[3] > box[1])
                merge_boxes.append((merge, box, chosen))
                if chosen:
                    plans.append({"node": merge, "box": box, "anchor": None, "has_value": False})
            if not plans:
                return 0, 0
            reject_overlapping_merges(merge_boxes)
            if builder.row_error is not None:
                raise builder.row_error
            data = children(root, "sheetData")[0]
            prefix = (b'<?xml version="1.0" encoding="utf-8"?>'
                      + opening_xml(root) + opening_xml(data))
            suffix = ("</" + data.tagName + "></" + root.tagName + ">").encode("utf-8")
            # 行索引只与行数有关，不为每个实体格保留全局 DOM/字典。
            row_numbers = set(builder.rows)
            for plan in plans:
                a, _, c, _ = plan["box"]
                row_numbers.update(range(a, c + 1))
            starts = sorted(plans, key=lambda plan: plan["box"][0])
            next_start, active, filled = 0, [], 0
            for number in sorted(row_numbers):
                still_active = []
                for plan in active:
                    if plan["box"][2] >= number:
                        still_active.append(plan)
                    elif plan["anchor"] is not None:
                        plan["anchor"].unlink()
                        plan["anchor"] = None
                active = still_active
                while next_start < len(starts) and starts[next_start]["box"][0] == number:
                    active.append(starts[next_start])
                    next_start += 1
                record = builder.rows.get(number)
                if record is not None:
                    original_rows.seek(record[0])
                    raw = original_rows.read(record[1])
                else:
                    raw = b""
                if not active:
                    result_rows.write(raw)
                    continue
                row_doc = minidom.parseString(prefix + raw + suffix)
                try:
                    row_data = children(row_doc.documentElement, "sheetData")[0]
                    rows = children(row_data, "row")
                    if rows:
                        row_node = rows[0]
                    else:
                        row_node = new_element(row_doc, "row", row_data)
                        row_node.setAttribute("r", str(number))
                        row_data.appendChild(row_node)
                    cells = {coordinate(cell.getAttribute("r"))[1]: cell
                             for cell in children(row_node, "c")}
                    # 每个合并锚点只保留一份，先保存上下文，再释放原行。
                    for plan in active:
                        r1, c1, _, _ = plan["box"]
                        if r1 != number:
                            continue
                        anchor = cells.get(c1)
                        ref = plan["node"].getAttribute("ref")
                        if anchor is not None and children(anchor, "f"):
                            raise ValueError("合并区域 " + ref + " 含公式，请先在 Excel 中复制并粘贴为值后重试。")
                        if has_metadata(anchor):
                            raise ValueError("合并区域 " + ref + " 含单元格图片或扩展元数据，暂不支持拆分填充。"
                                             "请先在副本中人工核对并转换为普通单元格值后重试。")
                        plan["anchor"] = detached_anchor(anchor)
                        plan["has_value"] = payload(anchor)
                    for plan in active:
                        r1, c1, _, c2 = plan["box"]
                        for col, cell in cells.items():
                            if (c1 <= col <= c2 and (number, col) != (r1, c1)
                                    and (payload(cell) or has_metadata(cell))):
                                raise ValueError("合并区域内的 " + cell.getAttribute("r")
                                                 + " 仍有独立内容或元数据，请先核对后再处理。")
                    for plan in active:
                        r1, c1, _, c2 = plan["box"]
                        anchor = plan["anchor"]
                        columns = range(c1, c2 + 1) if all_merges else (c1,)
                        for col in columns:
                            if (number, col) == (r1, c1):
                                continue
                            cell = new_element(row_doc, "c", row_node)
                            cell.setAttribute("r", address(number, col))
                            old = cells.get(col)
                            if old is not None:
                                row_node.replaceChild(cell, old)
                            else:
                                row_node.appendChild(cell)
                            if anchor is not None:
                                for attribute in ("s", "t"):
                                    if anchor.hasAttribute(attribute):
                                        cell.setAttribute(attribute, anchor.getAttribute(attribute))
                                for child in anchor.childNodes:
                                    if child.nodeType == Node.ELEMENT_NODE and child.localName in ("v", "is"):
                                        cell.appendChild(clone_payload(child, cell))
                            cells[col] = cell
                            filled += int(plan["has_value"])
                    if row_node.hasAttribute("spans"):
                        row_node.removeAttribute("spans")
                    other = [node for node in row_node.childNodes
                             if node.nodeType != Node.ELEMENT_NODE or node.localName != "c"]
                    reorder_children(row_node, [cells[col] for col in sorted(cells)] + other)
                    result_rows.write(serialize_xml(row_node))
                finally:
                    row_doc.unlink()
            for plan in plans:
                merges.removeChild(plan["node"])
            remaining = children(merges, "mergeCell")
            horizontal_count = sum(c - a + 1 for a, b, c, d in (plan["box"] for plan in plans)
                                   if not all_merges and d > b)
            if remaining or horizontal_count:
                merges.setAttribute("count", str(len(remaining) + horizontal_count))
            else:
                root.removeChild(merges)
            dimensions = children(root, "dimension")
            if dimensions:
                r1, c1, r2, c2 = bounds(dimensions[0].getAttribute("ref"))
                boxes = [plan["box"] for plan in plans] + [bounds(node.getAttribute("ref")) for node in remaining]
                if builder.extent is not None:
                    boxes.append(builder.extent)
                for a, b, c, d in boxes:
                    r1, c1, r2, c2 = min(r1, a), min(c1, b), max(r2, c), max(c2, d)
                dimensions[0].setAttribute("ref", address(r1, c1) + ":" + address(r2, c2))

            def copy_rows():
                result_rows.seek(0)
                shutil.copyfileobj(result_rows, output, 1024 * 1024)

            def write_horizontal_merges():
                # 延迟逐项生成，避免宽矩形拆分后的 mergeCells 重新形成大 DOM。
                node = new_element(doc, "mergeCell", merges)
                try:
                    for plan in plans:
                        a, b, c, d = plan["box"]
                        if not all_merges and d > b:
                            for row in range(a, c + 1):
                                node.setAttribute("ref", address(row, b) + ":" + address(row, d))
                                output.write(serialize_xml(node))
                finally:
                    node.unlink()

            insertions = {data: (copy_rows, None)}
            if horizontal_count:
                insertions[merges] = (None, write_horizontal_merges)
            serialize_xml(doc, output, insertions)
            return len(plans), filled
        finally:
            for plan in plans:
                if plan["anchor"] is not None:
                    plan["anchor"].unlink()
            if doc is not None:
                doc.unlink()
            builder.document.unlink()
            builder._parser = None


def worksheets(archive):
    """通过关系文件读取真正的工作表路径，包括中文名称和隐藏表。"""
    workbook = minidom.parseString(archive.read("xl/workbook.xml"))
    relationships = minidom.parseString(archive.read("xl/_rels/workbook.xml.rels"))
    try:
        paths = {}
        for relation in children(relationships.documentElement, "Relationship"):
            if (relation.getAttribute("Type").endswith("/worksheet")
                    and relation.getAttribute("TargetMode") != "External"):
                target = relation.getAttribute("Target")
                paths[relation.getAttribute("Id")] = (
                    target.lstrip("/") if target.startswith("/")
                    else posixpath.normpath(posixpath.join("xl", target)))
        result = []
        for group in children(workbook.documentElement, "sheets"):
            for sheet in children(group, "sheet"):
                rid = next((a.value for a in sheet.attributes.values()
                            if a.localName == "id" and a.namespaceURI), None)
                if rid in paths:
                    result.append((sheet.getAttribute("name"), paths[rid]))
        return result
    finally:
        workbook.unlink()
        relationships.unlink()


def process_file(source, all_merges=False, sheets=None):
    """返回 (新文件路径或 None, [(工作表名, 拆分区域数, 新填充格数)])。"""
    source = Path(source).expanduser().resolve()
    if source.suffix.lower() not in (".xlsx", ".xlsm"):
        raise ValueError("仅支持 .xlsx / .xlsm；旧版 .xls 请先在 Excel 中另存为 .xlsx。")
    if not source.is_file():
        raise ValueError("找不到文件：" + str(source))
    with ZipFile(source) as original, ExitStack() as temporary:
        if any(n.startswith("_xmlsignatures/") for n in original.namelist()):
            raise ValueError("文件含文档数字签名，请使用未签名副本处理。")
        available = worksheets(original)
        requested = set(sheets or [])
        unknown = requested - {name for name, _ in available}
        if unknown:
            raise ValueError("找不到工作表：" + "、".join(sorted(unknown)))
        updates, stats = {}, []
        for name, path in available:
            if requested and name not in requested:
                continue
            try:
                transformed = temporary.enter_context(TemporaryFile(mode="w+b"))
                if original.getinfo(path).file_size >= STREAM_THRESHOLD:
                    with original.open(path) as incoming:
                        count, filled = transform_sheet_stream(incoming, transformed, all_merges)
                else:
                    raw, count, filled = transform_sheet(original.read(path), all_merges)
                    if count:
                        transformed.write(raw)
                    del raw
            except ValueError as error:
                raise ValueError("工作表“" + name + "”：" + str(error)) from error
            stats.append((name, count, filled))
            if count:
                updates[path] = transformed
        if not updates:
            return None, stats
        # 独占创建，重复运行也不会覆盖原文件或上一次结果。
        serial = 1
        while True:
            suffix = "" if serial == 1 else "_" + str(serial)
            output = source.with_name(source.stem + "_拆分填充" + suffix + source.suffix)
            try:
                stream = output.open("xb")
                break
            except FileExistsError:
                serial += 1
            except PermissionError:
                # Windows 对同名目录报告 PermissionError；实际写入权限错误仍需抛出。
                if not output.is_dir():
                    raise
                serial += 1
        try:
            with stream, ZipFile(stream, "w") as result:
                result.comment = original.comment
                for info in original.infolist():
                    outgoing = copy.copy(info)
                    contents = updates.get(info.filename)
                    if contents is not None:
                        outgoing.file_size = contents.seek(0, 2)
                        contents.seek(0)
                        with result.open(outgoing, "w") as destination:
                            shutil.copyfileobj(contents, destination, 1024 * 1024)
                    else:
                        # 共享字符串、图片等未修改部件也不整体读入内存。
                        with original.open(info) as incoming, result.open(outgoing, "w") as destination:
                            shutil.copyfileobj(incoming, destination, 1024 * 1024)
        except BaseException:
            output.unlink(missing_ok=True)
            raise
    return output, stats


def choose_files():
    if sys.platform == "darwin":
        script = ('set chosen to choose file with prompt "选择要拆分合并行的 Excel 文件" '
                  'with multiple selections allowed\n'
                  'set paths to ""\nrepeat with itemPath in chosen\n'
                  'set paths to paths & POSIX path of itemPath & linefeed\n'
                  'end repeat\nreturn paths')
        result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        if result.returncode:
            if "-128" in result.stderr:
                return []
            raise RuntimeError("无法打开文件选择框，请通过命令行指定文件路径。")
        return result.stdout.strip().splitlines()
    try:
        import tkinter as tk
        from tkinter.filedialog import askopenfilenames
        root = tk.Tk()
        root.withdraw()
        try:
            return list(askopenfilenames(title="选择 Excel 文件", filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm")]))
        finally:
            root.destroy()
    except ImportError as error:
        raise RuntimeError("请通过命令行指定 Excel 文件路径。") from error


def main():
    parser = argparse.ArgumentParser(description="拆分 Excel 合并行并填充原值，另存为新文件。")
    parser.add_argument("files", nargs="*", help="Excel 文件路径；不填时打开文件选择框")
    parser.add_argument("--all-merges", action="store_true", help="也拆开横向合并，填满区域内每个单元格")
    parser.add_argument("--sheet", action="append", help="仅处理指定工作表，可重复；默认处理全部工作表")
    args = parser.parse_args()
    try:
        files = args.files or choose_files()
    except (OSError, RuntimeError) as error:
        print("无法选择文件：" + str(error), file=sys.stderr)
        return 1
    if not files:
        print("已取消，未处理文件。")
        return 0
    failures = 0
    for file in files:
        print("正在处理：" + str(file), flush=True)
        try:
            output, stats = process_file(file, args.all_merges, args.sheet)
            for name, count, filled in stats:
                print("  {}：拆分 {} 个合并区域，新填充 {} 个单元格".format(name, count, filled))
            print("已保存：" + str(output) if output else "没有符合条件的合并区域，未生成新文件。")
        except Exception as error:
            # 单个文件失败后继续处理其他文件；不保留不完整结果。
            failures += 1
            print("处理失败：" + str(error), file=sys.stderr)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
