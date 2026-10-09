#!/usr/bin/env python3
"""拆分 Excel 合并行并填充原值。Python 3.8+，只使用标准库。"""

import argparse
import copy
import posixpath
import re
import subprocess
import sys
from pathlib import Path
from xml.dom import Node, minidom
from zipfile import ZipFile


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


def new_element(doc, name):
    root = doc.documentElement
    qualified = root.prefix + ":" + name if root.prefix else name
    return doc.createElementNS(root.namespaceURI, qualified)


def payload(cell):
    if cell is None:
        return False
    if children(cell, "f"):
        return True
    # 空的 <v/> / <is><t/></is> 是空白；"0"、" " 都是原有内容。
    values = children(cell, "v") + list(cell.getElementsByTagNameNS(cell.namespaceURI, "t"))
    return any(n.nodeValue for value in values for n in value.childNodes
               if n.nodeType in (Node.TEXT_NODE, Node.CDATA_SECTION_NODE))


def transform_sheet(raw, all_merges=False):
    """返回 (XML bytes, 拆分区域数, 新填充单元格数)。不修改普通空白。"""
    doc = minidom.parseString(raw)
    try:
        root = doc.documentElement
        merge_groups = children(root, "mergeCells")
        if not merge_groups:
            return raw, 0, 0
        merges = merge_groups[0]
        selected = []
        for merge in children(merges, "mergeCell"):
            box = bounds(merge.getAttribute("ref"))
            if box[2] > box[0] or (all_merges and box[3] > box[1]):
                selected.append((merge, box))
        if not selected:
            return raw, 0, 0
        data = children(root, "sheetData")[0]
        row_nodes = {int(row.getAttribute("r")): row for row in children(data, "row")}
        cells = {cell.getAttribute("r"): cell for row in row_nodes.values()
                 for cell in children(row, "c")}
        changed_rows = set()
        filled = 0
        for merge, (r1, c1, r2, c2) in selected:
            anchor_ref = address(r1, c1)
            anchor = cells.get(anchor_ref)
            if anchor is not None and children(anchor, "f"):
                raise ValueError("合并区域 " + merge.getAttribute("ref")
                                 + " 含公式，请先在 Excel 中复制并粘贴为值后重试。")
            # 先检查被合并掩盖的内容，避免静默覆盖原有数据。
            for row in range(r1, r2 + 1):
                for col in range(c1, c2 + 1):
                    ref = address(row, col)
                    cell = cells.get(ref)
                    if ref != anchor_ref and payload(cell):
                        raise ValueError("合并区域内的 " + ref
                                         + " 仍有独立内容，请先核对后再处理。")
            for row in range(r1, r2 + 1):
                if row not in row_nodes:
                    row_node = new_element(doc, "row")
                    row_node.setAttribute("r", str(row))
                    row_nodes[row] = row_node
                row_node = row_nodes[row]
                changed_rows.add(row)
                # 默认只拆纵向；矩形区域每一行仍保留横向合并。
                columns = range(c1, c2 + 1) if all_merges else (c1,)
                for col in columns:
                    ref = address(row, col)
                    if ref == anchor_ref:
                        continue
                    old = cells.get(ref)
                    cell = new_element(doc, "c")
                    cell.setAttribute("r", ref)
                    if anchor is not None:
                        for attr in ("s", "t"):
                            if anchor.hasAttribute(attr):
                                cell.setAttribute(attr, anchor.getAttribute(attr))
                        for child in anchor.childNodes:
                            if (child.nodeType == Node.ELEMENT_NODE
                                    and child.localName in ("v", "is")):
                                cell.appendChild(child.cloneNode(True))
                    if old is not None:
                        row_node.replaceChild(cell, old)
                    else:
                        row_node.appendChild(cell)
                    cells[ref] = cell
                    if payload(anchor):
                        filled += 1
                if not all_merges and c2 > c1:
                    horizontal = new_element(doc, "mergeCell")
                    horizontal.setAttribute("ref", address(row, c1) + ":" + address(row, c2))
                    merges.appendChild(horizontal)
            merges.removeChild(merge)

        # OOXML 要求行、单元格按坐标递增排列；新增行保留其他行属性。
        for row in sorted(changed_rows):
            row_node = row_nodes[row]
            if row_node.hasAttribute("spans"):
                row_node.removeAttribute("spans")
            tail = next((n for n in row_node.childNodes
                         if n.nodeType == Node.ELEMENT_NODE and n.localName != "c"), None)
            for cell in sorted(children(row_node, "c"),
                               key=lambda c: coordinate(c.getAttribute("r"))[1]):
                row_node.insertBefore(cell, tail)
        for row in sorted(row_nodes):
            data.appendChild(row_nodes[row])
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
            dimensions[0].setAttribute("ref", address(r1, c1) + ":" + address(r2, c2))
        return doc.toxml(encoding="utf-8"), len(selected), filled
    finally:
        doc.unlink()


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
    with ZipFile(source) as original:
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
                raw, count, filled = transform_sheet(original.read(path), all_merges)
            except ValueError as error:
                raise ValueError("工作表“" + name + "”：" + str(error)) from error
            stats.append((name, count, filled))
            if count:
                updates[path] = raw
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
        try:
            with stream, ZipFile(stream, "w") as result:
                result.comment = original.comment
                for info in original.infolist():
                    contents = updates.get(info.filename)
                    if contents is None:
                        contents = original.read(info)
                    result.writestr(copy.copy(info), contents)
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
