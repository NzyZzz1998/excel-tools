"""Bounded-memory scalar XLSX I/O for table operations (Python 3.10+).

This module creates new data workbooks. It retains cell types and number formats,
not workbook layout, drawings or formulas. Callers publish completed files.
"""
from __future__ import annotations

from collections import OrderedDict
from contextlib import ExitStack
from dataclasses import dataclass
from datetime import date, datetime, time as datetime_time
from decimal import Decimal, InvalidOperation
from pathlib import Path
import posixpath
import re
import sqlite3
import sys
from tempfile import TemporaryDirectory, TemporaryFile
import time
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from excel_unmerge_fill import _cleanup_part, address, bounds, coordinate


NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
MAX_DATA_ROWS = 1_048_575
MAX_COLUMNS = 16_384
_ESCAPE = re.compile(r"_x([0-9a-fA-F]{4})_")
_NUMBER = re.compile(r"[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[Ee][+-]?[0-9]+)?\Z")
_ISO_DATE = re.compile(r"[0-9]{4}-[0-9]{2}-[0-9]{2}(?:T[0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?(?:Z|[+-][0-9]{2}:[0-9]{2})?)?\Z")
_ISO_TIME = re.compile(r"[0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?(?:Z|[+-][0-9]{2}:[0-9]{2})?\Z")
_DATE_IDS = frozenset((*range(14, 23), *range(27, 37), *range(45, 48), *range(50, 59)))


class TableCancelled(Exception):
    """The caller's cancellation event was set; no result should be published."""


@dataclass(frozen=True, slots=True)
class Cell:
    kind: str
    value: str | None
    number_format: int | str = 0


def is_date_format(number_format: int | str) -> bool:
    """Recognize date/time tokens outside quoted, escaped and decorative text."""
    if isinstance(number_format, int):
        return number_format in _DATE_IDS
    index = 0
    quoted = False
    while index < len(number_format):
        char = number_format[index]
        if char == '"':
            quoted = not quoted
        elif char == "\\":
            index += 1
        elif not quoted:
            if char in "_*":
                index += 1
            elif char == "[":
                end = number_format.find("]", index + 1)
                if end == -1:
                    return False
                token = number_format[index + 1:end].lower()
                if token and len(set(token)) == 1 and token[0] in "hms":
                    return True
                index = end
            elif char.lower() in "ymdhs":
                return True
        index += 1
    return False


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _child(element, name):
    return next((node for node in element if _local(node.tag) == name), None)


def _decode_text(value):
    # One pass is essential: _x005F_x0041_ denotes the literal text _x0041_.
    value = _ESCAPE.sub(lambda match: chr(int(match[1], 16)), value)
    return value.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")


def _encode_text(value):
    value = _ESCAPE.sub(lambda match: "_x005F_" + match[0][1:], value)
    return "".join(
        f"_x{ord(char):04X}_" if (ord(char) < 32 and char not in "\t\n\r")
        or 0xD800 <= ord(char) <= 0xDFFF or ord(char) in (0xFFFE, 0xFFFF)
        else char for char in value
    )


def _xml(value, attribute=False):
    value = value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\r", "&#13;")
    if attribute:
        value = value.replace('"', "&quot;").replace("\n", "&#10;").replace("\t", "&#9;")
    return value


def _string_content(element):
    if element is None:
        return ""
    pieces = []
    for node in element:
        if _local(node.tag) == "t":
            pieces.append(node.text or "")
        elif _local(node.tag) == "r":
            pieces.extend(child.text or "" for child in node if _local(child.tag) == "t")
    # Phonetic annotations (rPh/t) are not part of the cell's value.
    return _decode_text("".join(pieces))


def _validated_value(kind, value):
    if value is None:
        return None
    if not isinstance(value, str):
        raise ValueError("单元格原值必须为字符串或空值。")
    if kind == "n":
        try:
            valid = (_NUMBER.fullmatch(value.strip()) is not None
                     and Decimal(value).copy_abs() <= Decimal("1.7976931348623157E308"))
        except InvalidOperation:
            valid = False
        if not valid:
            raise ValueError("数值单元格内容无效。")
    elif kind == "b":
        if value not in ("0", "1", "false", "true"):
            raise ValueError("布尔单元格内容无效。")
        value = {"false": "0", "true": "1"}.get(value, value)
    elif kind == "d":
        try:
            if _ISO_DATE.fullmatch(value):
                (datetime if "T" in value else date).fromisoformat(value.replace("Z", "+00:00"))
            elif _ISO_TIME.fullmatch(value):
                datetime_time.fromisoformat(value.replace("Z", "+00:00"))
            else:
                raise ValueError()
        except ValueError:
            raise ValueError("日期单元格内容无效。") from None
    return value


def _number_formats(archive, part, observer):
    custom, bases, xfs, parents = {}, [], [], []
    with archive.open(part) as source:
        for event, element in ET.iterparse(source, events=("start", "end")):
            if event == "start":
                parents.append(element)
                continue
            tag = _local(element.tag)
            parent_tag = _local(parents[-2].tag) if len(parents) > 1 else None
            if tag == "numFmt" and parent_tag == "numFmts":
                custom[int(element.get("numFmtId"))] = element.get("formatCode", "")
            elif tag == "xf" and parent_tag in ("cellStyleXfs", "cellXfs"):
                observer.check()
                record = (element.get("numFmtId"), int(element.get("xfId", "0")), element.get("applyNumberFormat"))
                (bases if parent_tag == "cellStyleXfs" else xfs).append(record)
            if len(parents) > 1:
                parents[-2].remove(element)
            element.clear()
            parents.pop()
    result = []
    for number, base, apply in xfs:
        if (number is None or apply in ("0", "false")) and 0 <= base < len(bases):
            number = bases[base][0]
        number = int(number or "0")
        if number < 0 or (number >= 164 and number not in custom):
            raise ValueError("数字格式定义无效或缺失。")
        result.append(custom.get(number, number))
    return result or [0]


class _Observer:
    def __init__(self, progress, cancel, sheet):
        self.callback, self.cancel, self.sheet = progress, cancel, sheet
        self.last_time = None
        self.last_phase = None

    def check(self):
        if self.cancel is not None and self.cancel.is_set():
            raise TableCancelled("已取消当前表格任务。")

    def emit(self, phase, completed=None, total=None, unit=None, force=False):
        self.check()
        if self.callback is None:
            return
        now = time.monotonic()
        if not force and phase == self.last_phase and self.last_time is not None and now - self.last_time < 0.3:
            return
        self.last_phase, self.last_time = phase, now
        try:
            self.callback(dict(phase=phase, sheet=self.sheet, completed=completed, total=total, unit=unit))
        except Exception:
            pass
        self.check()


def _parse_part(archive, path):
    with archive.open(path) as source:
        return ET.parse(source).getroot()


def _target(base, target):
    result = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(base), target))
    if result.startswith("../") or result == "..":
        raise ValueError("工作簿关联路径无效。")
    return result


def _metadata(archive):
    workbook_path = "xl/workbook.xml"
    if "_rels/.rels" in archive.namelist():
        package = _parse_part(archive, "_rels/.rels")
        for item in package:
            if item.get("Type", "").endswith("/officeDocument") and item.get("TargetMode") != "External":
                workbook_path = _target("", item.get("Target", ""))
                break
    workbook = _parse_part(archive, workbook_path)
    folder, filename = posixpath.split(workbook_path)
    relationships = _parse_part(archive, posixpath.join(folder, "_rels", filename + ".rels"))
    paths, auxiliary = {}, {}
    for item in relationships:
        if item.get("TargetMode") == "External":
            continue
        kind = item.get("Type", "").rsplit("/", 1)[-1]
        target = _target(workbook_path, item.get("Target", ""))
        if kind == "worksheet":
            paths[item.get("Id")] = target
        elif kind in ("styles", "sharedStrings"):
            auxiliary[kind] = target
    sheets = {}
    for group in workbook:
        if _local(group.tag) != "sheets":
            continue
        for sheet in group:
            rid = next((value for name, value in sheet.attrib.items() if _local(name) == "id"), None)
            if rid in paths:
                name = sheet.get("name", "")
                if name in sheets:
                    raise ValueError("工作簿含重复工作表名称。")
                sheets[name] = paths[rid]
    properties = _child(workbook, "workbookPr")
    date1904 = properties is not None and properties.get("date1904", "0").lower() in ("1", "true")
    return sheets, auxiliary, date1904


def list_sheets(path):
    with ZipFile(path) as archive:
        return list(_metadata(archive)[0])


class _SharedStrings:
    """Disk index plus a byte- and entry-bounded cache; never a full text list."""
    def __init__(self, stack, archive, part, observer):
        directory = stack.enter_context(TemporaryDirectory(prefix="excel-table-strings-"))
        self.connection = sqlite3.connect(str(Path(directory) / "strings.sqlite"))
        stack.callback(self.connection.close)
        self.connection.execute("PRAGMA cache_size=-4096")
        self.connection.execute("PRAGMA journal_mode=OFF")
        self.connection.execute("PRAGMA synchronous=OFF")
        self.connection.execute("PRAGMA temp_store=FILE")
        self.connection.execute("CREATE TABLE strings (id INTEGER PRIMARY KEY, value BLOB NOT NULL)")
        self.cache = OrderedDict()
        self.cache_bytes = 0
        self.count = 0
        with archive.open(part) as source:
            parents = []
            for event, element in ET.iterparse(source, events=("start", "end")):
                if event == "start":
                    parents.append(element)
                    continue
                if _local(element.tag) == "si":
                    observer.check()
                    value = _string_content(element).encode("utf-8", "surrogatepass")
                    self.connection.execute("INSERT INTO strings VALUES (?, ?)", (self.count, value))
                    self.count += 1
                    parents[-2].remove(element)
                    element.clear()
                    observer.emit("reading")
                elif len(parents) > 1 and not any(_local(parent.tag) == "si" for parent in parents[:-1]):
                    parents[-2].remove(element)
                    element.clear()
                parents.pop()
        self.connection.commit()

    def get(self, index):
        if index in self.cache:
            self.cache.move_to_end(index)
            return self.cache[index]
        row = self.connection.execute("SELECT value FROM strings WHERE id=?", (index,)).fetchone()
        if row is None:
            raise ValueError("共享字符串索引不存在：" + str(index))
        value = row[0].decode("utf-8", "surrogatepass")
        size = sys.getsizeof(value)
        if size <= 4 * 1024 * 1024:
            self.cache[index] = value
            self.cache_bytes += size
            while len(self.cache) > 1024 or self.cache_bytes > 4 * 1024 * 1024:
                _, expired = self.cache.popitem(last=False)
                self.cache_bytes -= sys.getsizeof(expired)
        return value


class TableReader:
    def __init__(self, path, sheet: str, header_row: int = 1, progress=None, cancel=None):
        self.path = Path(path)
        self.sheet = sheet
        if not isinstance(header_row, int) or not 1 <= header_row <= MAX_DATA_ROWS + 1:
            raise ValueError("表头行必须在 1 至 1048576 之间。")
        self.header_row = header_row
        self.headers = []
        self.date1904 = False
        self.observer = _Observer(progress, cancel, sheet)
        self._stack = None
        self._rows = None
        self._consumed = False

    def _error(self, message, coordinate=None):
        location = self.path.name + " / " + self.sheet
        if coordinate is not None:
            location += " / " + str(coordinate)
        return ValueError(location + "：" + message)

    def __enter__(self):
        if self._stack is not None:
            raise RuntimeError("TableReader 已打开。")
        self._stack = ExitStack()
        try:
            self.observer.emit("reading", force=True)
            self.archive = self._stack.enter_context(ZipFile(self.path))
            sheets, auxiliary, self.date1904 = _metadata(self.archive)
            if self.sheet not in sheets:
                raise self._error("工作表不存在。")
            self._sheet_part = sheets[self.sheet]
            self._formats = [0]
            if "styles" in auxiliary:
                try:
                    self._formats = _number_formats(self.archive, auxiliary["styles"], self.observer)
                except (ValueError, ET.ParseError) as error:
                    raise self._error("数字格式定义无效或缺失。") from error
            self._strings = (_SharedStrings(self._stack, self.archive, auxiliary["sharedStrings"], self.observer)
                             if "sharedStrings" in auxiliary else None)
            self._rows = self._raw_rows()
            self._stack.callback(self._rows.close)
            for number, cells in self._rows:
                if number < self.header_row:
                    continue
                if number != self.header_row:
                    raise self._error("指定表头行为空。", self.header_row)
                values = ["" if cell is None or cell.value is None else cell.value.strip() for cell in cells]
                while values and not values[-1]:
                    values.pop()
                if not values:
                    raise self._error("表头不能为空。", self.header_row)
                if any(not value for value in values):
                    raise self._error("表头含空列名。", self.header_row)
                if len(values) != len(set(values)):
                    raise self._error("表头去除首尾空白后存在重复列名。", self.header_row)
                self.headers = values
                break
            else:
                raise self._error("找不到指定表头行。", self.header_row)
            return self
        except BaseException:
            self.close()
            raise

    def _read_cell(self, element, row, col, inherited_style):
        ref = address(row, col)
        if _child(element, "f") is not None:
            raise self._error("含公式，请先在副本中复制并粘贴为值。", ref)
        if any(name in element.attrib for name in ("vm", "cm")) or _child(element, "extLst") is not None:
            raise self._error("含单元格图片或扩展元数据，不能作为普通表格值处理。", ref)
        try:
            style = int(element.get("s", inherited_style))
            if style < 0:
                raise ValueError()
            number_format = self._formats[style]
        except (ValueError, IndexError):
            raise self._error("单元格数字格式索引无效。", ref) from None
        kind = element.get("t", "n")
        value_node = _child(element, "v")
        value = value_node.text if value_node is not None else None
        if kind == "inlineStr":
            return Cell("s", _string_content(_child(element, "is")), number_format)
        if kind == "s":
            if value is None:
                return Cell("s", None, number_format)
            try:
                if self._strings is None or int(value) < 0:
                    raise ValueError()
                value = self._strings.get(int(value))
            except ValueError:
                raise self._error("共享字符串索引无效或缺失。", ref) from None
            kind = "s"
        elif kind == "str":
            kind, value = "s", _decode_text(value or "")
        elif kind not in ("n", "b", "e", "d"):
            raise self._error("不支持的单元格值类型。", ref)
        try:
            value = _validated_value(kind, value)
        except ValueError as error:
            raise self._error(str(error), ref) from None
        return Cell(kind, value, number_format)

    def _raw_rows(self):
        try:
            yield from self._parse_rows()
        except ET.ParseError as error:
            raise self._error("工作表 XML 无效。") from error

    def _parse_rows(self):
        last_row = 0
        physical = 0
        columns = []
        merged = None
        row_depth = None
        with self.archive.open(self._sheet_part) as source:
            parents = []
            for event, element in ET.iterparse(source, events=("start", "end")):
                if event == "start":
                    parents.append(element)
                    if (_local(element.tag) == "row" and len(parents) > 1
                            and _local(parents[-2].tag) == "sheetData"):
                        row_depth = len(parents)
                    continue
                tag = _local(element.tag)
                parent_tag = _local(parents[-2].tag) if len(parents) > 1 else None
                if tag == "col" and parent_tag == "cols":
                    if "style" in element.attrib:
                        columns.append((int(element.get("min")), int(element.get("max")), element.get("style")))
                    parents[-2].remove(element)
                elif tag == "mergeCell" and parent_tag == "mergeCells":
                    ref = element.get("ref", "")
                    try:
                        end_row = bounds(ref)[2]
                    except ValueError:
                        raise self._error("合并区域地址无效。", ref) from None
                    if end_row >= self.header_row:
                        merged = merged or ref
                    parents[-2].remove(element)
                elif tag == "row" and parent_tag == "sheetData":
                    self.observer.check()
                    try:
                        number = int(element.get("r", str(last_row + 1)))
                    except ValueError:
                        raise self._error("行号无效。") from None
                    if not last_row < number <= MAX_DATA_ROWS + 1:
                        raise self._error("行号重复、乱序或超出 Excel 范围。", number)
                    last_row = number
                    cells, seen, next_column = {}, set(), 1
                    row_style = element.get("s") if element.get("customFormat", "0").lower() in ("1", "true") else None
                    if number >= self.header_row:
                        for node in element:
                            if _local(node.tag) != "c":
                                continue
                            ref = node.get("r")
                            try:
                                cell_row, col = coordinate(ref) if ref else (number, next_column)
                            except ValueError:
                                raise self._error("单元格坐标无效。", ref) from None
                            if cell_row != number or col in seen or col > MAX_COLUMNS:
                                raise self._error("单元格坐标重复、超范围或与所在行不符。", ref or number)
                            seen.add(col)
                            next_column = col + 1
                            style = row_style
                            if style is None:
                                style = next((s for start, end, s in reversed(columns) if start <= col <= end), "0")
                            cells[col] = self._read_cell(node, number, col, style)
                    width = max(cells, default=0)
                    result = [cells.get(col) for col in range(1, width + 1)]
                    parents[-2].remove(element)
                    element.clear()
                    row_depth = None
                    physical += 1
                    self.observer.emit("reading", physical, None, "rows")
                    yield number, result
                elif row_depth is None and len(parents) > 1:
                    # Layout-only metadata is not copied to new data workbooks.
                    # Clear completed nodes too, including large tail collections.
                    parents[-2].remove(element)
                    element.clear()
                parents.pop()
        if merged is not None:
            raise self._error("表头或数据区包含合并单元格，请先使用拆分填充功能。", merged)
        self.observer.emit("reading", physical, physical, "rows", force=True)

    def iter_rows(self):
        if self._stack is None or self._consumed:
            raise RuntimeError("TableReader 必须处于打开状态，且数据只能迭代一次。")
        self._consumed = True
        width = len(self.headers)
        for number, cells in self._rows:
            if any(cell is not None and cell.value is not None for cell in cells[width:]):
                raise self._error("数据超出表头列数。", address(number, width + 1))
            cells = cells[:width] + [None] * max(0, width - len(cells))
            if any(cell is not None and cell.value is not None for cell in cells):
                yield number, cells

    def close(self):
        if self._stack is not None:
            stack, self._stack = self._stack, None
            stack.close()

    def __exit__(self, exc_type, exc_value, traceback):
        self.close()


class XlsxWriter:
    """Append-only single-sheet writer. A caller publishes the completed path."""
    def __init__(self, path, headers, date1904=False, *, progress=None, cancel=None):
        self.path = Path(path)
        self.headers = list(headers)
        if not self.headers or len(self.headers) > MAX_COLUMNS:
            raise ValueError("输出表头列数必须在 1 至 16384 之间。")
        self.date1904 = bool(date1904)
        self.data_rows = 0
        self._closed = False
        self._created = False
        self._rows = TemporaryFile(mode="w+b")
        self._formats = {0: 0}
        self.observer = _Observer(progress, cancel, "数据")
        try:
            self._write_row(1, [Cell("s", str(value)) for value in self.headers])
        except BaseException:
            self._rows.close()
            raise

    def __enter__(self):
        if self._closed:
            raise RuntimeError("XlsxWriter 已关闭。")
        return self

    def _write_row(self, number, cells):
        parts = ['<row r="' + str(number) + '">']
        for column, cell in enumerate(cells, 1):
            if cell is None or cell.value is None:
                continue
            if cell.kind not in ("n", "s", "b", "e", "d") or not isinstance(cell.value, str):
                raise ValueError("输出单元格类型或值无效：" + address(number, column))
            try:
                value = _validated_value(cell.kind, cell.value)
            except ValueError as error:
                raise ValueError(str(error) + "：" + address(number, column)) from None
            style = self._formats.setdefault(cell.number_format, len(self._formats))
            attributes = ' r="' + address(number, column) + '" s="' + str(style) + '"'
            if cell.kind == "s":
                if len(cell.value.encode("utf-16-le", "surrogatepass")) // 2 > 32767:
                    raise ValueError("输出文本超出 Excel 单元格长度限制：" + address(number, column))
                parts.append('<c' + attributes + ' t="inlineStr"><is><t xml:space="preserve">'
                             + _xml(_encode_text(cell.value)) + '</t></is></c>')
            else:
                parts.append('<c' + attributes + ' t="' + cell.kind + '"><v>'
                             + _xml(value) + '</v></c>')
        parts.append('</row>')
        self._rows.write("".join(parts).encode("utf-8"))

    def append(self, cells):
        if self._closed:
            raise RuntimeError("XlsxWriter 已关闭。")
        self.observer.check()
        if self.data_rows >= MAX_DATA_ROWS:
            raise ValueError("单个工作表最多容纳 1048575 行数据及一行表头，请拆分输出。")
        cells = list(cells)
        if len(cells) > len(self.headers):
            raise ValueError("输出数据列数超过表头。")
        self._write_row(self.data_rows + 2, cells)
        self.data_rows += 1
        self.observer.emit("filling", self.data_rows, None, "rows")

    def _styles_xml(self):
        custom, xfs = [], []
        for value in self._formats:
            if isinstance(value, str):
                number = 164 + len(custom)
                custom.append('<numFmt numFmtId="' + str(number) + '" formatCode="' + _xml(value, True) + '"/>')
            elif isinstance(value, int) and 0 <= value < 164:
                number = value
            else:
                raise ValueError("输出数字格式无效。")
            xfs.append('<xf numFmtId="' + str(number) + '" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
        return ('<styleSheet xmlns="' + NS + '"><numFmts count="' + str(len(custom)) + '">' + ''.join(custom)
                + '</numFmts><fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts>'
                + '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
                + '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                + '<cellXfs count="' + str(len(xfs)) + '">' + ''.join(xfs) + '</cellXfs>'
                + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>')

    def close(self):
        if self._closed:
            return
        try:
            self.observer.check()
            styles = self._styles_xml()
            total = self._rows.tell()
            self.observer.emit("saving", 0, total, "bytes", force=True)
            with self.path.open("xb") as destination:
                self._created = True
                with ZipFile(destination, "w", ZIP_DEFLATED) as archive:
                    archive.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                        '<Default Extension="xml" ContentType="application/xml"/>'
                        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
                    archive.writestr('_rels/.rels', '<Relationships xmlns="' + PKG_REL + '"><Relationship Id="rId1" Type="' + REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>')
                    archive.writestr('xl/workbook.xml', '<workbook xmlns="' + NS + '" xmlns:r="' + REL + '"><workbookPr date1904="' + str(int(self.date1904)) + '"/><sheets><sheet name="数据" sheetId="1" r:id="rId1"/></sheets></workbook>')
                    archive.writestr('xl/_rels/workbook.xml.rels', '<Relationships xmlns="' + PKG_REL + '"><Relationship Id="rId1" Type="' + REL + '/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="' + REL + '/styles" Target="styles.xml"/></Relationships>')
                    archive.writestr('xl/styles.xml', styles)
                    with archive.open('xl/worksheets/sheet1.xml', 'w', force_zip64=True) as stream:
                        stream.write(('<worksheet xmlns="' + NS + '"><dimension ref="A1:' + address(self.data_rows + 1, len(self.headers)) + '"/><sheetData>').encode('utf-8'))
                        self._rows.seek(0)
                        completed = 0
                        while chunk := self._rows.read(1024 * 1024):
                            self.observer.check()
                            stream.write(chunk)
                            completed += len(chunk)
                            self.observer.emit("saving", completed, total, "bytes")
                        stream.write(b'</sheetData></worksheet>')
            self.observer.emit("saving", total, total, "bytes", force=True)
        except BaseException as error:
            if self._created:
                _cleanup_part(self.path, error)
            raise
        finally:
            self._closed = True
            self._rows.close()

    def __exit__(self, exc_type, exc_value, traceback):
        if exc_type is None:
            self.close()
        else:
            self._closed = True
            self._rows.close()
