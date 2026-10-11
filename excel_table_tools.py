"""本地数据型合表与拆表：有限内存、整批发布、源文件只读。"""

import json
import os
import re
import shutil
import sqlite3
import tempfile
import time
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path

from excel_table_io import (Cell, TableCancelled, TableReader, XlsxWriter,
                            is_date_format, list_sheets)

MAX_DATA_ROWS = 1_048_575
MAX_COLUMNS = 16_384
MAX_OUTPUT_FILES = 1_000


@dataclass(frozen=True)
class TableSource:
    path: Path
    sheet: str
    header_row: int = 1


@dataclass(frozen=True)
class TaskResult:
    directory: Path
    files: tuple[Path, ...]
    rows: int
    manifest: Path


def _check_cancel(cancel):
    if cancel is not None and cancel.is_set():
        raise TableCancelled("任务已取消，未生成正式结果。")


class _Progress:
    def __init__(self, callback, cancel):
        self.callback, self.cancel = callback, cancel
        self.last = 0
        self.phase = None

    def emit(self, phase, completed=None, sheet=None, force=False):
        _check_cancel(self.cancel)
        now = time.monotonic()
        if self.callback and (force or phase != self.phase or now - self.last >= .3):
            self.last, self.phase = now, phase
            self.callback(dict(phase=phase, completed=completed, total=None,
                               sheet=sheet, unit="rows"))


def inspect_workbook(path):
    return list_sheets(Path(path))


def _sources(sources):
    result, seen = [], set()
    for source in sources:
        path = Path(source.path).expanduser().resolve(strict=True)
        if path.suffix.lower() not in (".xlsx", ".xlsm"):
            raise ValueError("只支持 .xlsx / .xlsm：{}".format(path.name))
        if not isinstance(source.header_row, int) or isinstance(source.header_row, bool) or not 1 <= source.header_row <= 1_048_576:
            raise ValueError("表头行须为 1 至 1048576 的整数。")
        key = (os.path.normcase(str(path)), source.sheet, source.header_row)
        if key in seen:
            raise ValueError("重复输入同一文件、工作表和表头行：{} / {}".format(path.name, source.sheet))
        seen.add(key)
        result.append(TableSource(path, source.sheet, source.header_row))
    if not result:
        raise ValueError("请先添加至少一个工作表。")
    return result


def _signature(source):
    stat = source.path.stat()
    return dict(path=str(source.path), sheet=source.sheet, header_row=source.header_row,
                size=stat.st_size, mtime_ns=stat.st_mtime_ns)


def _validate_headers(headers, expected, source):
    if set(headers) != set(expected):
        missing = [h for h in expected if h not in headers]
        extra = [h for h in headers if h not in expected]
        raise ValueError("{} / {} 的表头不一致；缺少列：{}；多出列：{}。".format(
            source.path.name, source.sheet, "、".join(missing) or "无", "、".join(extra) or "无"))


def preview_sources(sources, progress=None, cancel=None):
    """只读取各表表头；完整数据校验在执行时进行，不估算虚假的行数。"""
    sources = _sources(sources)
    headers, epoch, snapshots = None, None, []
    reporter = _Progress(progress, cancel)
    for source in sources:
        reporter.emit("checking", sheet=source.sheet, force=True)
        before = _signature(source)
        with TableReader(source.path, source.sheet, source.header_row,
                         progress=progress, cancel=cancel) as reader:
            if headers is None:
                headers, epoch = list(reader.headers), reader.date1904
            else:
                _validate_headers(reader.headers, headers, source)
                if reader.date1904 != epoch:
                    raise ValueError("输入混用了 1900 / 1904 日期系统，请在 Excel 中统一日期系统后重试。")
        if before != _signature(source):
            raise ValueError("读取期间源文件发生变化，请关闭编辑并重试：{}".format(source.path.name))
        snapshots.append(before)
    return dict(headers=headers, date1904=epoch, sources=snapshots)


def _provenance_headers(headers):
    used, names = set(headers), []
    for base in ("来源文件", "来源工作表", "来源行号"):
        name, index = base, 2
        while name in used:
            name = "{}_{}".format(base, index)
            index += 1
        used.add(name)
        names.append(name)
    return names


def _group_key(cell):
    if cell is None or cell.value is None:
        return '["blank"]', dict(kind="blank", value=None), "空白"
    kind, value = cell.kind, cell.value
    display = str(value)
    if kind == "n":
        try:
            number = Decimal(value)
            if not number.is_finite():
                raise ValueError("分组列包含非有限数值。")
            sign, digits, exponent = number.as_tuple()
            digits = list(digits)
            while len(digits) > 1 and digits[-1] == 0:
                digits.pop()
                exponent += 1
            if not any(digits):
                sign, digits, exponent = 0, [0], 0
            normalized = [sign, digits, exponent]
        except InvalidOperation as error:
            raise ValueError("分组列包含无法识别的数值。") from error
        kind = "date_serial" if is_date_format(cell.number_format) else "n"
    else:
        normalized = value
        if kind == "b":
            normalized = "1" if value in ("1", "true", "True") else "0"
            display = "TRUE" if normalized == "1" else "FALSE"
        elif kind == "s" and value == "":
            display = "空文本"
    return json.dumps([kind, normalized], ensure_ascii=False, separators=(",", ":")), dict(kind=kind, value=value), display


def _safe_label(value):
    # 序号在最前，避免清洗同名、大小写同名和 Windows 保留设备名碰撞。
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" .")
    return value[:36].rstrip(" .") or "空白"


class _Parts:
    def __init__(self, directory, headers, epoch, limit, prefix, records, group_id=None,
                 progress=None, cancel=None):
        self.directory, self.headers, self.epoch = directory, headers, epoch
        self.limit, self.prefix, self.records = limit, prefix, records
        self.group_id, self.writer, self.count, self.part = group_id, None, 0, 0
        self.path = None
        self.progress, self.cancel = progress, cancel

    def _open(self):
        _check_cancel(self.cancel)
        if len(self.records) >= MAX_OUTPUT_FILES:
            raise ValueError("本次结果超过 {} 个文件，请增大每份行数或选择分类更少的列。".format(MAX_OUTPUT_FILES))
        self.part += 1
        self.path = self.directory / "{}_{:03d}.xlsx".format(self.prefix, self.part)
        self.writer = XlsxWriter(self.path, self.headers, date1904=self.epoch,
                                 progress=self.progress, cancel=self.cancel)
        self.writer.__enter__()
        self.count = 0

    def append(self, cells):
        if self.writer is None:
            self._open()
        elif self.count >= self.limit:
            self.close()
            self._open()
        self.writer.append(cells)
        self.count += 1

    def close(self):
        if self.writer is not None:
            writer, self.writer = self.writer, None
            writer.close()
            record = dict(name=self.path.name, rows=self.count, part=self.part)
            if self.group_id is not None:
                record["group_id"] = self.group_id
            self.records.append(record)

    def finish(self):
        if self.writer is None and self.part == 0:
            self._open()
        self.close()


def _publish(stage, output_dir, mode):
    """排他取得父目录，再一次 rename 发布整套结果；任何已有路径都不覆盖。"""
    base = ("合表" if mode == "merge" else "拆表") + time.strftime("_%Y%m%d_%H%M%S")
    index = 1
    while True:
        target = output_dir / (base if index == 1 else "{} ({})".format(base, index))
        try:
            target.mkdir()
            break
        except FileExistsError:
            index += 1
    destination = target / "结果"
    try:
        os.rename(stage, destination)
    except BaseException:
        # 只删除本次排他新建的空目录，不删除别人写入的内容。
        try:
            target.rmdir()
        except OSError:
            pass
        raise
    return destination


def _cleanup_stage(stage, output_dir, error):
    try:
        if stage.parent.resolve() != output_dir.resolve() or not stage.name.startswith(".excel-tools-"):
            raise RuntimeError("临时目录不在本次结果目录内。")
        if stage.is_symlink():
            stage.unlink()
        elif stage.exists():
            shutil.rmtree(stage)
    except OSError as cleanup_error:
        error.partial_path = str(stage)
        error.cleanup_error = str(cleanup_error)
        error.add_note("临时文件清理失败：{}；{}".format(stage, cleanup_error))


def run_tables(sources, output_dir, mode="merge", rows_per_file=100_000,
               split_column=None, provenance=True, progress=None, cancel=None):
    """从真实输入重新校验并执行；失败或取消不会返回部分成功。"""
    if mode not in ("merge", "split_rows", "split_column"):
        raise ValueError("未知处理模式。")
    sources = _sources(sources)
    if mode != "merge" and len(sources) != 1:
        raise ValueError("拆表每次只能选择一个工作表。")
    if mode == "split_rows" and (not isinstance(rows_per_file, int) or isinstance(rows_per_file, bool) or not 1 <= rows_per_file <= MAX_DATA_ROWS):
        raise ValueError("每份数据行数须为 1 至 {} 的整数（不含表头）。".format(MAX_DATA_ROWS))
    output_dir = Path(output_dir).expanduser().resolve(strict=True)
    if not output_dir.is_dir():
        raise ValueError("请选择已存在的结果目录。")
    reporter = _Progress(progress, cancel)
    reporter.emit("checking", force=True)
    plan = preview_sources(sources, progress=progress, cancel=cancel)
    headers = plan["headers"]
    if mode == "split_column" and split_column not in headers:
        raise ValueError("请选择表头中存在的拆分列。")
    extra_headers = _provenance_headers(headers) if provenance else []
    out_headers = headers + extra_headers
    if len(out_headers) > MAX_COLUMNS:
        raise ValueError("加上来源列后超过 Excel 的 16384 列上限，请关闭来源列后重试。")
    _check_cancel(cancel)
    stage = Path(tempfile.mkdtemp(prefix=".excel-tools-", suffix=".part", dir=output_dir))
    records, groups, parts, database, total_rows = [], [], None, None, 0
    try:
        if mode == "split_column":
            database = sqlite3.connect(stage / "rows.sqlite")
            database.execute("PRAGMA cache_size=-8192")
            database.execute("PRAGMA temp_store=FILE")
            database.execute("CREATE TABLE groups (id INTEGER PRIMARY KEY, key TEXT UNIQUE, detail TEXT, label TEXT)")
            database.execute("CREATE TABLE rows (seq INTEGER PRIMARY KEY, group_id INTEGER, data TEXT)")
            split_index = headers.index(split_column)
            cached_key, cached_id = None, None
        else:
            limit = min(rows_per_file, MAX_DATA_ROWS) if mode == "split_rows" else MAX_DATA_ROWS
            parts = _Parts(stage, out_headers, plan["date1904"], limit,
                           "合表" if mode == "merge" else "按行拆表", records,
                           progress=progress, cancel=cancel)
        for source, snapshot in zip(sources, plan["sources"]):
            _check_cancel(cancel)
            if snapshot != _signature(source):
                raise ValueError("源文件在检查后发生变化，请重试：{}".format(source.path.name))
            with TableReader(source.path, source.sheet, source.header_row,
                             progress=progress, cancel=cancel) as reader:
                _validate_headers(reader.headers, headers, source)
                if reader.date1904 != plan["date1904"]:
                    raise ValueError("源文件日期系统发生变化，请重试。")
                mapping = [reader.headers.index(header) for header in headers]
                for row_number, row in reader.iter_rows():
                    _check_cancel(cancel)
                    cells = [row[index] for index in mapping]
                    group_cell = cells[split_index] if mode == "split_column" else None
                    if provenance:
                        cells += [Cell("s", str(source.path)), Cell("s", source.sheet), Cell("n", str(row_number))]
                    if database is None:
                        parts.append(cells)
                    else:
                        key, detail, label = _group_key(group_cell)
                        if key == cached_key:
                            group_id = cached_id
                        else:
                            database.execute("INSERT OR IGNORE INTO groups(key,detail,label) VALUES (?,?,?)",
                                             (key, json.dumps(detail, ensure_ascii=False), _safe_label(label)))
                            group_id = database.execute("SELECT id FROM groups WHERE key=?", (key,)).fetchone()[0]
                            cached_key, cached_id = key, group_id
                        data = [[cell.kind, cell.value, cell.number_format] if cell is not None else None for cell in cells]
                        database.execute("INSERT INTO rows(group_id,data) VALUES (?,?)",
                                         (group_id, json.dumps(data, ensure_ascii=False, separators=(",", ":"))))
                        if total_rows % 10_000 == 0:
                            database.commit()
                    total_rows += 1
                    reporter.emit("grouping" if database is not None else "writing", total_rows, source.sheet)
            if snapshot != _signature(source):
                raise ValueError("处理期间源文件发生变化，请重试：{}".format(source.path.name))
        if database is not None:
            reporter.emit("indexing", total_rows, force=True)
            database.commit()
            database.set_progress_handler(lambda: 1 if cancel is not None and cancel.is_set() else 0, 10_000)
            file_count = database.execute(
                "SELECT COALESCE(SUM((n + ? - 1) / ?), 0) FROM "
                "(SELECT COUNT(*) AS n FROM rows GROUP BY group_id)",
                (MAX_DATA_ROWS, MAX_DATA_ROWS)).fetchone()[0]
            if file_count > MAX_OUTPUT_FILES:
                raise ValueError("按此列将生成 {} 个文件，超过单次 {} 个文件的限制；请选择分类更少的列。".format(file_count, MAX_OUTPUT_FILES))
            try:
                database.execute("CREATE INDEX group_rows ON rows(group_id,seq)")
            except sqlite3.OperationalError:
                _check_cancel(cancel)
                raise
            saved_rows = 0
            for group_id, detail, label in database.execute("SELECT id,detail,label FROM groups ORDER BY id"):
                _check_cancel(cancel)
                group = json.loads(detail)
                group["id"] = group_id
                groups.append(group)
                parts = _Parts(stage, out_headers, plan["date1904"], MAX_DATA_ROWS,
                               "分组_{:06d}_{}".format(group_id, label), records, group_id,
                               progress=progress, cancel=cancel)
                for (data,) in database.execute("SELECT data FROM rows WHERE group_id=? ORDER BY seq", (group_id,)):
                    _check_cancel(cancel)
                    parts.append([Cell(*cell) if cell is not None else None for cell in json.loads(data)])
                    saved_rows += 1
                    reporter.emit("saving", saved_rows)
                parts.finish()
            database.close()
            database = None
            (stage / "rows.sqlite").unlink()
            if not groups:
                parts = _Parts(stage, out_headers, plan["date1904"], MAX_DATA_ROWS, "空结果", records,
                               progress=progress, cancel=cancel)
        if parts is not None:
            parts.finish()
        manifest = dict(format_version=1, mode=mode, rows=total_rows, sources=plan["sources"],
                        files=records, groups=groups, provenance_columns=extra_headers,
                        headers=out_headers, date1904=plan["date1904"],
                        contract="仅保留值、类型和数字格式；公式与参与区域合并单元格拒绝。")
        (stage / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
        reporter.emit("publishing", total_rows, force=True)
        _check_cancel(cancel)
        directory = _publish(stage, output_dir, mode)
        return TaskResult(directory, tuple(directory / item["name"] for item in records),
                          total_rows, directory / "manifest.json")
    except BaseException as error:
        if parts is not None and parts.writer is not None:
            try:
                parts.writer.__exit__(type(error), error, error.__traceback__)
                parts.writer = None
            except BaseException:
                pass
        if database is not None:
            try:
                database.close()
            except Exception:
                pass
        _cleanup_stage(stage, output_dir, error)
        if isinstance(error, sqlite3.OperationalError) and getattr(error, "sqlite_errorcode", None) == sqlite3.SQLITE_INTERRUPT:
            _check_cancel(cancel)
        raise
