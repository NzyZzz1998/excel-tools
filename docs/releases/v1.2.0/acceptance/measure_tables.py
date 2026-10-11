"""Local acceptance helper; no business values are printed or saved to reports.

Run with repository on PYTHONPATH. Execution needs psutil for measurement;
--verify needs test-only openpyxl. Inputs/outputs remain in ignored testfile/.
"""
import argparse
from datetime import date, datetime, time as datetime_time, timedelta
import hashlib
import json
from pathlib import Path
import re
import threading
import time
import xml.etree.ElementTree as ET
import zipfile


def decode_inline_xstring(value):
    """Independent, single-pass ST_Xstring interpretation for this oracle.

    In particular, an escaped underscore must not trigger recursive decoding.
    No application reader/decoder is imported by verification.
    """
    decoded = re.sub(r"_x([0-9A-Fa-f]{4})_", lambda match: chr(int(match[1], 16)), value)
    return decoded.encode("utf-16-le", "surrogatepass").decode("utf-16-le", "surrogatepass")


def require_inline_strings(path):
    """Fail closed before normalizing openpyxl text: SST is already transformed.

    Scan all worksheet cell storage types, without retaining cells or rows.
    This deliberately limits the corrected monthly oracle to inlineStr inputs.
    """
    counts = {}
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.startswith("xl/worksheets/") or not name.endswith(".xml"):
                continue
            with archive.open(name) as stream:
                sheet_data = None
                for event, element in ET.iterparse(stream, events=("start", "end")):
                    local = element.tag.rsplit("}", 1)[-1]
                    if event == "start":
                        if local == "sheetData":
                            sheet_data = element
                        continue
                    if local == "c":
                        storage = element.get("t", "n")
                        counts[storage] = counts.get(storage, 0) + 1
                        if storage in ("s", "str"):
                            raise ValueError("Inline-only oracle cannot normalize shared/cached string storage")
                    elif local == "row" and sheet_data is not None:
                        sheet_data.remove(element)
                    element.clear()
    return counts


def sha(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(4 * 1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def verify(paths, inline_xstring=False):
    import openpyxl
    digest, count, unordered = hashlib.sha256(), 0, 0
    for path in paths:
        if inline_xstring:
            require_inline_strings(path)
        book = openpyxl.load_workbook(path, read_only=True, data_only=False)
        try:
            table = book.worksheets[0]
            table.reset_dimensions()
            for row_number, row in enumerate(table.iter_rows(), 1):
                if row_number == 1:
                    continue
                values = []
                for cell in row:
                    value = cell.value
                    if inline_xstring and cell.data_type == "s" and isinstance(value, str):
                        value = decode_inline_xstring(value)
                    if isinstance(value, (datetime, date, datetime_time)):
                        value = value.isoformat()
                    elif isinstance(value, timedelta):
                        value = value.total_seconds()
                    values.append([cell.data_type, value, cell.number_format] if value is not None else None)
                while values and values[-1] is None:
                    values.pop()
                if not values:
                    continue
                encoded = json.dumps(values, ensure_ascii=True, separators=(",", ":")).encode()
                digest.update(encoded)
                digest.update(b"\n")
                unordered = (unordered + int.from_bytes(hashlib.sha256(encoded).digest(), "big")) % (1 << 256)
                count += 1
                if count % 200_000 == 0:
                    print(json.dumps(dict(phase="independent_oracle", checked_rows=count)), flush=True)
        finally:
            book.close()
    return dict(rows=count, ordered_value_type_format_sha256=digest.hexdigest(),
                unordered_row_sha256_sum="{:064x}".format(unordered))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path)
    parser.add_argument("--input-report", type=Path)
    parser.add_argument("--merge-report", type=Path)
    parser.add_argument("--group-report", type=Path)
    parser.add_argument("--column-index", type=int, default=0)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--mode", choices=("merge", "split_rows", "split_column"), default="split_rows")
    parser.add_argument("--verify", action="store_true")
    parser.add_argument("--inline-xstring", action="store_true",
                        help="Decode inlineStr ST_Xstring once, rejecting SST/cached-string cells")
    args = parser.parse_args()
    if args.verify:
        prior = json.loads(args.input_report.read_text(encoding="utf-8"))
        started = time.perf_counter()
        result = dict(source=verify([args.source], args.inline_xstring),
                      output=verify(prior["output_paths"], args.inline_xstring))
        result["equal"] = result["source"] == result["output"]
        if args.merge_report:
            merged = json.loads(args.merge_report.read_text(encoding="utf-8"))
            result["merged"] = verify(merged["output_paths"], args.inline_xstring)
            result["equal"] = result["equal"] and result["source"] == result["merged"]
        if args.group_report:
            grouped = json.loads(args.group_report.read_text(encoding="utf-8"))
            result["grouped"] = verify(grouped["output_paths"], args.inline_xstring)
            result["grouped_equal"] = all(result["source"][key] == result["grouped"][key]
                                          for key in ("rows", "unordered_row_sha256_sum"))
            result["equal"] = result["equal"] and result["grouped_equal"]
        result["seconds"] = round(time.perf_counter() - started, 3)
        result["oracle"] = dict(reader="openpyxl", version=__import__("openpyxl").__version__,
                                inline_xstring_single_pass=args.inline_xstring,
                                storage_guard="reject s/str cells in all worksheets" if args.inline_xstring else None,
                                script_sha256=sha(__file__))
        args.report.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(result), flush=True)
        if not result["equal"]:
            raise SystemExit(1)
        return

    import psutil
    from excel_table_tools import TableSource, inspect_workbook, run_tables
    process, finished, peak = psutil.Process(), threading.Event(), {"rss": 0, "private": 0}
    def sample():
        while not finished.wait(.1):
            memory = process.memory_info()
            peak["rss"] = max(peak["rss"], memory.rss)
            peak["private"] = max(peak["private"], getattr(memory, "private", 0))
    watcher = threading.Thread(target=sample)
    watcher.start()
    last = 0
    def progress(detail):
        nonlocal last
        if time.monotonic() - last >= 30:
            last = time.monotonic()
            # Do not print sheet names or business values.
            print(json.dumps({k: detail.get(k) for k in ("phase", "completed", "unit")}), flush=True)
    sources = ([TableSource(Path(path), inspect_workbook(path)[0]) for path in
                json.loads(args.input_report.read_text(encoding="utf-8"))["output_paths"]]
               if args.input_report else [TableSource(args.source, inspect_workbook(args.source)[0])])
    identity = [{"size": source.path.stat().st_size, "sha256": sha(source.path)} for source in sources]
    started = time.perf_counter()
    try:
        split_column = None
        if args.mode == "split_column":
            from excel_table_io import TableReader
            with TableReader(sources[0].path, sources[0].sheet) as reader:
                split_column = reader.headers[args.column_index]
        result = run_tables(sources, args.output, mode=args.mode, rows_per_file=100_000,
                            provenance=False, split_column=split_column, progress=progress)
    finally:
        seconds = time.perf_counter() - started
        finished.set()
        watcher.join()
    unchanged = all(sha(source.path) == before["sha256"] for source, before in zip(sources, identity))
    report = dict(mode=args.mode, seconds=round(seconds, 3), peak_rss_mib=round(peak["rss"] / 2**20, 2),
                  peak_private_mib=round(peak["private"] / 2**20, 2), input=identity, source_unchanged=unchanged,
                  rows=result.rows, output_files=len(result.files),
                  output=[dict(size=path.stat().st_size, sha256=sha(path)) for path in result.files],
                  output_paths=[str(path) for path in result.files],
                  code={name: sha(name) for name in ("excel_table_io.py", "excel_table_tools.py")})
    args.report.write_text(json.dumps(report, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps({k: v for k, v in report.items() if k not in ("output_paths", "input", "output", "code")}), flush=True)


if __name__ == "__main__":
    main()
