"""Independent, batch-specific monthly headers and group-membership oracle.

Reads only local workbooks with openpyxl, not application readers. Reports never
contain source paths, sheet names, grouping labels, or business cell values.
The accompanying storage inventory and full-content oracle establish the
inline-string precondition for this specific batch; this is not a generic XLSX
storage validator or a performance measurement.
"""

import argparse
from datetime import date, datetime, time as datetime_time, timedelta
from decimal import Decimal
import hashlib
import json
from pathlib import Path
import time
import warnings

import openpyxl

import measure_tables


def file_sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def decoded_value(cell):
    value = cell.value
    if cell.data_type == "s" and isinstance(value, str):
        value = measure_tables.decode_inline_xstring(value)
    return value


def value_key(cell):
    """Semantic type/value key; number formats do not split ordinary numbers."""
    value = decoded_value(cell)
    if value is None:
        return ("blank",)
    if cell.data_type == "n":
        value = Decimal(str(value))
    elif isinstance(value, (datetime, date, datetime_time)):
        value = value.isoformat()
    elif isinstance(value, timedelta):
        value = (value.days, value.seconds, value.microseconds)
    return cell.data_type, value


def header_key(row):
    # Preserve the full observed width; do not trim or coerce header strings.
    return tuple((value_key(cell), cell.number_format) for cell in row)


def open_rows(path):
    book = openpyxl.load_workbook(path, read_only=True, data_only=False)
    try:
        sheet = book.worksheets[0]
        sheet.reset_dimensions()
        return book, sheet.iter_rows()
    except BaseException:
        book.close()
        raise


def read_header(path):
    book, rows = open_rows(path)
    try:
        return header_key(next(rows))
    finally:
        rows.close()
        book.close()


def check(condition, message):
    if not condition:
        raise AssertionError(message)


def verify(base, inventory_path):
    inventory = json.loads(inventory_path.read_text(encoding="utf-8"))
    check(inventory["summary"]["all_worksheet_parts_covered_by_guard"],
          "Prior worksheet storage inventory did not pass")
    check(inventory["summary"]["all_present_sst_empty"],
          "Prior shared-string storage inventory did not pass")
    decoder_sha = file_sha(measure_tables.__file__)
    check(inventory["script_sha256"] == decoder_sha,
          "Decoder changed after the storage review")
    reports = {}
    paths = {}
    for operation in ("split", "merge", "group"):
        report_path = base / ("monthly-" + operation + ".json")
        check(file_sha(report_path) == inventory["operation_report_sha256"][operation],
              "Operation report changed after the storage review")
        reports[operation] = json.loads(report_path.read_text(encoding="utf-8"))
        paths[operation] = [Path(path) for path in reports[operation]["output_paths"]]
    check([len(paths[name]) for name in ("split", "merge", "group")] == [8, 1, 29],
          "Unexpected output-file counts for this fixed batch")
    source = base / "monthly-base_拆分填充.xlsx"
    all_paths = [source] + [path for group in paths.values() for path in group]
    signatures = [(path.stat().st_size, path.stat().st_mtime_ns) for path in all_paths]
    expected_header = read_header(source)
    check(len(expected_header) == 22, "Unexpected source-header width")
    headers_checked = 0
    for operation in ("split", "merge"):
        for index, path in enumerate(paths[operation], 1):
            check(read_header(path) == expected_header,
                  "Header mismatch in {} output index {}".format(operation, index))
            headers_checked += 1
    keys = set()
    counts = []
    total_rows = blank_key_rows = fully_empty_rows = 0
    last_update = time.monotonic()
    for index, path in enumerate(paths["group"], 1):
        book, rows = open_rows(path)
        group_key = None
        group_rows = 0
        try:
            check(header_key(next(rows)) == expected_header,
                  "Header mismatch in group output index {}".format(index))
            headers_checked += 1
            for row_number, row in enumerate(rows, 2):
                # A missing first-column value never makes an otherwise populated
                # row disappear. Empty strings, zero and False are also data.
                if not any(cell.value is not None for cell in row):
                    fully_empty_rows += 1
                    continue
                key = value_key(row[0])
                if group_key is None:
                    group_key = key
                check(key == group_key,
                      "Mixed grouping keys in output index {} at row {}".format(index, row_number))
                group_rows += 1
                total_rows += 1
                blank_key_rows += key == ("blank",)
                if time.monotonic() - last_update >= 30:
                    print(json.dumps({"checked_group_files": index - 1,
                                      "checked_group_rows": total_rows}), flush=True)
                    last_update = time.monotonic()
        finally:
            rows.close()
            book.close()
        check(group_rows > 0, "Empty group output index {}".format(index))
        check(group_key not in keys, "Repeated grouping key in output index {}".format(index))
        keys.add(group_key)
        counts.append(group_rows)
    check(headers_checked == 38, "Not every output header was checked")
    check(len(keys) == 29, "Unexpected distinct grouping-key count")
    check(total_rows == 757635, "Unexpected total data-row count")
    check(total_rows == reports["group"]["rows"], "Count differs from operation result")
    check(signatures == [(path.stat().st_size, path.stat().st_mtime_ns) for path in all_paths],
          "A workbook changed while the oracle was reading it")
    return {
        "passed": True,
        "source_header_columns": len(expected_header),
        "output_headers_checked": headers_checked,
        "all_headers_value_type_number_format_equal": True,
        "group_files_checked": len(counts),
        "each_file_has_one_decoded_type_value_key": True,
        "all_group_file_keys_distinct": True,
        "distinct_group_keys": len(keys),
        "group_rows_by_output_index": counts,
        "group_data_rows": total_rows,
        "rows_with_blank_first_column": blank_key_rows,
        "fully_empty_rows_skipped": fully_empty_rows,
        "expected_total_rows_match": True,
        "workbook_size_mtime_signatures_unchanged": True,
        "decoder_sha256": decoder_sha,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--base", type=Path, default=Path("testfile/v1.2.0-acceptance"))
    parser.add_argument("--report", type=Path,
                        default=Path(__file__).with_name("monthly-groups.json"))
    parser.add_argument("--inventory", type=Path,
                        default=Path(__file__).with_name("oracle-storage-inventory.json"))
    args = parser.parse_args()
    warnings.filterwarnings("ignore", message="Workbook contains no default style, apply openpyxl's default")
    started = time.perf_counter()
    try:
        result = verify(args.base, args.inventory)
    except Exception as error:
        # Only our assertions carry safe indices/counts. Other exception messages
        # might contain a business pathname and are deliberately not serialized.
        result = {"passed": False, "exception_type": type(error).__name__}
        if isinstance(error, AssertionError):
            result["check_failure"] = str(error)
    result["helper_sha256"] = file_sha(__file__)
    result["elapsed_seconds"] = round(time.perf_counter() - started, 3)
    args.report.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(result), flush=True)
    if not result["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
