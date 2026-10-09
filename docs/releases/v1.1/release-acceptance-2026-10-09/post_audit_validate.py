"""Revalidate saved performance outputs without rerunning timing or exposing values."""
import json
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import performance_benchmark as benchmark


def addresses(path, part):
    cells, merges = set(), []
    count = 0
    with ZipFile(path) as archive, archive.open(part) as stream:
        for _, node in ET.iterparse(stream, events=("end",)):
            if node.tag == benchmark.NS + "c":
                cells.add(benchmark.point(node.get("r")))
                count += 1
                node.clear()
            elif node.tag == benchmark.NS + "mergeCell":
                merges.append(node.get("ref"))
            elif node.tag == benchmark.NS + "row":
                node.clear()
    assert count == len(cells), "Duplicate physical cell address"
    return cells, merges


def moved_cell_rejected():
    namespace = benchmark.NS[1:-1]
    raw = ('<worksheet xmlns="' + namespace + '"><sheetData><row r="1">'
           '<c r="A1"><v>42</v></c></row></sheetData></worksheet>').encode()
    wrong = raw.replace(b'r="1"', b'r="2"').replace(b'r="A1"', b'r="A2"')
    with tempfile.TemporaryDirectory(prefix="excel-perf-validator-audit-") as directory:
        source, output = Path(directory) / "source.xlsx", Path(directory) / "wrong.xlsx"
        for path, body in ((source, raw), (output, wrong)):
            with ZipFile(path, "w") as archive:
                archive.writestr("xl/worksheets/sheet1.xml", body)
        cell = ET.fromstring(raw).find(".//" + benchmark.NS + "c")
        templates = {"xl/worksheets/sheet1.xml": ({"A1": benchmark.cell_value(cell)}, 0, 0, [], 1)}
        try:
            benchmark.verify(source, output, 1, templates)
        except AssertionError:
            return True
    return False


def main():
    evidence = Path(__file__).resolve().parent
    destination = evidence / "post-audit-validation.json"
    assert not destination.exists(), "Preserve prior audit evidence"
    reports = [evidence / "performance-initial.json", evidence / "performance-final.json"]
    original_report_hashes = {path.name: benchmark.sha(path) for path in reports}
    source_hash = benchmark.sha(benchmark.SOURCE)
    templates = {}
    with ZipFile(benchmark.SOURCE) as archive:
        for part in ("xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
            raw = archive.read(part)
            span = max(benchmark.point(node.get("r"))[0]
                       for node in ET.fromstring(raw).findall(".//" + benchmark.NS + "c"))
            templates[part] = (*benchmark.expected_template(raw), span)
    results, fixtures = [], {}
    for path in reports:
        report = json.loads(path.read_text(encoding="utf-8"))
        assert report["input_sha256_before"] == report["input_sha256_after"] == source_hash
        for item in report["results"]:
            assert item["passed"] and item["exit_code"] == 0 and item["stop_reason"] is None
            output = Path(item["output"])
            factor = item["factor"]
            source = output.parent / ("daily-x" + str(factor) + ".xlsx")
            assert benchmark.sha(output) == item["output_sha256"]
            input_hash = benchmark.sha(source)
            fixtures.setdefault(factor, set()).add(input_hash)
            cells_checked = benchmark.verify(source, output, factor, templates)
            assert cells_checked == item["validated_cells"]
            assert item["regions"] == templates["xl/worksheets/sheet1.xml"][1] * factor + templates["xl/worksheets/sheet2.xml"][1]
            assert item["filled"] == templates["xl/worksheets/sheet1.xml"][2] * factor + templates["xl/worksheets/sheet2.xml"][2]
            address_count = 0
            for part in templates:
                expected, merges = addresses(source, part)
                for ref in merges:
                    first, last = ref.split(":")
                    r1, c1 = benchmark.point(first)
                    r2, c2 = benchmark.point(last)
                    if r2 > r1:
                        assert c1 == c2
                        expected.update((row, c1) for row in range(r1, r2 + 1))
                actual, _ = addresses(output, part)
                assert actual == expected, "Output physical addresses differ from input plus selected fills"
                address_count += len(actual)
            assert address_count == cells_checked
            results.append(dict(report=path.name, implementation=item["implementation"], factor=factor,
                                validated_cells=cells_checked, exact_address_set_matches=True,
                                output_hash_matches_record=True, input_sha256=input_hash, passed=True))
    assert all(len(hashes) == 1 for hashes in fixtures.values()), "Versions used different fixtures"
    rejected = moved_cell_rejected()
    assert rejected
    assert benchmark.sha(benchmark.SOURCE) == source_hash
    assert all(benchmark.sha(path) == original_report_hashes[path.name] for path in reports)
    filesystem_path = evidence / "filesystem-results-final-audit.json"
    filesystem = json.loads(filesystem_path.read_text(encoding="utf-8"))
    assert filesystem["passed"] and filesystem["temporary_directory_removed"]
    assert all(case["passed"] for case in filesystem["cases"])
    collision = next(case for case in filesystem["cases"] if case["case"] == "existing_directory_at_first_output_name")
    assert collision["auto_suffix"] and collision["source_preserved"]
    current_engine_hash = benchmark.sha(benchmark.REPO / "excel_unmerge_fill.py")
    assert filesystem["engine_sha256"] == current_engine_hash
    result = dict(recorded_utc=datetime.now(timezone.utc).isoformat(),
                  scope="Existing timed outputs revalidated; original timing reports unchanged; no new timing runs.",
                  engine_sha256=current_engine_hash,
                  validator_sha256=benchmark.sha(evidence / "performance_benchmark.py"),
                  timing_report_sha256=original_report_hashes,
                  original_timing_reports_unchanged=True, original_source_unchanged=True,
                  same_factor_inputs_byte_identical_across_versions=True,
                  synthetic_out_of_range_cell_rejected=True,
                  filesystem_audit_report=filesystem_path.name, filesystem_audit_passed=True,
                  results=results, passed=True)
    destination.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(dict(results=len(results), total_cell_comparisons=sum(r["validated_cells"] for r in results),
                          passed=True, report=str(destination)), ensure_ascii=False))


if __name__ == "__main__":
    main()
