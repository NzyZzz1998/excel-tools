"""Generate synthetic Excel probes; never reads business inputs.

Run with the development openpyxl dependency available, then run probe_excel.ps1
with the printed manifest path. Baseline engine is the recorded 96cfbfe revision.
"""
import hashlib
import importlib.util
import json
import subprocess
import sys
import tempfile
import types
from pathlib import Path

REPO = Path(__file__).resolve().parents[5]
sys.path.insert(0, str(REPO / "tests"))
from test_excel_unmerge import NS, makebook

baseline_source = subprocess.check_output(
    ["git", "show", "96cfbfe:excel_unmerge_fill.py"], cwd=REPO)
baseline = types.ModuleType("engine_before_deep_review")
exec(compile(baseline_source, "96cfbfe:excel_unmerge_fill.py", "exec"), baseline.__dict__)
spec = importlib.util.spec_from_file_location("engine_after_deep_review", REPO / "excel_unmerge_fill.py")
current = importlib.util.module_from_spec(spec)
spec.loader.exec_module(current)
work = Path(tempfile.mkdtemp(prefix="excel-tools-release-engine-fixed-"))

cases = [
    ("row_style", '<row r="1" s="1" customFormat="1"><c r="A1"><v>45292</v></c></row>',
     "A1:A2", "", "", False),
    ("column_style", '<row r="1"><c r="A1"><v>45292</v></c></row>',
     "A1:B1", '<cols><col min="1" max="1" style="1"/></cols>', "", True),
    ("xmlspace_cell", '<row r="1"><c r="A1" t="inlineStr" xml:space="preserve">'
     '<is><t>  padded  </t></is></c></row>', "A1:A2", "", "", False),
    ("xmlspace_row", '<row r="1" xml:space="preserve"><c r="A1" t="inlineStr">'
     '<is><t>  padded  </t></is></c></row>', "A1:A2", "", "", False),
    ("xmlspace_target", '<row r="1"><c r="A1" t="inlineStr"><is><t>  padded  </t></is></c></row>'
     '<row r="2" xml:space="preserve"/>', "A1:A2", "", "", False),
    ("xmlspace_root_target_default", '<row r="1"><c r="A1" t="inlineStr">'
     '<is><t>  padded  </t></is></c></row><row r="2" xml:space="default"/>',
     "A1:A2", "", ' xml:space="preserve"', False),
    ("xmlspace_text_default", '<row r="1"><c r="A1" t="inlineStr" xml:space="preserve">'
     '<is><t xml:space="default">  padded  </t></is></c></row>', "A1:A2", "", "", False),
    ("lowercase_cell", '<row r="1"><c r="a1" t="inlineStr"><is><t>sentinel</t></is></c></row>'
     '<row r="2"><c r="a2"/></row>', "A1:A2", "", "", False),
    ("absolute_cell", '<row r="1"><c r="$A$1" t="inlineStr"><is><t>sentinel</t></is></c></row>'
     '<row r="2"><c r="$A$2"/></row>', "A1:A2", "", "", False),
]
results = []
for name, rows, merge, before_rows, attributes, all_merges in cases:
    raw = ('<worksheet xmlns="' + NS + '"' + attributes + '><dimension ref="A1:B2"/>'
           + before_rows + '<sheetData>' + rows + '</sheetData><mergeCells count="1">'
           '<mergeCell ref="' + merge + '"/></mergeCells></worksheet>').encode()
    source = makebook(work / (name + ".xlsx"), [("Data", raw)])
    before, before_stats = baseline.process_file(source, all_merges)
    after, after_stats = current.process_file(source, all_merges)
    results.append(dict(name=name, source=str(source), before=str(before), after=str(after),
                        before_stats=before_stats, after_stats=after_stats,
                        source_sha256=hashlib.sha256(source.read_bytes()).hexdigest(),
                        target="B1" if all_merges else "A2"))
report = dict(baseline_revision="96cfbfe",
              baseline_engine_sha256=hashlib.sha256(baseline_source).hexdigest(),
              current_engine_sha256=hashlib.sha256((REPO / "excel_unmerge_fill.py").read_bytes()).hexdigest(),
              synthetic_only=True, cases=results)
manifest = work / "manifest.json"
manifest.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
print(manifest)
