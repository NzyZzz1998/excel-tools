"""Recreate the synthetic fixtures used to disprove the inherited-style candidate."""
import json
from pathlib import Path
import sys
import tempfile

repo = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(repo / 'tests'))
import test_excel_unmerge as fixture

root = Path(tempfile.mkdtemp(prefix='post-release-style-', dir=repo / 'testfile'))
cases = []
for name, prefix, row_attrs, merge, target, all_merges in (
    ('row_style', '', ' s="1" customFormat="1"', 'A1:A2', 'A2', False),
    ('column_style', '<cols><col min="1" max="1" style="1"/></cols>', '', 'A1:B1', 'B1', True),
):
    raw = ('<worksheet xmlns="' + fixture.NS + '"><dimension ref="' + merge + '"/>'
           + prefix + '<sheetData><row r="1"' + row_attrs + '><c r="A1"><v>45292</v></c>'
           '</row></sheetData><mergeCells count="1"><mergeCell ref="' + merge
           + '"/></mergeCells></worksheet>').encode()
    source = fixture.makebook(root / (name + '.xlsx'), [('SYNTHETIC', raw)])
    before = fixture.sha(source)
    output, stats = fixture.mod.process_file(source, all_merges=all_merges)
    cases.append({'name': name, 'source': str(source), 'output': str(output), 'stats': stats,
                  'source_unchanged': fixture.sha(source) == before, 'target': target})
(root / 'fixtures.json').write_text(json.dumps(cases, ensure_ascii=False, indent=2), encoding='utf-8')
print(root)
