"""本轮评审的合成取证脚本；不读取业务文件，不修改程序代码。

从任意目录运行：python <此脚本绝对路径>。
只依赖标准库，输出 JSON；性能数字是当前机器单次样本，不是产品门禁。
"""

import cProfile
import hashlib
import io
import json
import platform
import pstats
import sys
import tempfile
import time
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
from excel_unmerge_fill import process_file, transform_sheet

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def fidelity_probe():
    raw = ('<worksheet xmlns="' + NS + '"><sheetData><row r="1">'
           '<c r="A1" t="inlineStr"><is><t>merge</t></is></c>'
           '<c r="B1" t="inlineStr"><is><t>first&#13;second</t></is></c>'
           '</row></sheetData><mergeCells count="1"><mergeCell ref="A1:A2"/>'
           '</mergeCells><dataValidations count="1"><dataValidation sqref="D1" '
           'type="whole" showInputMessage="1" prompt="one&#10;two">'
           '<formula1>0</formula1></dataValidation></dataValidations></worksheet>')
    parts = {
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',
        "_rels/.rels": '<Relationships xmlns="' + PKG + '"><Relationship Id="wb" Type="' + REL + '/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": '<workbook xmlns="' + NS + '" xmlns:r="' + REL + '"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<Relationships xmlns="' + PKG + '"><Relationship Id="rId1" Type="' + REL + '/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": raw,
    }
    with tempfile.TemporaryDirectory(prefix="excel-review-entities-") as directory:
        source = Path(directory) / "probe.xlsx"
        with ZipFile(source, "w", ZIP_DEFLATED) as archive:
            for name, value in parts.items():
                archive.writestr(name, value.encode("utf-8"))
        original = source.read_bytes()
        output, stats = process_file(source)
        with ZipFile(output) as archive:
            after = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        before = ET.fromstring(raw)
        values = []
        for label, root in (("before", before), ("after", after)):
            text = root.find('.//{' + NS + '}c[@r="B1"]/{' + NS + '}is/{' + NS + '}t').text
            prompt = root.find('{' + NS + '}dataValidations/{' + NS + '}dataValidation').get("prompt")
            values.append({"stage": label, "B1": text, "validation_prompt": prompt})
        return {"values": values, "stats": stats, "source_unchanged": source.read_bytes() == original}


def performance_probe():
    samples = []
    profile_raw = None
    for rows in (5000, 10000, 20000, 40000):
        body = ''.join('<row r="%s"><c r="B%s"><v>1</v></c></row>' % (r, r)
                       for r in range(1, rows + 1))
        for merged in (False, True):
            raw = ('<worksheet xmlns="' + NS + '"><sheetData>' + body + '</sheetData>'
                   + ('<mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells>' if merged else '')
                   + '</worksheet>').encode("utf-8")
            start = time.perf_counter()
            _, count, filled = transform_sheet(raw)
            samples.append({"rows": rows, "merged": merged, "bytes": len(raw),
                            "seconds": round(time.perf_counter() - start, 6),
                            "regions": count, "filled": filled})
            if rows == 20000 and merged:
                profile_raw = raw
    profile = cProfile.Profile()
    profile.runcall(transform_sheet, profile_raw)
    summary = io.StringIO()
    pstats.Stats(profile, stream=summary).strip_dirs().sort_stats("cumulative").print_stats(12)
    return {"samples": samples, "profile_20000_rows": summary.getvalue()}


if __name__ == "__main__":
    print(json.dumps({"python": sys.version, "platform": platform.platform(),
                      "engine_sha256": hashlib.sha256((ROOT / "excel_unmerge_fill.py").read_bytes()).hexdigest(),
                      "fidelity": fidelity_probe(), "performance": performance_probe()},
                     ensure_ascii=False, indent=2))
