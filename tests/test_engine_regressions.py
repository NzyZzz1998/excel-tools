"""回归范围：XML 语义保真、富值保护、稀疏处理和写入失败。"""

import tempfile
import unittest
from pathlib import Path
from unittest import mock
from xml.dom import minidom
from xml.etree import ElementTree as ET

from test_excel_unmerge import NS, c, load, makebook, mod, sha, sheet, xml


class EngineRegressionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory(prefix="excel-engine-regression-")
        self.root = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_text_carriage_returns_and_unrelated_cells_keep_their_values(self):
        raw = sheet({1: [
            '<c r="A1" t="inlineStr"><is><t xml:space="preserve">'
            'first&#13;second &amp; &lt; &gt; 雪</t></is></c>',
            '<c r="B1" t="inlineStr"><is><t xml:space="preserve">'
            'untouched&#13;row&#10;line&#9;tab</t></is></c>',
        ]}, ["A1:A2"])
        source = makebook(self.root / "text.xlsx", [("数据", raw)])
        output, stats = mod.process_file(source)
        with load(output) as workbook:
            self.assertEqual(workbook.active["A1"].value, "first\rsecond & < > 雪")
            self.assertEqual(workbook.active["A2"].value, "first\rsecond & < > 雪")
            self.assertEqual(workbook.active["B1"].value, "untouched\rrow\nline\ttab")
        self.assertEqual(stats, [("数据", 1, 1)])

    def test_attribute_whitespace_and_xml_escapes_survive(self):
        extra = ('<dataValidations count="1"><dataValidation type="whole" '
                 'sqref="B1" prompt="first&#10;second&#9;tab&#13;end '
                 '&amp; &quot;quoted&quot; &lt;value&gt;">'
                 '<formula1>0</formula1></dataValidation></dataValidations>')
        raw = sheet({1: [c("A1", "value")]}, ["A1:A2"], extra)
        transformed, _, _ = mod.transform_sheet(raw)
        path = "{" + NS + "}dataValidations/{" + NS + "}dataValidation"
        before = ET.fromstring(raw).find(path)
        after = ET.fromstring(transformed).find(path)
        self.assertEqual(before.attrib, after.attrib)
        self.assertEqual(after.attrib["prompt"], 'first\nsecond\ttab\rend & "quoted" <value>')

    def test_prefixed_namespace_cdata_and_processing_instruction_survive(self):
        raw = ('<?xml version="1.0"?><?probe unchanged?>'
               '<x:worksheet xmlns:x="' + NS + '"><x:sheetData><x:row r="1">'
               '<x:c r="A1" t="inlineStr"><x:is><x:t xml:space="preserve">'
               '<![CDATA[ 雪<&"\tfirst\nsecond ]]></x:t></x:is></x:c>'
               '<!-- keep row comment --></x:row></x:sheetData>'
               '<x:mergeCells count="1"><x:mergeCell ref="A1:A2"/>'
               '</x:mergeCells></x:worksheet>').encode("utf-8")
        transformed, count, filled = mod.transform_sheet(raw)
        document = minidom.parseString(transformed)
        try:
            self.assertEqual(document.documentElement.prefix, "x")
            values = document.getElementsByTagNameNS(NS, "t")
            self.assertEqual([node.firstChild.data for node in values],
                             [' 雪<&"\tfirst\nsecond '] * 2)
            self.assertIn(b"<?probe unchanged?>", transformed)
            self.assertIn(b"<!-- keep row comment -->", transformed)
            self.assertEqual((count, filled), (1, 1))
        finally:
            document.unlink()

    def test_copied_payload_keeps_local_and_shadowed_namespaces(self):
        cases = [
            # 前缀只在源单元格声明，离开源格后仍需可解析。
            ('<worksheet xmlns="' + NS + '"><sheetData><row r="1">'
             '<c r="A1" xmlns:s="' + NS + '" t="inlineStr">'
             '<s:is><s:t>value</s:t></s:is></c></row></sheetData>'
             '<mergeCells><mergeCell ref="A1:A2"/></mergeCells></worksheet>'),
            # 源行与目标行给同名前缀不同含义，复制子树需保留源含义。
            ('<worksheet xmlns="' + NS + '"><sheetData><row r="1" xmlns:s="' + NS + '">'
             '<c r="A1" t="inlineStr"><s:is><s:t>value</s:t></s:is></c></row>'
             '<row r="2" xmlns:s="urn:unrelated"><c r="B2"/></row></sheetData>'
             '<mergeCells><mergeCell ref="A1:A2"/></mergeCells></worksheet>'),
            # worksheet前缀在目标行被重绑定，新c仍需属于worksheet NS。
            ('<s:worksheet xmlns:s="' + NS + '"><s:sheetData><s:row r="1">'
             '<s:c r="A1" t="inlineStr"><s:is><s:t>value</s:t></s:is></s:c></s:row>'
             '<x:row xmlns:x="' + NS + '" xmlns:s="urn:unrelated" r="2">'
             '<x:c r="B2"/></x:row></s:sheetData>'
             '<s:mergeCells><s:mergeCell ref="A1:A2"/></s:mergeCells></s:worksheet>'),
            # 默认namespace在源格内被重绑定；不能覆盖新c的namespace。
            ('<worksheet xmlns="' + NS + '" xmlns:s="' + NS + '"><sheetData><row r="1">'
             '<s:c r="A1" xmlns="urn:unrelated" t="inlineStr">'
             '<s:is><s:t>value</s:t></s:is></s:c></row></sheetData>'
             '<mergeCells><mergeCell ref="A1:A2"/></mergeCells></worksheet>'),
        ]
        for raw in cases:
            with self.subTest(raw=raw):
                transformed, count, filled = mod.transform_sheet(raw.encode())
                root = ET.fromstring(transformed)
                target = root.find('.//{' + NS + '}c[@r="A2"]')
                self.assertIsNotNone(target)
                text = target.find('{' + NS + '}is/{' + NS + '}t')
                self.assertIsNotNone(text)
                self.assertEqual(text.text, "value")
                self.assertEqual((count, filled), (1, 1))

    def test_selected_metadata_anchor_rejects_whole_workbook_without_output(self):
        metadata_cells = [
            '<c r="A1" t="e" vm="1"><v>#VALUE!</v></c>',
            '<c r="A1" cm="1"><v>5</v></c>',
            '<c r="A1"><v>5</v><extLst><ext uri="future-value"/></extLst></c>',
        ]
        for index, anchor in enumerate(metadata_cells):
            with self.subTest(anchor=anchor):
                source = makebook(self.root / ("metadata%d.xlsx" % index), [
                    ("普通", sheet({1: [c("A1", "first")]}, ["A1:A2"])),
                    ("富值", sheet({1: [anchor]}, ["A1:A2"])),
                ])
                before = set(self.root.iterdir())
                digest = sha(source)
                with self.assertRaisesRegex(ValueError, "富值.*元数据"):
                    mod.process_file(source)
                self.assertEqual(set(self.root.iterdir()), before)
                self.assertEqual(sha(source), digest)

    def test_unselected_metadata_cells_and_horizontal_anchor_are_preserved(self):
        raw = sheet({1: [
            c("A1", "value"),
            '<c r="B1" t="e" vm="1"><v>#VALUE!</v></c>',
            '<c r="C1" cm="1"><v>5</v></c>',
        ]}, ["A1:A2", "C1:D1"])
        source = makebook(self.root / "unselected.xlsx", [("数据", raw)])
        output, stats = mod.process_file(source)
        for reference in ("B1", "C1"):
            path = './/{' + NS + '}c[@r="' + reference + '"]'
            self.assertEqual(ET.tostring(xml(source).find(path)),
                             ET.tostring(xml(output).find(path)))
        self.assertEqual(stats, [("数据", 1, 1)])
        with self.assertRaisesRegex(ValueError, "C1:D1.*元数据"):
            mod.process_file(source, all_merges=True)

    def test_sheet_selection_does_not_inspect_unselected_metadata_anchor(self):
        source = makebook(self.root / "selection.xlsx", [
            ("普通", sheet({1: [c("A1", "value")]}, ["A1:A2"])),
            ("富值", sheet({1: ['<c r="A1" vm="1"><v>1</v></c>']}, ["A1:A2"])),
        ])
        output, stats = mod.process_file(source, sheets=["普通"])
        self.assertTrue(output.is_file())
        self.assertEqual(stats, [("普通", 1, 1)])

    def test_sparse_wide_merge_does_not_enumerate_the_empty_grid(self):
        raw = sheet({1: [c("A1", "value")]}, ["A1:XFD200"])
        real_address = mod.address
        address_calls = 0

        def bounded_address(row, column):
            nonlocal address_calls
            address_calls += 1
            # 工作量应跟新增的 200 行有关，而不是 327 万个空坐标。
            if address_calls > 10000:
                raise AssertionError("稀疏合并仍在枚举整块空白网格")
            return real_address(row, column)

        with mock.patch.object(mod, "address", new=bounded_address):
            transformed, count, filled = mod.transform_sheet(raw)
        root = ET.fromstring(transformed)
        merges = root.findall("{" + NS + "}mergeCells/{" + NS + "}mergeCell")
        self.assertEqual((count, filled), (1, 199))
        self.assertEqual(len(merges), 200)
        self.assertEqual(merges[-1].attrib["ref"], "A200:XFD200")

    def test_wide_merge_still_rejects_far_right_hidden_content(self):
        raw = sheet({1: [c("A1", "value")], 199: [c("XFD199", 0, "num")]},
                    ["A1:XFD200"])
        with self.assertRaisesRegex(ValueError, "XFD199.*独立内容"):
            mod.transform_sheet(raw)

    def test_hidden_metadata_without_value_is_not_overwritten(self):
        targets = ['<c r="A2" vm="1"/>', '<c r="A2" cm="1"/>',
                   '<c r="A2"><extLst><ext uri="future-value"/></extLst></c>']
        for index, target in enumerate(targets):
            with self.subTest(target=target):
                source = makebook(self.root / ("hidden-metadata%d.xlsx" % index), [
                    ("数据", sheet({1: [c("A1", "value")], 2: [target]}, ["A1:A2"]))])
                before = {path: sha(path) for path in self.root.iterdir()}
                with self.assertRaisesRegex(ValueError, "A2.*元数据"):
                    mod.process_file(source)
                self.assertEqual({path: sha(path) for path in self.root.iterdir()}, before)

    def test_many_rows_keep_values_order_and_row_attributes(self):
        raw = sheet({row: [c("B" + str(row), row, "num")]
                     for row in range(1, 5001)}, ["A1:A2"])
        raw = raw.replace(b'<row r="1">', b'<row r="1" hidden="1" spans="1:2">'
                          b'<c r="A1" t="inlineStr"><is><t>group</t></is></c>')
        transformed, count, filled = mod.transform_sheet(raw)
        root = ET.fromstring(transformed)
        rows = root.findall("{" + NS + "}sheetData/{" + NS + "}row")
        self.assertEqual([int(row.attrib["r"]) for row in rows], list(range(1, 5001)))
        self.assertEqual(rows[0].attrib, {"r": "1", "hidden": "1"})
        for row in rows:
            number = int(row.attrib["r"])
            value = row.find('{'+NS+'}c[@r="B'+str(number)+'"]/{'+NS+'}v')
            self.assertEqual(value.text, str(number))
        self.assertEqual([cell.attrib["r"] for cell in rows[1]], ["A2", "B2"])
        self.assertEqual((count, filled), (1, 1))

    def test_inaccurate_dimension_keeps_remote_values_readable(self):
        raw = sheet({1: [c("A1", "group")], 20: [c("Z20", 73, "num")]},
                    ["A1:A2"]).replace(b'dimension ref="A1:K20"', b'dimension ref="A1"')
        source = makebook(self.root / "small-dimension.xlsx", [("数据", raw)])
        output, stats = mod.process_file(source)
        self.assertEqual(xml(output).find("{" + NS + "}dimension").attrib["ref"], "A1:Z20")
        # 默认 read_only 遍历必须看到远离拆分区域的普通数值，不手动 reset_dimensions。
        with load(output) as workbook:
            values = list(workbook.active.values)
            self.assertEqual((len(values), len(values[0])), (20, 26))
            self.assertEqual(values[19][25], 73)
            self.assertEqual(values[1][0], "group")
        self.assertEqual(stats, [("数据", 1, 1)])

    def test_dimension_also_covers_retained_empty_horizontal_merge(self):
        raw = sheet({1: [c("A1", "group")], 20: [c("Z20", 73, "num")]},
                    ["A1:A2", "AA30:AC30"]).replace(
                        b'dimension ref="A1:K20"', b'dimension ref="A1"')
        transformed, count, filled = mod.transform_sheet(raw)
        root = ET.fromstring(transformed)
        self.assertEqual(root.find("{" + NS + "}dimension").attrib["ref"], "A1:AC30")
        remaining = root.findall("{" + NS + "}mergeCells/{" + NS + "}mergeCell")
        self.assertEqual([node.attrib["ref"] for node in remaining], ["AA30:AC30"])
        self.assertEqual((count, filled), (1, 1))

    def test_dimension_is_not_shrunk_or_rewritten_without_matching_merges(self):
        broad = sheet({1: [c("A1", "group")]}, ["A1:A2"]).replace(
            b'dimension ref="A1:K20"', b'dimension ref="A1:AZ100"')
        transformed, _, _ = mod.transform_sheet(broad)
        self.assertEqual(ET.fromstring(transformed).find("{" + NS + "}dimension").attrib["ref"],
                         "A1:AZ100")
        unmatched = sheet({20: [c("Z20", 73, "num")]}, ["A1:B1"]).replace(
            b'dimension ref="A1:K20"', b'dimension ref="A1"')
        self.assertEqual(mod.transform_sheet(unmatched), (unmatched, 0, 0))

    def test_partial_write_failure_removes_only_the_new_output(self):
        source = makebook(self.root / "source.xlsx", [
            ("数据", sheet({1: [c("A1", "value")]}, ["A1:A2"]))])
        previous = self.root / "source_拆分填充.xlsx"
        previous.write_bytes(b"previous result must survive")
        before = {path: sha(path) for path in self.root.iterdir()}
        real_write = mod.ZipFile.writestr
        writes = 0

        def failing_write(archive, *args, **kwargs):
            nonlocal writes
            writes += 1
            if writes == 2:
                raise OSError("injected disk write failure")
            return real_write(archive, *args, **kwargs)

        with mock.patch.object(mod.ZipFile, "writestr", new=failing_write):
            with self.assertRaisesRegex(OSError, "injected disk write failure"):
                mod.process_file(source)
        self.assertEqual({path: sha(path) for path in self.root.iterdir()}, before)


if __name__ == "__main__":
    unittest.main(verbosity=2)
