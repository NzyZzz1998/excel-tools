"""Force the disk-backed path through the existing semantic regression suite."""
import io
import unittest
from unittest import mock
from xml.etree import ElementTree as ET

import test_engine_regressions as regressions
import test_excel_unmerge as baseline
from test_excel_unmerge import NS, c, makebook, mod, sha, sheet

DOM_TRANSFORM = mod.transform_sheet


class ForcedStreamingWorkbookTests(baseline.UnmergeTests):
    def setUp(self):
        super().setUp()
        patcher = mock.patch.object(mod, "STREAM_THRESHOLD", 0)
        patcher.start()
        self.addCleanup(patcher.stop)


class ForcedStreamingRegressions(regressions.EngineRegressionTests):
    def setUp(self):
        super().setUp()
        self.addCleanup(mock.patch.stopall)
        mock.patch.object(mod, "STREAM_THRESHOLD", 0).start()
        mock.patch.object(mod, "transform_sheet", new=self.stream_bytes).start()

    @staticmethod
    def stream_bytes(raw, all_merges=False):
        output = io.BytesIO()
        count, filled = mod.transform_sheet_stream(io.BytesIO(raw), output, all_merges)
        return (output.getvalue() if count else raw), count, filled

    def test_unordered_rows_and_missing_target_rows_are_sorted(self):
        raw = ('<worksheet xmlns="' + NS + '"><dimension ref="A1"/><sheetData>'
               '<row r="5">' + c("C5", "untouched") + '</row>'
               '<row r="1" xmlns:s="' + NS + '" xml:space="preserve">'
               '<c r="A1" t="inlineStr"><s:is><s:t>  group  </s:t></s:is></c></row>'
               '<row r="3" xml:space="default">' + c("B3", "ordinary") + '</row>'
               '</sheetData><mergeCells><mergeCell ref="A1:A4"/></mergeCells></worksheet>').encode()
        transformed, count, filled = self.stream_bytes(raw)
        from xml.dom import minidom
        document = minidom.parseString(transformed)
        try:
            rows = document.getElementsByTagNameNS(NS, "row")
            self.assertEqual([row.getAttribute("r") for row in rows], ["1", "2", "3", "4", "5"])
            values = [node.firstChild.data for node in document.getElementsByTagNameNS(NS, "t")]
            self.assertEqual(values, ["  group  ", "  group  ", "  group  ", "ordinary", "  group  ", "untouched"])
            self.assertEqual((count, filled), (1, 3))
        finally:
            document.unlink()

    def test_streaming_matches_dom_for_values_merges_attributes_and_comments(self):
        fixtures = [
            sheet({1: [c("A1", "header"), c("E1", 45292, "num", 1)],
                   2: [c("B2", "rectangle"), c("E2", 0, "bool"), c("G2", 0, "shared")],
                   3: [c("F3", "1+1", "formula")], 9: [c("Z9", "remote")]},
                  ["A1:C1", "B2:D4", "E1:E2", "G2:G4"]),
            sheet({1: [c("A1", kind="blank")], 3: [c("C3", "unrelated")]}, ["A1:B2"]),
        ]
        # 第一组不把 E2 的独立布尔值放入 E1:E2 合并中。
        fixtures[0] = fixtures[0].replace(b'<c r="E2" t="b"><v>0</v></c>',
                                         b'<c r="F2" t="b"><v>0</v></c>')

        def semantic(node):
            return node.tag, sorted(node.attrib.items()), node.text, tuple(semantic(child) for child in node)

        for raw in fixtures:
            for all_merges in (False, True):
                with self.subTest(all_merges=all_merges):
                    expected, regions, filled = DOM_TRANSFORM(raw, all_merges)
                    actual, actual_regions, actual_filled = self.stream_bytes(raw, all_merges)
                    self.assertEqual((actual_regions, actual_filled), (regions, filled))
                    self.assertEqual(semantic(ET.fromstring(actual)), semantic(ET.fromstring(expected)))

    def test_no_matching_merges_do_not_validate_unused_coordinates(self):
        invalid_rows = [
            '<row><c r="A1"><v>1</v></c></row>',
            '<row r="1"><c><v>1</v></c></row>',
            '<row r="1">' + c("A1", "first") + c("a1", "second") + '</row>',
        ]
        for rows in invalid_rows:
            raw = ('<worksheet xmlns="' + NS + '"><sheetData>' + rows + '</sheetData>'
                   '<mergeCells><mergeCell ref="A1:B1"/></mergeCells></worksheet>').encode()
            with self.subTest(rows=rows):
                self.assertEqual(self.stream_bytes(raw), (raw, 0, 0))
                self.assertEqual(DOM_TRANSFORM(raw), (raw, 0, 0))
                with self.assertRaises(ValueError):
                    self.stream_bytes(raw, all_merges=True)

    def test_utf16_and_later_sheet_rejection_leave_no_output_and_close_spools(self):
        raw = sheet({1: [c("A1", "group")]}, ["A1:A3"])
        raw = ('<?xml version="1.0" encoding="utf-16"?>' + raw.decode()).encode("utf-16")
        output, count, filled = self.stream_bytes(raw)
        self.assertEqual((count, filled), (1, 2))
        self.assertIn(b'encoding="utf-8"', output)
        source = makebook(self.root / "later-rejection.xlsx", [
            ("first", raw), ("second", sheet({1: [c("A1", "1+1", "formula")]}, ["A1:A2"]))])
        before = sha(source)
        real_temp = mod.TemporaryFile
        opened = []

        def tracked_temp(*args, **kwargs):
            value = real_temp(*args, **kwargs)
            opened.append(value)
            return value

        with mock.patch.object(mod, "TemporaryFile", new=tracked_temp):
            with self.assertRaisesRegex(ValueError, "公式"):
                mod.process_file(source)
        self.assertTrue(opened)
        self.assertTrue(all(value.closed for value in opened))
        self.assertEqual(list(self.root.iterdir()), [source])
        self.assertEqual(sha(source), before)

    def test_each_spool_write_failure_closes_temporary_files_without_output(self):
        source = makebook(self.root / "spool-failure.xlsx", [
            ("data", sheet({1: [c("A1", "group")]}, ["A1:A3"]))])
        before = sha(source)
        real_temp = mod.TemporaryFile

        class FailingSpool:
            def __init__(self, handle, fail):
                self.handle, self.fail = handle, fail

            def __getattr__(self, name):
                return getattr(self.handle, name)

            def __enter__(self):
                return self

            def __exit__(self, *args):
                self.handle.close()

            def write(self, data):
                if self.fail:
                    raise OSError("injected temporary disk write failure")
                return self.handle.write(data)

        for failing_number in (1, 2, 3):
            opened = []

            def failing_temp(*args, **kwargs):
                handle = real_temp(*args, **kwargs)
                opened.append(handle)
                return FailingSpool(handle, len(opened) == failing_number)

            with self.subTest(spool=failing_number), mock.patch.object(mod, "TemporaryFile", new=failing_temp):
                with self.assertRaisesRegex(OSError, "temporary disk write failure"):
                    mod.process_file(source)
            self.assertTrue(opened)
            self.assertTrue(all(handle.closed for handle in opened))
            self.assertEqual(list(self.root.iterdir()), [source])
            self.assertEqual(sha(source), before)

    def test_zip_payloads_are_copied_in_chunks_with_metadata_preserved(self):
        from zipfile import ZipFile
        source = makebook(self.root / "chunks.xlsm", [
            ("data", sheet({1: [c("A1", "group")]}, ["A1:A2"]))],
            {"opaque/large.bin": b"synthetic" * (512 * 1024)})
        original_read = ZipFile.read

        def metadata_only_read(archive, name, *args, **kwargs):
            self.assertIn(name, ("xl/workbook.xml", "xl/_rels/workbook.xml.rels"))
            return original_read(archive, name, *args, **kwargs)

        with mock.patch.object(ZipFile, "read", new=metadata_only_read):
            output, stats = mod.process_file(source)
        self.assertEqual(stats, [("data", 1, 1)])
        with ZipFile(source) as before, ZipFile(output) as after:
            self.assertEqual(before.namelist(), after.namelist())
            for name in before.namelist():
                if name != "xl/worksheets/sheet1.xml":
                    self.assertEqual(before.read(name), after.read(name))
                a, b = before.getinfo(name), after.getinfo(name)
                for attr in ("date_time", "comment", "extra", "create_system", "external_attr",
                             "internal_attr", "compress_type"):
                    self.assertEqual(getattr(a, attr), getattr(b, attr), (name, attr))


if __name__ == "__main__":
    unittest.main(verbosity=2)
