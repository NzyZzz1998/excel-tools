"""Independent oracle regressions; run directly, separately from product tests."""
from pathlib import Path
import tempfile
import unittest
import zipfile

from measure_tables import decode_inline_xstring, require_inline_strings


class OracleTests(unittest.TestCase):
    def test_excel_checked_xstring_cases(self):
        for raw, expected in (
            ("A_x003D_B", "A=B"),
            ("_x003D_text", "=text"),
            ("_x005F_x003D_", "_x003D_"),
            ("A_x000D_B", "A\rB"),
            ("A\rB", "A\rB"),
            ("_xD83D__xDE00_", "\U0001f600"),
        ):
            with self.subTest(raw=raw):
                self.assertEqual(decode_inline_xstring(raw), expected)

    def storage_fixture(self, folder, cell):
        path = Path(folder) / "fixture.xlsx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("xl/worksheets/sheet1.xml", (
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData><row r="1">' + cell + '</row></sheetData></worksheet>'
            ))
        return path

    def test_inline_guard_accepts_inline_and_numbers(self):
        with tempfile.TemporaryDirectory() as folder:
            path = self.storage_fixture(folder, '<c t="inlineStr"><is><t>A</t></is></c><c><v>1</v></c>')
            self.assertEqual(require_inline_strings(path), {"inlineStr": 1, "n": 1})

    def test_guard_rejects_shared_or_cached_strings(self):
        with tempfile.TemporaryDirectory() as folder:
            for kind in ("s", "str"):
                with self.subTest(kind=kind):
                    path = self.storage_fixture(folder, '<c t="' + kind + '"><v>0</v></c>')
                    with self.assertRaises(ValueError):
                        require_inline_strings(path)


if __name__ == "__main__":
    unittest.main(verbosity=2)
