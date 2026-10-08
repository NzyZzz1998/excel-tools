"""Exercise the real Tk window; business fixtures stay in temporary directories."""

import tempfile
import time
import tkinter as tk
import unittest
from pathlib import Path

from openpyxl import Workbook

from excel_unmerge_gui import Application


class WindowTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.attributes("-alpha", 0)
        self.app = Application(self.root)
        self.root.update()
        self.directory = tempfile.TemporaryDirectory(prefix="excel-gui-test-")
        self.folder = Path(self.directory.name)

    def tearDown(self):
        self.root.destroy()
        self.directory.cleanup()

    def workbook(self, name, merged):
        path = self.folder / name
        book = Workbook()
        book.active["A1"] = "测试内容"
        if merged:
            book.active.merge_cells("A1:A3")
        book.save(path)
        book.close()
        return path

    def test_result_directory_unavailable_until_output_exists(self):
        button = getattr(self.app, "open_button", None)
        self.assertIsNotNone(button, "Results need an open-directory action")
        self.assertIn("disabled", button.state())
        for control in (self.app.choose_button, self.app.file_list, self.app.option,
                        self.app.start_button, self.app.progress, self.app.results,
                        button):
            self.assertGreater(control.winfo_width(), 1)
            self.assertGreater(control.winfo_height(), 1)

    def test_batch_reports_completed_files_and_enables_output_directory(self):
        successful = self.workbook("合并 数据.xlsx", True)
        unchanged = self.workbook("无需处理.xlsx", False)
        missing = self.folder / "不存在.xlsx"
        self.app.files = tuple(map(str, (successful, unchanged, missing)))
        self.app.start()
        self.assertEqual(str(self.app.progress["mode"]), "determinate")
        deadline = time.monotonic() + 10
        while self.app.running and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.app.running, "Batch did not finish")
        self.assertEqual(float(self.app.progress["value"]), 3)
        self.assertEqual(float(self.app.progress["maximum"]), 3)
        self.assertEqual(self.app.completed, 3)
        self.assertEqual(self.app.succeeded, 1)
        self.assertEqual(self.app.unchanged, 1)
        self.assertEqual(self.app.failed, 1)
        self.assertEqual(self.app.output_directory, self.folder.resolve())
        self.assertNotIn("disabled", self.app.open_button.state())
        output = successful.with_name("合并 数据_拆分填充.xlsx")
        self.assertTrue(output.is_file())
        log = self.app.results.get("1.0", "end")
        self.assertIn(str(output), log)
        self.assertIn(str(missing), log)
        self.assertIn(str(unchanged), log)


if __name__ == "__main__":
    unittest.main()
