"""Exercise the real Tk window; business fixtures stay in temporary directories."""

import tempfile
import threading
import time
import tkinter as tk
from tkinter import font as tkfont
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

from excel_unmerge_gui import Application
from excel_unmerge_fill import process_file


class WindowTests(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.attributes("-alpha", 0)
        self.app = Application(self.root)
        self.root.update()
        self.directory = tempfile.TemporaryDirectory(prefix="excel-gui-test-")
        self.folder = Path(self.directory.name)

    def tearDown(self):
        if self.root is not None:
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

    def finish(self):
        deadline = time.monotonic() + 10
        while self.app.running and time.monotonic() < deadline:
            self.root.update()
            time.sleep(0.01)
        self.assertFalse(self.app.running, "Batch did not finish")

    def test_result_directory_unavailable_until_output_exists(self):
        button = getattr(self.app, "open_button", None)
        self.assertIsNotNone(button, "Results need an open-directory action")
        self.assertIn("disabled", button.state())
        for control in (self.app.choose_button, self.app.file_list, self.app.option,
                        self.app.start_button, self.app.progress, self.app.results,
                        self.app.retry_button, self.app.resume_button,
                        self.app.stop_button, button):
            self.assertGreater(control.winfo_width(), 1)
            self.assertGreater(control.winfo_height(), 1)
        for button in (self.app.retry_button, self.app.resume_button, self.app.stop_button):
            self.assertIn("disabled", button.state())

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
        output = successful.with_name("合并 数据_拆分填充.xlsx").resolve()
        self.assertTrue(output.is_file())
        log = self.app.results.get("1.0", "end")
        self.assertIn(str(output), log)
        self.assertIn(str(missing), log)
        self.assertIn(str(unchanged), log)

    def test_retry_only_failed_files_keeps_history_output_and_initial_options(self):
        successful = self.workbook("成功.xlsx", True)
        repaired = self.folder / "待修复.xlsx"
        repaired.write_bytes(b"not an xlsx")
        self.app.files = tuple(map(str, (successful, repaired)))
        self.app.start()
        self.finish()
        previous_log = self.app.results.get("1.0", "end-1c")
        self.assertIn("BadZipFile", previous_log)
        self.assertEqual(self.app.failed_files, (str(repaired),))
        self.assertNotIn("disabled", self.app.retry_button.state())

        book = Workbook()
        book.active["A1"] = "只保留横向"
        book.active.merge_cells("A1:B1")
        book.save(repaired)
        book.close()
        # Changing the next normal run's option must not change a retry's rule.
        self.app.all_merges.set(True)
        self.app.retry_failed()
        self.assertIs(self.app.all_merges.get(), False)
        self.finish()
        self.assertEqual(self.app.failed_files, ())
        self.assertEqual(self.app.succeeded, 1)
        self.assertEqual(self.app.unchanged, 1)
        self.assertEqual(self.app.completed, 2)
        self.assertEqual(self.app.output_directory, self.folder.resolve())
        self.assertEqual(len(list(self.folder.glob("成功_拆分填充*.xlsx"))), 1)
        self.assertFalse(repaired.with_name("待修复_拆分填充.xlsx").exists())
        self.assertTrue(self.app.results.get("1.0", "end").startswith(previous_log))
        self.assertIn("disabled", self.app.retry_button.state())

    def test_stop_and_continue_keep_current_file_safe_and_report_pending(self):
        first = self.workbook("第一份.xlsx", True)
        later = self.workbook("第二份.xlsx", True)
        self.app.files = tuple(map(str, (first, later)))
        entered = threading.Event()
        release = threading.Event()

        def delayed_process(file, all_merges):
            if file == str(first):
                entered.set()
                if not release.wait(5):
                    raise AssertionError("Current file was not released")
            return process_file(file, all_merges=all_merges)

        with patch("excel_unmerge_gui.process_file", side_effect=delayed_process) as mocked:
            self.app.start()
            try:
                self.assertTrue(entered.wait(5))
                deadline = time.monotonic() + 2
                while self.app.current_file is None and time.monotonic() < deadline:
                    self.root.update()
                    time.sleep(.01)
                self.assertEqual(self.app.current_file, str(first))
                self.assertIn(first.name, self.app.status.get())
                self.app.close()
                self.assertTrue(self.app.running)
                self.assertTrue(self.root.winfo_exists())
                self.assertTrue(self.app.stop_event.is_set())
            finally:
                release.set()
                self.finish()
            self.assertEqual(mocked.call_count, 1)
            self.assertEqual(self.app.completed, 1)
            self.assertEqual(self.app.pending_files, (str(later),))
            self.assertIn("已停止", self.app.status.get())
            self.assertIn("待处理 1", self.app.status.get())
            self.assertNotIn("disabled", self.app.resume_button.state())
            self.assertFalse(later.with_name("第二份_拆分填充.xlsx").exists())
            self.app.all_merges.set(True)
            self.app.resume_pending()
            self.finish()
            self.assertEqual(mocked.call_count, 2)
            self.assertIs(mocked.call_args.kwargs["all_merges"], False)
        self.assertEqual(self.app.completed, 2)
        self.assertEqual(self.app.succeeded, 2)
        self.assertEqual(self.app.pending_files, ())
        self.assertIn("disabled", self.app.resume_button.state())

    def test_failure_and_unprocessed_files_have_independent_recovery_buttons(self):
        first = self.folder / "失败.xlsx"
        later = self.workbook("继续.xlsx", True)
        self.app.files = tuple(map(str, (first, later)))

        def fail_and_stop(file, all_merges):
            self.app.stop_event.set()
            raise PermissionError("test access denied")

        with patch("excel_unmerge_gui.process_file", side_effect=fail_and_stop):
            self.app.start()
            self.finish()
        self.assertEqual(self.app.failed_files, (str(first),))
        self.assertEqual(self.app.pending_files, (str(later),))
        self.assertNotIn("disabled", self.app.retry_button.state())
        self.assertNotIn("disabled", self.app.resume_button.state())
        self.assertIn("失败 1", self.app.status.get())
        self.assertIn("待处理 1", self.app.status.get())
        self.workbook(first.name, True)
        self.app.retry_failed()
        self.finish()
        self.assertEqual(self.app.failed_files, ())
        self.assertEqual(self.app.pending_files, (str(later),))
        self.assertIn("待处理 1", self.app.status.get())
        self.assertNotIn("disabled", self.app.resume_button.state())

    def check_thread_memory_failure_recovery(self, target, stage):
        source = self.workbook("线程内存-" + stage + ".xlsx", True)
        original = source.read_bytes()
        callback_errors = []
        self.root.report_callback_exception = lambda *args: callback_errors.append(args[0].__name__)
        self.app.files = (str(source),)
        with patch(target, side_effect=MemoryError("synthetic thread allocation failure")):
            self.root.after(0, self.app.start)
            self.root.update()
        self.assertEqual(callback_errors, [])
        self.finish()
        self.assertEqual(self.app.pending_files, (str(source),))
        self.assertEqual(self.app.failed_files, ())
        self.assertEqual((self.app.completed, self.app.succeeded, self.app.failed), (0, 0, 0))
        self.assertEqual(list(self.folder.glob("*_拆分填充*.xlsx")), [])
        self.assertEqual(source.read_bytes(), original)
        self.assertNotIn("disabled", self.app.start_button.state())
        self.assertNotIn("disabled", self.app.resume_button.state())
        self.assertIn("disabled", self.app.stop_button.state())
        self.assertIn("内存", self.app.results.get("1.0", "end"))
        self.assertIn("MemoryError", self.app.results.get("1.0", "end"))

        self.app.resume_pending()
        self.finish()
        self.assertEqual(self.app.succeeded, 1)
        self.assertEqual(self.app.pending_files, ())
        fresh = self.workbook("新批次-" + stage + ".xlsx", True)
        self.app.files = (str(fresh),)
        self.app.start()
        self.finish()
        self.assertEqual(self.app.succeeded, 1)
        self.assertEqual(tuple(self.app.file_states), (str(fresh),))
        self.assertEqual(callback_errors, [])
        # Exercise a real idle-window close, then leave teardown only the files.
        with patch.object(self.root, "destroy", wraps=self.root.destroy) as destroy:
            self.app.close()
            destroy.assert_called_once_with()
        self.root = None

    def test_thread_construction_memory_failure_can_resume_and_close(self):
        self.check_thread_memory_failure_recovery("excel_unmerge_gui.threading.Thread", "构造")

    def test_thread_start_memory_failure_can_resume_and_close(self):
        self.check_thread_memory_failure_recovery("excel_unmerge_gui.threading.Thread.start", "启动")

    def test_new_controls_fit_minimum_window_width(self):
        self.root.geometry("{}x{}".format(*self.root.minsize()))
        self.root.update()
        for button in (self.app.start_button, self.app.open_button,
                       self.app.retry_button, self.app.resume_button, self.app.stop_button):
            self.assertTrue(button.winfo_ismapped())
            self.assertLessEqual(button.winfo_rootx() + button.winfo_width(),
                                 self.root.winfo_rootx() + self.root.winfo_width())

    def test_long_status_wraps_to_available_width_after_resizing(self):
        wide_size = (self.root.winfo_width(), self.root.winfo_height())
        narrow_size = self.root.minsize()
        widths = []
        for size in (wide_size, narrow_size):
            self.root.geometry("{}x{}".format(*size))
            self.root.update()
            widths.append(self.app.status_label.winfo_width())
        self.assertGreater(widths[0], widths[1], "The test must use distinct window widths")
        # Physical sizes and font metrics vary with DPI. Choose text between
        # two narrow lines and two wide lines so narrowing must add a line.
        label_font = tkfont.nametofont("TkDefaultFont", root=self.root)
        message = "正在处理："
        suffix = ".xlsx（本轮第 1/3 个，已处理 0 个）。"
        while label_font.measure(message + suffix) < sum(widths) - 20:
            message += "明细"
        self.app.status.set(message + suffix)
        heights = []
        for size in (wide_size, narrow_size, wide_size):
            self.root.geometry("{}x{}".format(*size))
            self.root.update()
            label = self.app.status_label
            self.assertLessEqual(label.winfo_reqwidth(), label.winfo_width())
            self.assertLessEqual(label.winfo_reqheight(), label.winfo_height())
            self.assertLessEqual(label.winfo_rootx() + label.winfo_width(),
                                 self.root.winfo_rootx() + self.root.winfo_width())
            heights.append(label.winfo_height())
        self.assertGreater(heights[1], heights[0], "A narrower status should use more lines")
        self.assertEqual(heights[2], heights[0])


if __name__ == "__main__":
    unittest.main()
