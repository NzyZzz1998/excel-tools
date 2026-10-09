import queue
import errno
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch
from zipfile import BadZipFile

from test_excel_unmerge import c, load, makebook, sheet
from excel_unmerge_gui import describe_error, main, process_batch


class GuiBatchTests(unittest.TestCase):
    def test_bad_file_does_not_stop_later_file(self):
        with tempfile.TemporaryDirectory() as folder:
            source = makebook(Path(folder) / "中文 路径.xlsx", [
                ("数据", sheet({1: [c("A1", "华东仓")]}, ["A1:A3"]))])
            missing = Path(folder) / "不存在.xlsx"
            events = queue.Queue()
            process_batch([str(missing), str(source)], False, events)
            self.assertEqual(events.get_nowait(), ("started", str(missing), 1, 2))
            failure = events.get_nowait()
            self.assertEqual(events.get_nowait(), ("started", str(source), 2, 2))
            success = events.get_nowait()
            done = events.get_nowait()
            self.assertEqual(failure[0], "result")
            self.assertTrue(failure[4])
            self.assertEqual(success[0], "result")
            self.assertIsNone(success[4])
            with load(success[2]) as book:
                self.assertEqual(book.active["A3"].value, "华东仓")
            self.assertEqual(done, ("done", 2, 1, ()))
            self.assertTrue(events.empty())

    def test_all_merges_and_no_match_results(self):
        with tempfile.TemporaryDirectory() as folder:
            source = makebook(Path(folder) / "横向.xlsx", [
                ("数据", sheet({1: [c("A1", "标题")]}, ["A1:C1"]))])
            no_match = queue.Queue()
            process_batch([str(source)], False, no_match)
            self.assertEqual(no_match.get_nowait(), ("started", str(source), 1, 1))
            result = no_match.get_nowait()
            self.assertIsNone(result[2])
            self.assertIsNone(result[4])
            self.assertEqual(no_match.get_nowait(), ("done", 1, 0, ()))

            events = queue.Queue()
            process_batch([str(source)], True, events)
            self.assertEqual(events.get_nowait(), ("started", str(source), 1, 1))
            result = events.get_nowait()
            self.assertIsNone(result[4])
            with load(result[2]) as book:
                self.assertEqual(book.active["C1"].value, "标题")
            self.assertEqual(events.get_nowait(), ("done", 1, 0, ()))

    def test_stop_waits_for_current_file_and_leaves_later_files_unprocessed(self):
        events = queue.Queue()
        stop = threading.Event()
        entered = threading.Event()
        release = threading.Event()

        def process(file, all_merges):
            entered.set()
            if not release.wait(5):
                raise AssertionError("Test did not release the current file")
            return Path("output.xlsx"), [("数据", 1, 2)]

        with patch("excel_unmerge_gui.process_file", side_effect=process) as mocked:
            worker = threading.Thread(target=process_batch,
                                      args=(("first.xlsx", "later.xlsx"), False, events, stop))
            worker.start()
            try:
                self.assertTrue(entered.wait(5))
                self.assertEqual(events.get_nowait(), ("started", "first.xlsx", 1, 2))
                stop.set()
                self.assertTrue(worker.is_alive())
                self.assertTrue(events.empty(), "A saving file must not be reported as stopped")
            finally:
                release.set()
                worker.join(5)
            self.assertFalse(worker.is_alive())
            self.assertEqual(mocked.call_count, 1)
            self.assertEqual(events.get_nowait()[0], "result")
            self.assertEqual(events.get_nowait(), ("done", 1, 0, ("later.xlsx",)))

    def test_stop_before_first_file_reports_no_completion(self):
        stop = threading.Event()
        stop.set()
        events = queue.Queue()
        with patch("excel_unmerge_gui.process_file") as mocked:
            process_batch(("first.xlsx",), False, events, stop)
        mocked.assert_not_called()
        self.assertEqual(events.get_nowait(), ("done", 0, 0, ("first.xlsx",)))
        self.assertTrue(events.empty())

    def test_common_failures_include_action_and_original_technical_details(self):
        cases = (
            (BadZipFile("File is not a zip file"), "加密"),
            (PermissionError("access denied"), "可写目录"),
            (OSError(errno.ENOSPC, "No space left on device"), "空间"),
            (MemoryError(), "内存"),
        )
        for error, action in cases:
            with self.subTest(error=type(error).__name__):
                message = describe_error(error)
                self.assertIn(action, message)
                self.assertIn("技术详情：", message)
                self.assertIn(type(error).__name__, message)
                self.assertIn(str(error), message)

    def test_self_test_failure_writes_diagnostic_without_stderr(self):
        with tempfile.TemporaryDirectory() as folder:
            log = Path(folder) / "中文 日志" / "self-test.log"
            with patch("sys.argv", ["ExcelTools", "--self-test", "--self-test-log", str(log)]), \
                    patch("sys.stderr", None), \
                    patch("excel_unmerge_gui.self_test", side_effect=RuntimeError("诊断测试")):
                result = main()
            self.assertEqual(result, 1)
            diagnostic = log.read_text(encoding="utf-8")
            self.assertIn("Traceback", diagnostic)
            self.assertIn("RuntimeError: 诊断测试", diagnostic)


if __name__ == "__main__":
    unittest.main()
