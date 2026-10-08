import queue
import tempfile
import unittest
from pathlib import Path

from test_excel_unmerge import c, load, makebook, sheet
from excel_unmerge_gui import process_batch


class GuiBatchTests(unittest.TestCase):
    def test_bad_file_does_not_stop_later_file(self):
        with tempfile.TemporaryDirectory() as folder:
            source = makebook(Path(folder) / "中文 路径.xlsx", [
                ("数据", sheet({1: [c("A1", "华东仓")]}, ["A1:A3"]))])
            missing = Path(folder) / "不存在.xlsx"
            events = queue.Queue()
            process_batch([str(missing), str(source)], False, events)
            failure = events.get_nowait()
            success = events.get_nowait()
            done = events.get_nowait()
            self.assertEqual(failure[0], "result")
            self.assertTrue(failure[4])
            self.assertEqual(success[0], "result")
            self.assertIsNone(success[4])
            with load(success[2]) as book:
                self.assertEqual(book.active["A3"].value, "华东仓")
            self.assertEqual(done, ("done", 2, 1))
            self.assertTrue(events.empty())

    def test_all_merges_and_no_match_results(self):
        with tempfile.TemporaryDirectory() as folder:
            source = makebook(Path(folder) / "横向.xlsx", [
                ("数据", sheet({1: [c("A1", "标题")]}, ["A1:C1"]))])
            no_match = queue.Queue()
            process_batch([str(source)], False, no_match)
            result = no_match.get_nowait()
            self.assertIsNone(result[2])
            self.assertIsNone(result[4])
            self.assertEqual(no_match.get_nowait(), ("done", 1, 0))

            events = queue.Queue()
            process_batch([str(source)], True, events)
            result = events.get_nowait()
            self.assertIsNone(result[4])
            with load(result[2]) as book:
                self.assertEqual(book.active["C1"].value, "标题")
            self.assertEqual(events.get_nowait(), ("done", 1, 0))


if __name__ == "__main__":
    unittest.main()
