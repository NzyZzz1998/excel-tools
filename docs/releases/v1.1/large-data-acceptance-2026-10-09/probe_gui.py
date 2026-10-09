"""Bounded GUI/GIL and memory-failure probes; no large or business workbooks."""
import hashlib
import json
import statistics
import sys
import tempfile
import time
import tkinter as tk
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook

REPO = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(REPO))
import excel_unmerge_gui as gui
from excel_unmerge_fill import process_file


def workbook(folder, name):
    path = folder / name
    book = Workbook()
    book.active["A1"] = "synthetic"
    book.active.merge_cells("A1:A3")
    book.save(path)
    book.close()
    return str(path)


def window():
    root = tk.Tk()
    root.attributes("-alpha", 0)
    app = gui.Application(root)
    root.update()
    errors = []
    root.report_callback_exception = lambda kind, value, traceback: errors.append(kind.__name__)
    return root, app, errors


def finish(root, app):
    deadline = time.perf_counter() + 5
    while app.running and time.perf_counter() < deadline:
        root.update()
        time.sleep(.005)
    assert not app.running


def cpu_probe(folder):
    root, app, errors = window()
    first = workbook(folder, "cpu-current.xlsx")
    later = workbook(folder, "cpu-later.xlsx")
    app.files = (first, later)
    timestamps = []
    events = {}
    started = time.perf_counter()

    def busy(file, all_merges):
        end = time.perf_counter() + 2
        value = 0
        while time.perf_counter() < end:
            value = (value + 1) % 1000003
        result = process_file(file, all_merges=all_merges)
        events["current_file_finished_seconds"] = time.perf_counter() - started
        return result

    def heartbeat():
        timestamps.append(time.perf_counter())
        if app.running:
            root.after(20, heartbeat)
        else:
            root.quit()

    def stop():
        app.request_stop()
        events["stop_request_executed_seconds"] = time.perf_counter() - started

    def close_again():
        app.close()
        events["close_request_executed_seconds"] = time.perf_counter() - started
        events["window_alive_after_close_request"] = bool(root.winfo_exists())

    try:
        with patch.object(gui, "process_file", side_effect=busy):
            app.start()
            root.after(20, heartbeat)
            root.after(200, stop)
            root.after(350, close_again)
            root.after(8000, root.quit)
            root.mainloop()
        assert not app.running
        assert app.succeeded == 1 and app.pending_files == (later,)
        gaps = [(b - a) * 1000 for a, b in zip(timestamps, timestamps[1:])]
        return {"workload": "2 seconds of bounded pure-Python arithmetic, then one tiny synthetic workbook",
                "heartbeat_target_ms": 20, "heartbeat_samples": len(timestamps),
                "max_heartbeat_gap_ms": max(gaps), "median_heartbeat_gap_ms": statistics.median(gaps),
                "events": events, "callback_errors": errors,
                "current_file_saved": Path(first).with_name("cpu-current_拆分填充.xlsx").exists(),
                "next_file_not_started": not Path(later).with_name("cpu-later_拆分填充.xlsx").exists(),
                "pending_count": len(app.pending_files), "passed": not errors}
    finally:
        root.destroy()


def worker_memory_probe(folder):
    root, app, errors = window()
    first = workbook(folder, "memory-failed.xlsx")
    later = workbook(folder, "memory-success.xlsx")
    app.files = (first, later)
    try:
        def injected(file, all_merges):
            if file == first:
                raise MemoryError()
            return process_file(file, all_merges=all_merges)
        with patch.object(gui, "process_file", side_effect=injected):
            app.start()
            finish(root, app)
        initial = {"failed": app.failed, "succeeded": app.succeeded,
                   "failed_item_retryable": app.failed_files == (first,),
                   "contains_memory_guidance": "内存" in app.results.get("1.0", "end")}
        app.retry_failed()
        finish(root, app)
        assert app.succeeded == 2 and app.failed == 0
        return {"initial": initial, "retry_succeeded": True,
                "success_not_reprocessed": len(list(folder.glob("memory-success_拆分填充*.xlsx"))) == 1,
                "callback_errors": errors, "passed": not errors}
    finally:
        root.destroy()


def thread_memory_probe(folder, stage):
    root, app, errors = window()
    source = workbook(folder, "thread-memory-" + stage + ".xlsx")
    app.files = (source,)
    target = "excel_unmerge_gui.threading.Thread" if stage == "construct" else "excel_unmerge_gui.threading.Thread.start"
    try:
        with patch(target, side_effect=MemoryError("synthetic thread allocation failure")):
            root.after(0, app.start)
            root.update()
        app.close()
        app.retry_failed()
        app.resume_pending()
        app.start()
        root.update()
        return {"injected_stage": stage, "callback_errors": errors,
                "running_remains_true": app.running,
                "queue_empty_without_done": app.events.empty(),
                "window_still_exists_after_close": bool(root.winfo_exists()),
                "start_disabled": "disabled" in app.start_button.state(),
                "retry_disabled": "disabled" in app.retry_button.state(),
                "resume_disabled": "disabled" in app.resume_button.state(),
                "no_output_generated": not list(folder.glob("thread-memory-" + stage + "_拆分填充*.xlsx")),
                "stuck_state_confirmed": app.running and app.events.empty() and "MemoryError" in errors}
    finally:
        # No worker was created or started; destroy only this probe's own window.
        root.destroy()


if __name__ == "__main__":
    source = REPO / "excel_unmerge_gui.py"
    before = hashlib.sha256(source.read_bytes()).hexdigest()
    with tempfile.TemporaryDirectory(prefix="excel-large-gui-probes-") as directory:
        folder = Path(directory)
        report = {"gui_sha256": before, "cpu_bound_worker": cpu_probe(folder),
                  "worker_memory_error": worker_memory_probe(folder),
                  "thread_allocation_memory_errors": [thread_memory_probe(folder, stage)
                                                       for stage in ("construct", "start")],
                  "limitations": "Bounded synthetic probes only. Does not measure multi-GB RSS, long C-extension/GIL or GC pauses, OS OOM termination, or real large-workbook processing."}
    report["gui_source_unchanged"] = hashlib.sha256(source.read_bytes()).hexdigest() == before
    destination = Path(__file__).with_name("gui-probe-results.json")
    assert not destination.exists(), "Preserve previous probe evidence"
    destination.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
