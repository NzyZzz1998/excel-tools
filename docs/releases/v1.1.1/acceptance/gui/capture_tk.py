"""Capture only this script's transparent Tk window with synthetic progress.

This is a GUI/layout probe, not a workbook-content or performance benchmark.
No user EXE, desktop capture, physical input, or business workbook is involved.
"""

import ctypes
from ctypes import wintypes as W
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import sys
import threading
import time
import tkinter as tk
from unittest.mock import patch


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[4]
sys.path.insert(0, str(REPO))
from excel_unmerge_gui import APP_VERSION, Application, enable_windows_dpi_awareness

helper = REPO / "docs/releases/v1.1/large-data-acceptance-2026-10-09/exe/native_gui_controller.py"
spec = importlib.util.spec_from_file_location("historical_capture_helper", helper)
capture_helper = importlib.util.module_from_spec(spec)
spec.loader.exec_module(capture_helper)
# Restrict the existing capture helper to our exact Python PID, never its EXE search.
capture_helper.our_pids = lambda: {os.getpid()}
user = capture_helper.USER
user.GetAncestor.argtypes = [W.HWND, W.UINT]
user.GetAncestor.restype = W.HWND
user.RedrawWindow.argtypes = [W.HWND, ctypes.c_void_p, W.HANDLE, W.UINT]


def main():
    enable_windows_dpi_awareness()
    root = tk.Tk()
    root.withdraw()
    root.attributes("-alpha", 0)
    app = Application(root)
    hwnd = user.GetAncestor(root.winfo_id(), 2)
    assert capture_helper.details(hwnd)["pid"] == os.getpid()
    user.SetWindowLongW(hwnd, -20, user.GetWindowLongW(hwnd, -20) | 0x08000000)
    root.geometry("+20+20")
    root.deiconify()
    root.update()
    hwnd = user.GetAncestor(root.winfo_id(), 2)
    assert capture_helper.details(hwnd)["pid"] == os.getpid()
    user.SetWindowLongW(hwnd, -20, user.GetWindowLongW(hwnd, -20) | 0x08000000)
    release = threading.Event()
    evidence = {"version": APP_VERSION, "pid": os.getpid(),
                "kind": "synthetic source Tk layout; not EXE or business acceptance",
                "gui_sha256": hashlib.sha256((REPO / "excel_unmerge_gui.py").read_bytes()).hexdigest(),
                "physical_input_used": False, "captures": []}

    def synthetic_process(file, all_merges, progress=None):
        progress(dict(phase="reading", sheet="合成测试表", completed=1200, total=None, unit="rows"))
        progress(dict(phase="filling", sheet="合成测试表", completed=1200, total=3000, unit="rows"))
        if not release.wait(15):
            raise AssertionError("Probe did not release its own worker")
        return None, []

    try:
        app.files = (r"C:\合成演示\待开始B.xlsx",)
        app.file_states = {r"C:\合成演示\上次任务A.xlsx": "failed"}
        app.update_counts()
        with patch("excel_unmerge_gui.process_file", side_effect=synthetic_process):
            app.retry_failed()
            deadline = time.monotonic() + 5
            while app.current_detail is None and time.monotonic() < deadline:
                root.update()
                time.sleep(.01)
            assert app.current_detail["phase"] == "filling"
            deadline = time.monotonic() + 3
            while "已用时 00:00" in app.status.get() and time.monotonic() < deadline:
                root.update()
                time.sleep(.01)
            assert "已用时 00:00" not in app.status.get()
            for label, size in (("default", (root.winfo_width(), root.winfo_height())),
                                ("minimum", root.minsize())):
                root.geometry("{}x{}+20+20".format(*size))
                root.update()
                user.RedrawWindow(hwnd, None, None, 0x0001 | 0x0080 | 0x0100 | 0x0400)
                root.update()
                controls = {}
                for name in ("choose_button", "file_list", "option", "start_button", "open_button",
                             "retry_button", "resume_button", "stop_button", "progress",
                             "status_label", "results"):
                    widget = getattr(app, name)
                    bounds = [widget.winfo_rootx() - root.winfo_rootx(),
                              widget.winfo_rooty() - root.winfo_rooty(),
                              widget.winfo_width(), widget.winfo_height()]
                    assert widget.winfo_ismapped() and all(value > 0 for value in bounds[2:])
                    assert bounds[0] >= 0 and bounds[0] + bounds[2] <= root.winfo_width()
                    assert bounds[1] >= 0 and bounds[1] + bounds[3] <= root.winfo_height()
                    if name.endswith("button"):
                        assert widget.winfo_width() >= widget.winfo_reqwidth()
                    controls[name] = bounds
                screenshot = capture_helper.capture(hwnd, HERE / ("ui-" + label + ".png"))
                evidence["captures"].append(dict(size_name=label, title=root.title(),
                    client_size=list(size), status=app.status.get(),
                    shown_files=app.file_list.get("1.0", "end-1c"), next_selection=app.files,
                    file_progress=[app.progress["value"], app.progress["maximum"]],
                    controls=controls, screenshot=screenshot))
            release.set()
            deadline = time.monotonic() + 5
            while app.running and time.monotonic() < deadline:
                root.update()
                time.sleep(.01)
            assert not app.running and app.files == (r"C:\合成演示\待开始B.xlsx",)
            assert app.file_list.get("1.0", "end-1c") == app.files[0]
            evidence["next_selection_restored"] = True
            evidence["idle_status"] = app.status.get()
    finally:
        release.set()
        root.destroy()
    evidence["own_window_closed"] = True
    (HERE / "layout-evidence.json").write_text(json.dumps(evidence, ensure_ascii=False, indent=2),
                                               encoding="utf-8")
    print(json.dumps(dict(passed=True, own_pid=os.getpid(), screenshots=2)))


if __name__ == "__main__":
    main()
