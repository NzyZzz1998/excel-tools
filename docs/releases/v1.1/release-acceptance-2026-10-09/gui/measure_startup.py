"""Time three hidden launches of the verified EXE; interact only with owned PIDs."""
import ctypes
import hashlib
import json
import os
import platform
import statistics
import subprocess
import time
from ctypes import wintypes as W
from datetime import datetime, timezone
from pathlib import Path

EXE = Path(r"E:\codex\excel-tools\testfile\验收_v1.1_2026-10-09\fixed\exe\app\ExcelTools.exe")
EXPECTED_SHA = "066ca4a1c7fe7fd2ac80a06ddc3969dbd178f5181983d8767d2c219078485d5c"
OUTPUT = Path(__file__).with_name("startup-timing.json")
USER = ctypes.WinDLL("user32", use_last_error=True)
KERNEL = ctypes.WinDLL("kernel32", use_last_error=True)
CALLBACK = ctypes.WINFUNCTYPE(W.BOOL, W.HWND, W.LPARAM)


class ProcessEntry(ctypes.Structure):
    _fields_ = [("dwSize", W.DWORD), ("cntUsage", W.DWORD), ("pid", W.DWORD),
                ("heap", ctypes.c_size_t), ("module", W.DWORD), ("threads", W.DWORD),
                ("parent_pid", W.DWORD), ("priority", W.LONG), ("flags", W.DWORD),
                ("exe_name", W.WCHAR * 260)]


KERNEL.CreateToolhelp32Snapshot.argtypes = [W.DWORD, W.DWORD]
KERNEL.CreateToolhelp32Snapshot.restype = W.HANDLE
KERNEL.Process32FirstW.argtypes = [W.HANDLE, ctypes.POINTER(ProcessEntry)]
KERNEL.Process32NextW.argtypes = [W.HANDLE, ctypes.POINTER(ProcessEntry)]
KERNEL.CloseHandle.argtypes = [W.HANDLE]
KERNEL.OpenProcess.argtypes = [W.DWORD, W.BOOL, W.DWORD]
KERNEL.OpenProcess.restype = W.HANDLE
KERNEL.QueryFullProcessImageNameW.argtypes = [W.HANDLE, W.DWORD, W.LPWSTR, ctypes.POINTER(W.DWORD)]
KERNEL.GetExitCodeProcess.argtypes = [W.HANDLE, ctypes.POINTER(W.DWORD)]
KERNEL.TerminateProcess.argtypes = [W.HANDLE, W.UINT]
KERNEL.WaitForSingleObject.argtypes = [W.HANDLE, W.DWORD]
USER.EnumWindows.argtypes = [CALLBACK, W.LPARAM]
USER.EnumChildWindows.argtypes = [W.HWND, CALLBACK, W.LPARAM]
USER.GetWindowThreadProcessId.argtypes = [W.HWND, ctypes.POINTER(W.DWORD)]
USER.GetWindowTextW.argtypes = [W.HWND, W.LPWSTR, ctypes.c_int]
USER.ShowWindow.argtypes = [W.HWND, ctypes.c_int]
USER.GetWindowLongPtrW.argtypes = [W.HWND, ctypes.c_int]
USER.GetWindowLongPtrW.restype = ctypes.c_ssize_t
USER.SetWindowLongPtrW.argtypes = [W.HWND, ctypes.c_int, ctypes.c_ssize_t]
USER.SetWindowLongPtrW.restype = ctypes.c_ssize_t
USER.SetLayeredWindowAttributes.argtypes = [W.HWND, W.DWORD, ctypes.c_ubyte, W.DWORD]
USER.SetWindowPos.argtypes = [W.HWND, W.HWND, ctypes.c_int, ctypes.c_int,
                              ctypes.c_int, ctypes.c_int, W.UINT]
USER.PostMessageW.argtypes = [W.HWND, W.UINT, W.WPARAM, W.LPARAM]
USER.SendMessageTimeoutW.argtypes = [W.HWND, W.UINT, W.WPARAM, W.LPARAM,
                                     W.UINT, W.UINT, ctypes.POINTER(ctypes.c_size_t)]
USER.SendMessageTimeoutW.restype = ctypes.c_ssize_t


def process_parents():
    snapshot = KERNEL.CreateToolhelp32Snapshot(2, 0)
    if snapshot == ctypes.c_void_p(-1).value:
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        entry = ProcessEntry()
        entry.dwSize = ctypes.sizeof(entry)
        available = KERNEL.Process32FirstW(snapshot, ctypes.byref(entry))
        result = {}
        while available:
            result[entry.pid] = entry.parent_pid
            available = KERNEL.Process32NextW(snapshot, ctypes.byref(entry))
        return result
    finally:
        KERNEL.CloseHandle(snapshot)


def update_owned_handles(launcher_pid, handles):
    parents = process_parents()
    owned = {launcher_pid} | set(handles)
    previous = set()
    while previous != owned:
        previous = set(owned)
        owned.update(pid for pid, parent in parents.items() if parent in owned)
    for pid in owned - set(handles):
        handle = KERNEL.OpenProcess(0x1000 | 0x100000 | 1, False, pid)
        if not handle:
            continue
        image_path = ctypes.create_unicode_buffer(32768)
        size = W.DWORD(len(image_path))
        if (KERNEL.QueryFullProcessImageNameW(handle, 0, image_path, ctypes.byref(size))
                and Path(image_path.value).resolve() == EXE.resolve()):
            handles[pid] = handle
        else:
            KERNEL.CloseHandle(handle)


def owned_windows(handles):
    result = []
    @CALLBACK
    def collect(hwnd, unused):
        pid = W.DWORD()
        USER.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value in handles:
            title = ctypes.create_unicode_buffer(256)
            USER.GetWindowTextW(hwnd, title, len(title))
            if title.value == "Excel 数据处理工具 v1.1":
                result.append((hwnd, pid.value))
        return True
    USER.EnumWindows(collect, 0)
    return result


def child_count(hwnd):
    result = []
    @CALLBACK
    def collect(child, unused):
        result.append(child)
        return True
    USER.EnumChildWindows(hwnd, collect, 0)
    return len(result)


def still_active(handle):
    status = W.DWORD()
    return bool(KERNEL.GetExitCodeProcess(handle, ctypes.byref(status))) and status.value == 259


def sample(index):
    startup = subprocess.STARTUPINFO()
    startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = 0  # SW_HIDE; no interactive window was requested.
    env = {key: value for key, value in os.environ.items()
           if not key.upper().startswith(("PYTHON", "TCL", "TK", "CONDA"))
           and key.upper() != "VIRTUAL_ENV"}
    env["PATH"] = os.environ["SystemRoot"] + "\\System32;" + os.environ["SystemRoot"]
    started = time.perf_counter()
    process = subprocess.Popen([str(EXE)], cwd=EXE.parent, env=env,
                               startupinfo=startup, creationflags=subprocess.CREATE_NO_WINDOW)
    handles = {}
    observations = {"run": index, "launcher_pid": process.pid,
                    "process_creation_seconds": time.perf_counter() - started}
    try:
        deadline = started + 20
        window_first_seen = None
        mapped = set()
        while time.perf_counter() < deadline:
            update_owned_handles(process.pid, handles)
            for hwnd, gui_pid in owned_windows(handles):
                if window_first_seen is None:
                    window_first_seen = time.perf_counter() - started
                # Restrict all mutations to a descendant whose executable matches EXE.
                if hwnd not in mapped:
                    observations["native_child_controls_before_offscreen_map"] = child_count(hwnd)
                    USER.SetWindowLongPtrW(hwnd, -20, USER.GetWindowLongPtrW(hwnd, -20) | 0x80000 | 0x08000000)
                    USER.SetLayeredWindowAttributes(hwnd, 0, 0, 2)
                    USER.SetWindowPos(hwnd, None, -20000, -20000, 0, 0, 0x0015)
                    USER.ShowWindow(hwnd, 4)  # Map transparently offscreen without activation.
                    mapped.add(hwnd)
                controls = child_count(hwnd)
                if controls >= 26:
                    response = ctypes.c_size_t()
                    response_started = time.perf_counter()
                    answered = USER.SendMessageTimeoutW(hwnd, 0, 0, 0, 0x0003, 500,
                                                       ctypes.byref(response))
                    if answered:
                        observations.update({"gui_pid": gui_pid,
                            "window_first_seen_seconds": window_first_seen,
                            "window_and_controls_responsive_seconds": time.perf_counter() - started,
                            "wm_null_response_seconds": time.perf_counter() - response_started,
                            "native_child_control_count": controls,
                            "responsive": True})
                        return observations
            if process.poll() is not None:
                raise RuntimeError("Owned launcher exited before a responsive GUI was found")
            time.sleep(.01)
        raise TimeoutError("Owned EXE did not become responsive within 20 seconds")
    finally:
        update_owned_handles(process.pid, handles)
        for hwnd, unused_pid in owned_windows(handles):
            USER.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE only to our idle application.
        forced = []
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            for pid, handle in reversed(list(handles.items())):
                if still_active(handle):
                    KERNEL.TerminateProcess(handle, 1)
                    KERNEL.WaitForSingleObject(handle, 2000)
                    forced.append(pid)
            process.wait(timeout=5)
        observations.update({"launcher_exit_code": process.returncode,
                             "forced_cleanup_pids": forced,
                             "owned_processes_still_active": [pid for pid, h in handles.items() if still_active(h)]})
        for handle in handles.values():
            KERNEL.CloseHandle(handle)


if __name__ == "__main__":
    if OUTPUT.exists():
        raise RuntimeError("Startup evidence already exists; choose a fresh output for another run")
    sha = hashlib.sha256(EXE.read_bytes()).hexdigest()
    assert sha == EXPECTED_SHA, sha
    runs = []
    for index in range(1, 4):
        runs.append(sample(index))
        time.sleep(.25)
    times = [run["window_and_controls_responsive_seconds"] for run in runs]
    report = {"captured_utc": datetime.now(timezone.utc).isoformat(),
              "exe": str(EXE), "exe_sha256": sha, "app_version": "1.1",
              "platform": platform.platform(), "runs": runs,
              "median_responsive_seconds": statistics.median(times),
              "min_responsive_seconds": min(times), "max_responsive_seconds": max(times),
              "measurement": "perf_counter before Popen through owned GUI title, at least 26 native child controls, and successful WM_NULL SendMessageTimeout response; 10ms polling",
              "isolation": "STARTUPINFO SW_HIDE, CREATE_NO_WINDOW, exact-image verified launcher descendants only; own window maps fully transparent at (-20000,-20000) with NOACTIVATE to let Tk create controls; WM_NULL/WM_CLOSE only; no keyboard/mouse input",
              "instrumentation_note": "A premeasurement probe that repeatedly hid the root timed out on the 26-control criterion; Tk needed an offscreen nonactivating map. This was not counted as an application startup duration.",
              "cache_caveat": "Sequential first and subsequent launches only. No reboot or OS cache clearing; prior acceptance and the prelaunch SHA read can warm file caches. No true cold-start comparison or service guarantee.",
              "v1_0_comparison": "No directly available v1.0 executable/package was found in the project EXE/dist inventory; no download or old-version speed claim.",
              "scope": "Startup only. No workbook processing, Excel opening, user windows, or application source edits."}
    OUTPUT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
