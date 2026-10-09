"""Control only this acceptance run's real EXE via native Win32 messages."""
import ctypes
import json
import os
import subprocess
import sys
import time
from ctypes import wintypes as W
from pathlib import Path

import psutil
from PIL import Image

ROOT = Path(__file__).resolve().parent
EXE = None  # Set only after the runner verifies the requested release identity.
LAUNCHER = None
OWNED = {}  # PID -> creation time; path equality alone does not authorize a PID.
USER = ctypes.WinDLL('user32', use_last_error=True)
GDI = ctypes.WinDLL('gdi32', use_last_error=True)
CALLBACK = ctypes.WINFUNCTYPE(W.BOOL, W.HWND, W.LPARAM)
USER.EnumWindows.argtypes = [CALLBACK, W.LPARAM]
USER.EnumChildWindows.argtypes = [W.HWND, CALLBACK, W.LPARAM]
USER.GetWindowThreadProcessId.argtypes = [W.HWND, ctypes.POINTER(W.DWORD)]
USER.GetWindowTextW.argtypes = [W.HWND, W.LPWSTR, ctypes.c_int]
USER.GetClassNameW.argtypes = [W.HWND, W.LPWSTR, ctypes.c_int]
USER.GetWindowRect.argtypes = [W.HWND, ctypes.POINTER(W.RECT)]
USER.GetClientRect.argtypes = [W.HWND, ctypes.POINTER(W.RECT)]
USER.ShowWindow.argtypes = [W.HWND, ctypes.c_int]
USER.SetWindowPos.argtypes = [W.HWND, W.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, W.UINT]
USER.SendMessageW.argtypes = [W.HWND, W.UINT, W.WPARAM, W.LPARAM]
USER.SendMessageW.restype = ctypes.c_ssize_t
USER.PostMessageW.argtypes = [W.HWND, W.UINT, W.WPARAM, W.LPARAM]
USER.IsWindowVisible.argtypes = [W.HWND]
USER.IsWindowEnabled.argtypes = [W.HWND]
USER.GetDlgCtrlID.argtypes = [W.HWND]
USER.GetForegroundWindow.restype = W.HWND
USER.SetForegroundWindow.argtypes = [W.HWND]
USER.IsWindow.argtypes = [W.HWND]
USER.GetWindowLongW.argtypes = [W.HWND, ctypes.c_int]
USER.GetWindowLongW.restype = ctypes.c_long
USER.SetWindowLongW.argtypes = [W.HWND, ctypes.c_int, ctypes.c_long]
USER.SetLayeredWindowAttributes.argtypes = [W.HWND, W.DWORD, ctypes.c_ubyte, W.DWORD]
USER.GetCursorPos.argtypes = [ctypes.POINTER(W.POINT)]
USER.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
USER.WindowFromPoint.argtypes = [W.POINT]
USER.WindowFromPoint.restype = W.HWND
USER.GetDC.argtypes = [W.HWND]
USER.GetDC.restype = W.HDC
USER.ReleaseDC.argtypes = [W.HWND, W.HDC]
USER.PrintWindow.argtypes = [W.HWND, W.HDC, W.UINT]
USER.PrintWindow.restype = W.BOOL
GDI.CreateCompatibleDC.argtypes = [W.HDC]
GDI.CreateCompatibleDC.restype = W.HDC
GDI.CreateCompatibleBitmap.argtypes = [W.HDC, ctypes.c_int, ctypes.c_int]
GDI.CreateCompatibleBitmap.restype = W.HBITMAP
GDI.SelectObject.argtypes = [W.HDC, W.HANDLE]
GDI.SelectObject.restype = W.HANDLE
GDI.DeleteObject.argtypes = [W.HANDLE]
GDI.DeleteDC.argtypes = [W.HDC]
GDI.GetDIBits.argtypes = [W.HDC, W.HBITMAP, W.UINT, W.UINT, ctypes.c_void_p, ctypes.c_void_p, W.UINT]


class BitmapHeader(ctypes.Structure):
    _fields_ = [('size', W.DWORD), ('width', W.LONG), ('height', W.LONG),
               ('planes', W.WORD), ('bits', W.WORD), ('compression', W.DWORD),
               ('image_size', W.DWORD), ('x', W.LONG), ('y', W.LONG),
               ('used', W.DWORD), ('important', W.DWORD)]


def our_pids():
    result = set()
    for pid, created in list(OWNED.items()):
        try:
            proc = psutil.Process(pid)
            if proc.create_time() != created or Path(proc.exe()).resolve() != EXE.resolve():
                continue
            result.add(pid)
            for child in proc.children(recursive=True):
                if Path(child.exe()).resolve() == EXE.resolve():
                    OWNED[child.pid] = child.create_time()
                    result.add(child.pid)
        except (psutil.Error, OSError):
            pass
    return result


def details(hwnd):
    pid = W.DWORD()
    USER.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    title, klass = ctypes.create_unicode_buffer(8192), ctypes.create_unicode_buffer(256)
    USER.GetWindowTextW(hwnd, title, len(title))
    USER.GetClassNameW(hwnd, klass, len(klass))
    rect = W.RECT()
    USER.GetWindowRect(hwnd, ctypes.byref(rect))
    return dict(hwnd=int(hwnd), pid=pid.value, title=title.value, klass=klass.value,
                rect=[rect.left, rect.top, rect.right, rect.bottom],
                visible=bool(USER.IsWindowVisible(hwnd)), enabled=bool(USER.IsWindowEnabled(hwnd)),
                control_id=USER.GetDlgCtrlID(hwnd))


def windows():
    pids = our_pids()
    result = []
    @CALLBACK
    def collect(hwnd, _):
        info = details(hwnd)
        if info['pid'] in pids:
            result.append(info)
        return True
    USER.EnumWindows(collect, 0)
    return result


def children(hwnd):
    assert details(hwnd)['pid'] in our_pids()
    result = []
    @CALLBACK
    def collect(child, _):
        result.append(details(child))
        return True
    USER.EnumChildWindows(hwnd, collect, 0)
    return result


def main_window():
    candidates = [w for w in windows() if w['title'] == 'Excel 数据处理工具 v1.1']
    assert len(candidates) == 1, candidates
    return candidates[0]['hwnd']


def capture(hwnd, destination):
    assert details(hwnd)['pid'] in our_pids()
    rect = W.RECT()
    USER.GetWindowRect(hwnd, ctypes.byref(rect))
    width, height = rect.right-rect.left, rect.bottom-rect.top
    src = USER.GetDC(hwnd)
    dc = GDI.CreateCompatibleDC(src)
    bmp = GDI.CreateCompatibleBitmap(src, width, height)
    previous = GDI.SelectObject(dc, bmp)
    try:
        captured = bool(USER.PrintWindow(hwnd, dc, 2))
        header = BitmapHeader(ctypes.sizeof(BitmapHeader), width, -height, 1, 32, 0, 0, 0, 0, 0, 0)
        pixels = ctypes.create_string_buffer(width*height*4)
        rows = GDI.GetDIBits(dc, bmp, 0, height, pixels, ctypes.byref(header), 0)
        if not captured or rows != height:
            raise RuntimeError('PrintWindow did not produce a complete fresh image')
        Image.frombytes('RGB', (width, height), pixels.raw, 'raw', 'BGRX').save(destination)
        return dict(path=str(destination), captured=captured, rows=rows, width=width, height=height)
    finally:
        GDI.SelectObject(dc, previous)
        GDI.DeleteObject(bmp)
        GDI.DeleteDC(dc)
        USER.ReleaseDC(hwnd, src)


def click(hwnd, x=None, y=None):
    assert details(hwnd)['pid'] in our_pids()
    rect = W.RECT()
    USER.GetClientRect(hwnd, ctypes.byref(rect))
    x = x if x is not None else (rect.right-rect.left)//2
    y = y if y is not None else (rect.bottom-rect.top)//2
    position = (y << 16) | x
    USER.PostMessageW(hwnd, 0x0200, 0, position)
    USER.PostMessageW(hwnd, 0x0201, 1, position)
    USER.PostMessageW(hwnd, 0x0202, 0, position)


def keys(keys_to_send):
    """Hardware keyboard input is guarded to this EXE's foreground PID."""
    previous = USER.GetForegroundWindow()
    previous_is_ours = previous and details(previous)['pid'] in our_pids()
    assert USER.SetForegroundWindow(main_window()), 'Cannot activate own EXE'
    time.sleep(.1)
    try:
        for key in keys_to_send:
            assert details(USER.GetForegroundWindow())['pid'] in our_pids(), 'Foreground changed; no further keys sent'
            USER.keybd_event(key, 0, 0, 0)
            time.sleep(.04)
            USER.keybd_event(key, 0, 2, 0)
            time.sleep(.12)
    finally:
        restored = False
        if previous and not previous_is_ours and USER.IsWindow(previous):
            restored = bool(USER.SetForegroundWindow(previous))
        with (ROOT/'keyboard-actions.jsonl').open('a', encoding='utf-8') as log:
            log.write(json.dumps(dict(keys=keys_to_send, previous_hwnd=int(previous or 0),
                                      previous_was_ours=bool(previous_is_ours),
                                      external_foreground_restored=restored))+'\n')


def guarded_mouse_click(target):
    """Brief near-transparent own-window overlay; never click another PID."""
    assert details(target)['pid'] in our_pids()
    main = main_window()
    original_rect = details(main)['rect']
    original_style = USER.GetWindowLongW(main, -20)
    previous = USER.GetForegroundWindow()
    previous_is_ours = previous and details(previous)['pid'] in our_pids()
    old_cursor = W.POINT()
    USER.GetCursorPos(ctypes.byref(old_cursor))
    hit = None
    try:
        USER.SetWindowLongW(main, -20, original_style | 0x80000)
        USER.SetLayeredWindowAttributes(main, 0, 1, 2)
        USER.SetWindowPos(main, W.HWND(-1), 10, 10, 0, 0, 0x0011)
        time.sleep(.15)
        rect = details(target)['rect']
        point = W.POINT((rect[0]+rect[2])//2, (rect[1]+rect[3])//2)
        USER.SetCursorPos(point.x, point.y)
        time.sleep(.1)
        hit = USER.WindowFromPoint(point)
        assert hit == target and details(hit)['pid'] in our_pids(), 'Mouse target is not our requested control'
        USER.mouse_event(0x0002, 0, 0, 0, 0)
        time.sleep(.04)
        USER.mouse_event(0x0004, 0, 0, 0, 0)
        time.sleep(.3)
    finally:
        USER.SetWindowPos(main, W.HWND(-2), original_rect[0], original_rect[1], 0, 0, 0x0011)
        USER.SetLayeredWindowAttributes(main, 0, 255, 2)
        USER.SetWindowLongW(main, -20, original_style)
        current_cursor = W.POINT()
        USER.GetCursorPos(ctypes.byref(current_cursor))
        if 'point' in locals() and current_cursor.x == point.x and current_cursor.y == point.y:
            USER.SetCursorPos(old_cursor.x, old_cursor.y)
        restored = False
        if previous and not previous_is_ours and USER.IsWindow(previous):
            restored = bool(USER.SetForegroundWindow(previous))
        with (ROOT/'actual-input-actions.jsonl').open('a', encoding='utf-8') as log:
            log.write(json.dumps(dict(target=int(target), hit_hwnd=int(hit or 0),
                                      foreground_before=int(previous or 0),
                                      previous_was_ours=bool(previous_is_ours),
                                      external_foreground_restored=restored))+'\n')


def capture_fresh(destination):
    """Force this EXE's own offscreen widgets to repaint before capture."""
    main = main_window()
    original_rect = details(main)['rect']
    original_style = USER.GetWindowLongW(main, -20)
    try:
        USER.SetWindowLongW(main, -20, original_style | 0x80000)
        USER.SetLayeredWindowAttributes(main, 0, 1, 2)
        USER.SetWindowPos(main, W.HWND(-1), 10, 10, 0, 0, 0x0011)
        USER.RedrawWindow.argtypes = [W.HWND, ctypes.c_void_p, W.HANDLE, W.UINT]
        USER.RedrawWindow(main, None, None, 0x185)
        time.sleep(.4)
        return capture(main, destination)
    finally:
        USER.SetWindowPos(main, W.HWND(-2), original_rect[0], original_rect[1], 0, 0, 0x0011)
        USER.SetLayeredWindowAttributes(main, 0, 255, 2)
        USER.SetWindowLongW(main, -20, original_style)


def start(on_launched=None):
    global LAUNCHER
    assert EXE is not None and EXE.is_file(), 'A verified candidate is required'
    assert not our_pids(), 'Candidate EXE already running'
    startup = subprocess.STARTUPINFO()
    startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startup.wShowWindow = 0
    env = {key: value for key, value in os.environ.items()
           if not key.upper().startswith(('PYTHON', 'TCL', 'TK', 'CONDA'))
           and key.upper() != 'VIRTUAL_ENV'}
    env['PATH'] = os.environ['SystemRoot'] + r'\System32;' + os.environ['SystemRoot']
    process = subprocess.Popen([str(EXE)], cwd=EXE.parent, env=env,
                               startupinfo=startup, creationflags=subprocess.CREATE_NO_WINDOW)
    LAUNCHER = process
    OWNED[process.pid] = psutil.Process(process.pid).create_time()
    if on_launched is not None:
        on_launched(process)
    end = time.monotonic()+15
    while time.monotonic() < end:
        matches = [w for w in windows() if w['title'] == 'Excel 数据处理工具 v1.1']
        if matches:
            hwnd = matches[0]['hwnd']
            USER.SetWindowPos(hwnd, None, -20000, -20000, 0, 0, 0x0015)
            USER.ShowWindow(hwnd, 4)
            time.sleep(.5)
            info = dict(launcher_pid=process.pid, windows=windows(), main_children=children(hwnd),
                        capture=capture_fresh(ROOT/'exe-initial.png'))
            (ROOT/'gui-probe.json').write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding='utf-8')
            return info
        time.sleep(.1)
    raise TimeoutError('The candidate EXE did not create its window')


def close_idle(timeout=15):
    """Never kill a worker: request WM_CLOSE on our main window and await exit."""
    main = main_window()
    USER.PostMessageW(main, 0x0010, 0, 0)
    end = time.monotonic() + timeout
    while our_pids() and time.monotonic() < end:
        time.sleep(.1)
    remaining = sorted(our_pids())
    return dict(remaining_owned_pids=remaining,
                launcher_exit=LAUNCHER.poll() if LAUNCHER else None,
                forced_termination=False)


if __name__ == '__main__':
    if sys.argv[1] == 'start':
        start()
    elif sys.argv[1] == 'inspect':
        print(json.dumps(dict(windows=windows(), children=children(main_window())), ensure_ascii=False, indent=2))
