"""Own-window source probe using the same DPI initialization as the EXE.

All rows, paths and errors below are synthetic UI state. No workbook is read.
PrintWindow captures only the Toplevel belonging to this process, never desktop.
Run with the release build's Python; uses only standard-library dependencies.
"""

import ctypes
from ctypes import wintypes as W
import hashlib
import json
import os
from pathlib import Path
import struct
import sys
import tkinter as tk
import zlib

REPO = Path(__file__).resolve().parents[5]
sys.path.insert(0, str(REPO))
from excel_unmerge_gui import Application, enable_windows_dpi_awareness
from excel_table_tools import TableSource


def screenshot(window, path):
    user, gdi = ctypes.WinDLL("user32"), ctypes.WinDLL("gdi32")
    user.GetAncestor.argtypes = [W.HWND, W.UINT]
    user.GetAncestor.restype = W.HWND
    hwnd = user.GetAncestor(window.winfo_id(), 2)
    user.GetWindowThreadProcessId.argtypes = [W.HWND, ctypes.POINTER(W.DWORD)]
    pid = W.DWORD()
    user.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    assert pid.value == os.getpid(), "Refuse another process's window"
    user.GetWindowRect.argtypes = [W.HWND, ctypes.POINTER(W.RECT)]
    rect = W.RECT()
    assert user.GetWindowRect(hwnd, ctypes.byref(rect))
    width, height = rect.right - rect.left, rect.bottom - rect.top
    user.GetDC.argtypes, user.GetDC.restype = [W.HWND], W.HDC
    user.ReleaseDC.argtypes = [W.HWND, W.HDC]
    user.PrintWindow.argtypes = [W.HWND, W.HDC, W.UINT]
    gdi.CreateCompatibleDC.argtypes, gdi.CreateCompatibleDC.restype = [W.HDC], W.HDC
    gdi.CreateCompatibleBitmap.argtypes = [W.HDC, ctypes.c_int, ctypes.c_int]
    gdi.CreateCompatibleBitmap.restype = W.HBITMAP
    gdi.SelectObject.argtypes, gdi.SelectObject.restype = [W.HDC, W.HANDLE], W.HANDLE
    gdi.GetDIBits.argtypes = [W.HDC, W.HBITMAP, W.UINT, W.UINT, ctypes.c_void_p, ctypes.c_void_p, W.UINT]
    gdi.DeleteObject.argtypes, gdi.DeleteDC.argtypes = [W.HANDLE], [W.HDC]
    src = user.GetDC(hwnd)
    dc = gdi.CreateCompatibleDC(src)
    bitmap = gdi.CreateCompatibleBitmap(src, width, height)
    previous = gdi.SelectObject(dc, bitmap)
    try:
        assert user.PrintWindow(hwnd, dc, 2), "PrintWindow failed"
        header = ctypes.create_string_buffer(struct.pack("<IiiHHIIiiII", 40, width, -height, 1, 32, 0, 0, 0, 0, 0, 0))
        pixels = ctypes.create_string_buffer(width * height * 4)
        assert gdi.GetDIBits(dc, bitmap, 0, height, pixels, header, 0) == height
        rgb = bytearray(width * height * 3)
        rgb[0::3], rgb[1::3], rgb[2::3] = pixels.raw[2::4], pixels.raw[1::4], pixels.raw[0::4]
        rows = b"".join(b"\0" + rgb[y * width * 3:(y + 1) * width * 3] for y in range(height))

        def chunk(kind, data):
            return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

        path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
                         + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))
        return dict(path=path.name, pid=pid.value, size=[width, height], fresh=True)
    finally:
        gdi.SelectObject(dc, previous)
        gdi.DeleteObject(bitmap)
        gdi.DeleteDC(dc)
        user.ReleaseDC(hwnd, src)


def main():
    enable_windows_dpi_awareness()
    user = ctypes.WinDLL("user32")
    user.GetThreadDpiAwarenessContext.restype = ctypes.c_void_p
    user.GetAwarenessFromDpiAwarenessContext.argtypes = [ctypes.c_void_p]
    root = tk.Tk()
    root.attributes("-alpha", 0)
    app = Application(root)
    app.open_table_tools()
    tool = app.table_tools
    tool.window.attributes("-alpha", 0)
    result = dict(scope="Source, real Windows DPI; synthetic long paths and errors, no workbook processing",
                  executable=sys.executable, python=sys.version, pid=os.getpid(),
                  os_dpi=user.GetDpiForSystem(),
                  awareness=user.GetAwarenessFromDpiAwarenessContext(user.GetThreadDpiAwarenessContext()),
                  tk_dpi=root.winfo_fpixels("1i"), tk_scaling=root.tk.call("tk", "scaling"),
                  tk_version=root.tk.call("info", "patchlevel"),
                  screen=[root.winfo_screenwidth(), root.winfo_screenheight()],
                  hashes={name: hashlib.sha256((REPO / name).read_bytes()).hexdigest()
                          for name in ("excel_unmerge_gui.py", "excel_table_gui.py")}, cases=[])
    try:
        prefix = Path("C:/合成测试_不存在的目录") / ("仅用于界面边界验证的长路径" * 7)
        tool.sources = [TableSource(prefix / ("合成来源_{:02d}.xlsx".format(i)), "合成工作表", 1) for i in range(12)]
        tool._sources_changed()
        tool.file.set(str(tool.sources[-1].path))
        for i in range(8):
            tool._append_log("合成多行错误 {}\n无法访问文件或目录，请检查后重试。\n技术详情：PermissionError: {}\n源文件未更改，未生成本次正式结果。".format(i, tool.sources[i].path))
        tool.status.set("处理失败：未生成本次正式结果。详情见处理记录。此前成功的结果仍可通过最近成功结果入口打开。")
        for size in ("900x760", "760x700"):
            tool.window.geometry(size)
            root.update()
            width, height = tool.window.winfo_width(), tool.window.winfo_height()
            widgets = {}
            for name in ("choose_button", "add_button", "remove_button", "output_button", "preview_button",
                         "run_button", "stop_button", "open_button", "log", "status_label"):
                widget = getattr(tool, name)
                x, y = widget.winfo_rootx() - tool.window.winfo_rootx(), widget.winfo_rooty() - tool.window.winfo_rooty()
                w, h = widget.winfo_width(), widget.winfo_height()
                fits = x >= 0 and y >= 0 and x + w <= width and y + h <= height
                widgets[name] = dict(rect=[x, y, w, h], mapped=bool(widget.winfo_ismapped()), fits=fits)
                assert fits and widget.winfo_ismapped(), (size, name, widgets[name])
            assert tool.log.winfo_height() >= 40, "Error log must retain readable height"
            assert tool.log.yview()[0] > 0, "Long error details must be scrollable"
            assert len(tool.source_list.get_children()) == 12
            assert tool.window.winfo_rootx() >= 0 and tool.window.winfo_rooty() >= 0
            assert tool.window.winfo_rootx() + width <= root.winfo_screenwidth()
            assert tool.window.winfo_rooty() + height <= root.winfo_screenheight()
            result["cases"].append(dict(requested=size, actual=[width, height], widgets=widgets,
                                       log_scroll=tool.log.yview(), source_scroll=tool.source_list.yview(),
                                       screenshot=screenshot(tool.window, Path(__file__).parent / ("dpi-" + size + ".png"))))
        result["passed"] = True
    finally:
        app.close()
    output = Path(__file__).with_name("dpi-validation.json")
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps(dict(passed=result["passed"], dpi=result["os_dpi"], screen=result["screen"], evidence=str(output))))


if __name__ == "__main__":
    main()
