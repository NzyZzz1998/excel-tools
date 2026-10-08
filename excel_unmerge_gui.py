#!/usr/bin/env python3
"""Excel 合并行拆分的图形入口；也提供打包后的离线自检。"""

import argparse
import os
import queue
import subprocess
import sys
import tempfile
import threading
import traceback
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import tkinter as tk
from tkinter import filedialog, font, ttk
from tkinter.scrolledtext import ScrolledText

from excel_unmerge_fill import process_file

APP_TITLE = "Excel 数据处理工具"
APP_VERSION = "1.0.0"


def enable_windows_dpi_awareness():
    """在创建 Tk 前启用系统 DPI 感知；已由 exe 清单设置时保持原设置。"""
    if sys.platform != "win32":
        return
    import ctypes
    from ctypes import wintypes

    try:
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        try:
            set_awareness = user32.SetProcessDpiAwarenessContext
        except AttributeError:
            user32.SetProcessDPIAware()
        else:
            set_awareness.argtypes = [ctypes.c_void_p]
            set_awareness.restype = wintypes.BOOL
            # SYSTEM_AWARE = -2；Tk 8.6 按启动显示器的 DPI 缩放。
            # API 返回失败也可能表示清单已经设定 DPI，不再重复覆盖。
            set_awareness(ctypes.c_void_p(-2))
    except (AttributeError, OSError):
        # 旧系统不支持该 API 时仍可正常打开工具。
        pass


def process_batch(files, all_merges, events):
    """后台处理只发送消息，不读取或操作任何 Tk 对象。"""
    failures = 0
    for file in files:
        try:
            output, stats = process_file(file, all_merges=all_merges)
            events.put(("result", file, output, stats, None))
        except Exception as error:
            failures += 1
            events.put(("result", file, None, [], str(error)))
    events.put(("done", len(files), failures))


class Application:
    def __init__(self, root):
        self.root = root
        self.files = ()
        self.running = False
        self.completed = self.succeeded = self.unchanged = self.failed = 0
        self.output_directory = None
        self.events = queue.Queue()
        self.all_merges = tk.BooleanVar(root, value=False)
        self.status = tk.StringVar(root, value="请选择 Excel 文件。")
        root.title("{} v{}".format(APP_TITLE, APP_VERSION))
        scale = max(1.0, root.winfo_fpixels("1i") / 96.0) if sys.platform == "win32" else 1
        width = min(round(840 * scale), root.winfo_screenwidth() - 60)
        height = min(round(660 * scale), root.winfo_screenheight() - 80)
        root.geometry("{}x{}".format(width, height))
        root.minsize(min(round(680 * scale), width), min(round(540 * scale), height))
        root.protocol("WM_DELETE_WINDOW", self.close)

        frame = ttk.Frame(root, padding=round(18 * scale))
        frame.pack(fill="both", expand=True)
        self.heading_font = font.nametofont("TkDefaultFont").copy()
        self.heading_font.configure(size=16, weight="bold")
        ttk.Label(frame, text="合并单元格拆分与填充", font=self.heading_font).pack(anchor="w")
        ttk.Label(frame, text="拆开合并行，让每一行都有原来的值。"
                  ).pack(anchor="w", pady=(6, 12))

        input_frame = ttk.LabelFrame(frame, text="1. 选择文件", padding=10)
        input_frame.pack(fill="x")
        input_toolbar = ttk.Frame(input_frame)
        input_toolbar.pack(fill="x")
        self.choose_button = ttk.Button(input_toolbar, text="选择 Excel 文件（可多选）",
                                        command=self.choose)
        self.choose_button.pack(side="left")
        ttk.Label(input_toolbar, text="支持 .xlsx / .xlsm").pack(side="left", padx=12)
        self.file_list = ScrolledText(input_frame, height=4, wrap="word", state="disabled")
        self.file_list.pack(fill="x", pady=8)
        self.option = ttk.Checkbutton(input_frame, variable=self.all_merges,
                                     text="同时拆开横向合并（默认保留横向表头）")
        self.option.pack(anchor="w")
        ttk.Label(input_frame, text="普通空白不填充；结果另存到原文件目录，不覆盖原表。"
                  ).pack(anchor="w", pady=(6, 0))

        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x", pady=(14, 10))
        self.start_button = ttk.Button(toolbar, text="开始处理", command=self.start,
                                       state="disabled")
        self.start_button.pack(side="left")
        self.open_button = ttk.Button(toolbar, text="打开结果目录", state="disabled",
                                      command=self.open_results)
        self.open_button.pack(side="right")
        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=1, value=0)
        self.progress.pack(fill="x")
        ttk.Label(frame, textvariable=self.status, wraplength=round(760 * scale)).pack(
            anchor="w", pady=(8, 12))
        result_frame = ttk.LabelFrame(frame, text="2. 处理结果", padding=8)
        result_frame.pack(fill="both", expand=True)
        self.results = ScrolledText(result_frame, height=10, wrap="word", state="disabled")
        self.results.pack(fill="both", expand=True)
        self.results.tag_configure("success", foreground="#17603b")
        self.results.tag_configure("unchanged", foreground="#5f6368")
        self.results.tag_configure("failure", foreground="#a72828")

    @staticmethod
    def replace_text(widget, text):
        widget.configure(state="normal")
        widget.delete("1.0", "end")
        widget.insert("end", text)
        widget.configure(state="disabled")

    def choose(self):
        files = filedialog.askopenfilenames(
            parent=self.root, title="选择要处理的 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm")])
        if files:
            self.files = tuple(files)
            self.replace_text(self.file_list, "\n".join(self.files))
            self.status.set("已选择 {} 个文件。".format(len(self.files)))
            self.start_button.configure(state="normal")

    def start(self):
        if self.running or not self.files:
            return
        self.running = True
        self.completed = self.succeeded = self.unchanged = self.failed = 0
        self.output_directory = None
        self.choose_button.configure(state="disabled")
        self.start_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.option.configure(state="disabled")
        self.replace_text(self.results, "")
        self.status.set("正在处理：已完成 0/{} 个文件，请稍候……".format(len(self.files)))
        self.progress.configure(maximum=len(self.files), value=0)
        # 将主线程读取的参数副本传给工作线程；线程不得访问 Tk 变量。
        worker = threading.Thread(target=process_batch,
                                  args=(self.files, self.all_merges.get(), self.events))
        worker.start()
        self.root.after(100, self.poll_results)

    def poll_results(self):
        while True:
            try:
                event = self.events.get_nowait()
            except queue.Empty:
                break
            if event[0] == "done":
                _, count, failures = event
                self.running = False
                self.choose_button.configure(state="normal")
                self.start_button.configure(state="normal")
                self.option.configure(state="normal")
                if self.output_directory is not None:
                    self.open_button.configure(state="normal")
                self.status.set("已完成 {}/{}：成功 {} 个，无需处理 {} 个，失败 {} 个。".format(
                    self.completed, count, self.succeeded, self.unchanged, failures))
            else:
                _, file, output, stats, error = event
                self.completed += 1
                self.progress.configure(value=self.completed)
                if error is not None:
                    self.failed += 1
                    tag = "failure"
                    lines = ["失败  |  " + str(file), "原因：" + error]
                else:
                    if output is not None:
                        self.succeeded += 1
                        self.output_directory = Path(output).parent
                        tag = "success"
                        lines = ["成功  |  " + str(file)]
                    else:
                        self.unchanged += 1
                        tag = "unchanged"
                        lines = ["无需处理  |  " + str(file)]
                    lines.extend("  {}：拆分 {} 个合并区域，填充 {} 个单元格".format(*stat)
                                 for stat in stats)
                    lines.append("已保存：" + str(output) if output else
                                 "没有符合条件的合并区域，未生成新文件。")
                self.results.configure(state="normal")
                self.results.insert("end", "\n".join(lines) + "\n\n", tag)
                self.results.see("end")
                self.results.configure(state="disabled")
                self.status.set("正在处理：已完成 {}/{} 个文件……".format(
                    self.completed, len(self.files)))
        if self.running:
            self.root.after(100, self.poll_results)

    def open_results(self):
        if self.output_directory is None:
            return
        try:
            if not self.output_directory.is_dir():
                raise FileNotFoundError("结果目录已移动或不可访问")
            if sys.platform == "win32":
                os.startfile(str(self.output_directory))
            else:
                subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open",
                                  str(self.output_directory)])
        except OSError as error:
            self.status.set("无法打开结果目录：{}。可从下方复制保存路径。".format(error))

    def close(self):
        if self.running:
            self.status.set("文件仍在处理中，请完成后再关闭窗口。")
        else:
            self.root.destroy()


def self_test():
    """验证打包的 Tcl/Tk 与真实文件处理；只使用临时合成数据。"""
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
    parts = {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>'),
        "_rels/.rels": (
            '<Relationships xmlns="' + pkg + '"><Relationship Id="wb" Type="' + rel
            + '/officeDocument" Target="xl/workbook.xml"/></Relationships>'),
        "xl/workbook.xml": (
            '<workbook xmlns="' + ns + '" xmlns:r="' + rel + '"><sheets>'
            '<sheet name="测试" sheetId="1" r:id="rId1"/></sheets></workbook>'),
        "xl/_rels/workbook.xml.rels": (
            '<Relationships xmlns="' + pkg + '"><Relationship Id="rId1" Type="'
            + rel + '/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'),
        "xl/worksheets/sheet1.xml": (
            '<worksheet xmlns="' + ns + '"><dimension ref="A1:D3"/><sheetData>'
            '<row r="1"><c r="A1" t="inlineStr"><is><t>合并测试</t></is></c>'
            '<c r="B1" t="inlineStr"><is><t>表头</t></is></c>'
            '<c r="D1"><v>0</v></c></row></sheetData><mergeCells count="2">'
            '<mergeCell ref="A1:A3"/><mergeCell ref="B1:C1"/></mergeCells></worksheet>'),
    }
    with tempfile.TemporaryDirectory(prefix="excel-unmerge-self-test-") as directory:
        source = Path(directory) / "中文 自检.xlsx"
        with ZipFile(source, "w", ZIP_DEFLATED) as archive:
            for name, value in parts.items():
                archive.writestr(name, value.encode("utf-8"))
        original = source.read_bytes()
        for all_merges in (False, True):
            output, stats = process_file(source, all_merges=all_merges)
            if output is None or output == source or not output.is_file():
                raise AssertionError("未生成独立输出文件")
            expected_stats = [("测试", 2, 3)] if all_merges else [("测试", 1, 2)]
            if stats != expected_stats:
                raise AssertionError("合并区域或填充数量不正确")
            with ZipFile(output) as archive:
                sheet = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            namespace = {"s": ns}
            values = {cell.attrib["r"]: "".join(cell.itertext())
                      for cell in sheet.findall("s:sheetData/s:row/s:c", namespace)}
            expected = {"A1": "合并测试", "A2": "合并测试", "A3": "合并测试",
                        "B1": "表头", "D1": "0"}
            if all_merges:
                expected["C1"] = "表头"
            if values != expected:
                raise AssertionError("输出单元格内容不正确")
            merges = [node.attrib["ref"] for node in sheet.findall(
                "s:mergeCells/s:mergeCell", namespace)]
            if merges != ([] if all_merges else ["B1:C1"]):
                raise AssertionError("横向合并处理不正确")
        if source.read_bytes() != original:
            raise AssertionError("原文件被修改")

        # 完整窗口以透明方式布局，检查打包后的控件及结果显示路径。
        root = tk.Tk()
        try:
            root.attributes("-alpha", 0)
            app = Application(root)
            root.update()
            controls = (app.choose_button, app.file_list, app.option, app.start_button,
                        app.progress, app.results, app.open_button)
            if any(control.winfo_width() <= 1 or control.winfo_height() <= 1
                   for control in controls):
                raise AssertionError("主窗口控件未正确布局")
            if "disabled" not in app.open_button.state():
                raise AssertionError("生成结果前不应启用结果目录按钮")
            app.files = (str(source),)
            app.running = True
            app.events.put(("result", str(source), output, stats, None))
            app.events.put(("done", 1, 0))
            app.poll_results()
            root.update()
            if (app.running or app.completed != 1 or app.succeeded != 1
                    or app.output_directory != output.parent
                    or "disabled" in app.open_button.state()
                    or str(output) not in app.results.get("1.0", "end")
                    or float(app.progress["value"]) != 1):
                raise AssertionError("成功处理后的窗口状态不正确")
        finally:
            root.destroy()


def main():
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--self-test", action="store_true", help="运行离线打包自检")
    args = parser.parse_args()
    enable_windows_dpi_awareness()
    if args.self_test:
        try:
            self_test()
        except Exception:
            if sys.stderr is not None:
                traceback.print_exc()
            return 1
        return 0
    root = tk.Tk()
    Application(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
