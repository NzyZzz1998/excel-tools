#!/usr/bin/env python3
"""Excel 合并行拆分的图形入口；也提供打包后的离线自检。"""

import argparse
import errno
import os
import queue
import subprocess
import sys
import tempfile
import threading
import traceback
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.parsers.expat import ExpatError
from zipfile import BadZipFile, ZIP_DEFLATED, ZipFile

import tkinter as tk
from tkinter import filedialog, font, ttk
from tkinter.scrolledtext import ScrolledText

from excel_unmerge_fill import process_file

APP_TITLE = "Excel 数据处理工具"
APP_VERSION = "1.1"


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


def describe_error(error):
    """提供恢复建议，同时保留可复制的原始异常供排查。"""
    if isinstance(error, BadZipFile):
        message = ("无法读取 Excel 文件；它可能已加密、已损坏，或并非标准 .xlsx / .xlsm。"
                   "请先在 Excel 中打开，另存为无密码的副本后重试。")
    elif isinstance(error, PermissionError):
        message = ("文件或结果目录不可访问。请关闭正在编辑文件的应用，"
                   "或把原文件复制到可写目录后重试。")
    elif isinstance(error, FileNotFoundError):
        message = "找不到文件或目录。请确认原文件未被移动、磁盘已连接，再重新选择文件。"
    elif isinstance(error, OSError) and error.errno == errno.ENOSPC:
        message = "处理所需的磁盘空间不足。请检查系统盘及结果目录所在磁盘的可用空间，再重试。"
    elif isinstance(error, OSError) and error.errno == errno.ENAMETOOLONG:
        message = "文件路径过长。请缩短文件名，或复制到较短的目录路径后重试。"
    elif isinstance(error, MemoryError):
        message = "可用内存不足。请关闭其他大型应用，或将工作簿拆成较小的文件后重试。"
    elif isinstance(error, (KeyError, ExpatError)):
        message = "工作簿内部结构无法读取。请在 Excel 中打开并另存为副本后重试。"
    elif isinstance(error, ValueError):
        # 核心处理器的校验错误已包含工作表、单元格和具体操作建议。
        message = str(error) or "文件内容无法处理。请在 Excel 中核对并另存为副本后重试。"
        if message.startswith("找不到文件："):
            message += "\n请确认文件未被移动、磁盘已连接，再重新选择文件。"
    else:
        message = ("处理未完成。请确认文件可在 Excel 中正常打开，并另存为副本后重试；"
                   "若仍失败，可复制下方技术详情排查。")
    return "{}\n技术详情：{}: {}".format(message, type(error).__name__, str(error))


def process_batch(files, all_merges, events, stop_event=None):
    """后台处理只发送消息，不读取或操作任何 Tk 对象。"""
    files = tuple(files)
    failures = completed = 0
    for index, file in enumerate(files):
        # 只在文件边界停止，保证当前文件完成安全保存或失败清理。
        if stop_event is not None and stop_event.is_set():
            break
        events.put(("started", file, index + 1, len(files)))
        try:
            output, stats = process_file(file, all_merges=all_merges)
            events.put(("result", file, output, stats, None))
        except Exception as error:
            failures += 1
            events.put(("result", file, None, [], describe_error(error)))
        completed += 1
    events.put(("done", completed, failures, files[completed:]))


class Application:
    def __init__(self, root):
        self.root = root
        self.files = ()
        self.running = False
        self.completed = self.succeeded = self.unchanged = self.failed = 0
        self.file_states = {}
        self.run_files = ()
        self.run_completed = 0
        self.current_file = None
        self.batch_all_merges = False
        self.stop_event = threading.Event()
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
        recovery_toolbar = ttk.Frame(frame)
        recovery_toolbar.pack(fill="x", pady=(0, 8))
        self.retry_button = ttk.Button(recovery_toolbar, text="仅重试失败项", state="disabled",
                                       command=self.retry_failed)
        self.retry_button.pack(side="left")
        self.resume_button = ttk.Button(recovery_toolbar, text="继续未处理项", state="disabled",
                                        command=self.resume_pending)
        self.resume_button.pack(side="left", padx=8)
        self.stop_button = ttk.Button(recovery_toolbar, text="当前文件完成后停止", state="disabled",
                                      command=self.request_stop)
        self.stop_button.pack(side="right")
        self.progress = ttk.Progressbar(frame, mode="determinate", maximum=1, value=0)
        self.progress.pack(fill="x")
        self.status_label = ttk.Label(frame, textvariable=self.status, anchor="w",
                                      justify="left", wraplength=round(760 * scale))
        self.status_label.pack(fill="x", pady=(8, 12))
        self.status_label.bind("<Configure>", self.resize_status_label)
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

    def resize_status_label(self, event):
        # 按实际可用宽度换行；给标签边缘留余量，避免窄窗裁切长文件名。
        width = max(1, event.width - 4)
        if int(self.status_label.cget("wraplength")) != width:
            self.status_label.configure(wraplength=width)

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
        self.file_states = dict.fromkeys(self.files, "pending")
        self.batch_all_merges = self.all_merges.get()
        self.output_directory = None
        self.replace_text(self.results, "")
        self.launch_batch(tuple(self.file_states), "开始处理")

    @property
    def failed_files(self):
        return tuple(file for file, state in self.file_states.items() if state == "failed")

    @property
    def pending_files(self):
        return tuple(file for file, state in self.file_states.items() if state == "pending")

    def retry_failed(self):
        if not self.running and self.failed_files:
            self.launch_batch(self.failed_files, "仅重试失败项（沿用首次处理规则）")

    def resume_pending(self):
        if not self.running and self.pending_files:
            self.launch_batch(self.pending_files, "继续未处理项（沿用首次处理规则）")

    def update_counts(self):
        states = tuple(self.file_states.values())
        self.succeeded = states.count("success")
        self.unchanged = states.count("unchanged")
        self.failed = states.count("failed")
        self.completed = self.succeeded + self.unchanged + self.failed

    def task_summary(self):
        return "任务已处理 {}/{}：成功 {} 个，无需处理 {} 个，失败 {} 个，待处理 {} 个。".format(
            self.completed, len(self.file_states), self.succeeded, self.unchanged,
            self.failed, len(self.pending_files))

    def append_result(self, text, tag=""):
        self.results.configure(state="normal")
        self.results.insert("end", text + "\n\n", tag)
        self.results.see("end")
        self.results.configure(state="disabled")

    def launch_batch(self, files, label):
        self.running = True
        self.run_files = tuple(files)
        self.run_completed = 0
        self.current_file = None
        self.stop_event = threading.Event()
        self.update_counts()
        # 重试/继续时，复选框也反映真正使用的首次规则。
        self.all_merges.set(self.batch_all_merges)
        self.choose_button.configure(state="disabled")
        self.start_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.retry_button.configure(state="disabled")
        self.resume_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.option.configure(state="disabled")
        rule = "同时拆开横向合并" if self.batch_all_merges else "只拆纵向合并，保留横向表头"
        self.append_result("{}：{} 个文件；规则：{}。".format(label, len(files), rule))
        self.status.set("{}：本轮已处理 0/{} 个文件。".format(label, len(files)))
        self.progress.configure(maximum=len(files), value=0)
        # 将主线程读取的参数副本传给工作线程；线程不得访问 Tk 变量。
        try:
            worker = threading.Thread(target=process_batch,
                                      args=(self.run_files, self.batch_all_merges,
                                            self.events, self.stop_event))
            worker.start()
        except (RuntimeError, MemoryError) as error:
            self.append_result("无法启动处理任务：" + describe_error(error), "failure")
            self.events.put(("done", 0, 0, self.run_files))
        self.root.after(100, self.poll_results)

    def request_stop(self):
        if self.running:
            self.stop_event.set()
            self.stop_button.configure(state="disabled")
            self.status.set("已请求停止；等待当前文件安全完成，尚未开始的文件会保留供继续处理。")

    def poll_results(self):
        while True:
            try:
                event = self.events.get_nowait()
            except queue.Empty:
                break
            if event[0] == "started":
                _, file, index, total = event
                self.current_file = file
                prefix = "正在完成当前文件后停止" if self.stop_event.is_set() else "正在处理"
                self.status.set("{}：{}（本轮第 {}/{} 个，已处理 {} 个）。".format(
                    prefix, Path(file).name, index, total, self.run_completed))
            elif event[0] == "done":
                _, count, failures, remaining = event
                self.running = False
                self.current_file = None
                self.choose_button.configure(state="normal")
                self.start_button.configure(state="normal")
                self.option.configure(state="normal")
                self.stop_button.configure(state="disabled")
                self.retry_button.configure(state="normal" if self.failed_files else "disabled")
                self.resume_button.configure(state="normal" if self.pending_files else "disabled")
                if self.output_directory is not None:
                    self.open_button.configure(state="normal")
                prefix = "已停止" if remaining else "本轮处理结束"
                summary = "{}；{}".format(prefix, self.task_summary())
                self.status.set(summary)
                self.append_result(summary)
            elif event[0] == "result":
                _, file, output, stats, error = event
                self.run_completed += 1
                self.progress.configure(value=self.run_completed)
                if error is not None:
                    self.file_states[file] = "failed"
                    tag = "failure"
                    lines = ["失败  |  " + str(file), "原因：" + error]
                else:
                    if output is not None:
                        self.file_states[file] = "success"
                        self.output_directory = Path(output).parent
                        tag = "success"
                        lines = ["成功  |  " + str(file)]
                    else:
                        self.file_states[file] = "unchanged"
                        tag = "unchanged"
                        lines = ["无需处理  |  " + str(file)]
                    lines.extend("  {}：拆分 {} 个合并区域，填充 {} 个单元格".format(*stat)
                                 for stat in stats)
                    lines.append("已保存：" + str(output) if output else
                                 "没有符合条件的合并区域，未生成新文件。")
                self.update_counts()
                self.append_result("\n".join(lines), tag)
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
            self.request_stop()
            self.status.set("已请求停止。当前文件安全完成后，可查看结果并关闭窗口。")
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
                        app.progress, app.results, app.open_button, app.retry_button,
                        app.resume_button, app.stop_button)
            if any(control.winfo_width() <= 1 or control.winfo_height() <= 1
                   for control in controls):
                raise AssertionError("主窗口控件未正确布局")
            if "disabled" not in app.open_button.state():
                raise AssertionError("生成结果前不应启用结果目录按钮")
            app.files = (str(source),)
            app.file_states = {str(source): "pending"}
            app.run_files = app.files
            app.running = True
            app.events.put(("started", str(source), 1, 1))
            app.events.put(("result", str(source), output, stats, None))
            app.events.put(("done", 1, 0, ()))
            app.poll_results()
            root.update()
            if (app.running or app.completed != 1 or app.succeeded != 1
                    or app.output_directory != output.parent
                    or "disabled" in app.open_button.state()
                    or str(output) not in app.results.get("1.0", "end")
                    or app.failed_files or app.pending_files
                    or any("disabled" not in button.state() for button in (
                        app.retry_button, app.resume_button, app.stop_button))
                    or float(app.progress["value"]) != 1):
                raise AssertionError("成功处理后的窗口状态不正确")
            stopped = threading.Event()
            stopped.set()
            stopped_events = queue.Queue()
            process_batch((str(source),), False, stopped_events, stopped)
            if stopped_events.get_nowait() != ("done", 0, 0, (str(source),)):
                raise AssertionError("停止后不应处理尚未开始的文件")
            error_text = describe_error(BadZipFile("self-test invalid workbook"))
            if "加密" not in error_text or "BadZipFile" not in error_text:
                raise AssertionError("读取失败应包含恢复建议与技术详情")
        finally:
            root.destroy()


def main():
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--self-test", action="store_true", help="运行离线打包自检")
    parser.add_argument("--self-test-log", type=Path, help="将自检结果和异常详情写入 UTF-8 日志")
    args = parser.parse_args()
    enable_windows_dpi_awareness()
    if args.self_test:
        try:
            self_test()
        except Exception:
            diagnostic = traceback.format_exc()
            if sys.stderr is not None:
                sys.stderr.write(diagnostic)
            result = 1
        else:
            diagnostic = "ExcelTools v{} self-test passed.\n".format(APP_VERSION)
            result = 0
        if args.self_test_log is not None:
            try:
                args.self_test_log.parent.mkdir(parents=True, exist_ok=True)
                args.self_test_log.write_text(diagnostic, encoding="utf-8")
            except OSError as error:
                if sys.stderr is not None:
                    sys.stderr.write("无法保存自检日志：{}\n".format(error))
                return 1
        return result
    root = tk.Tk()
    Application(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
