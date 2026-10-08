#!/usr/bin/env python3
"""Excel 合并行拆分的图形入口；也提供打包后的离线自检。"""

import argparse
import queue
import sys
import tempfile
import threading
import traceback
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import tkinter as tk
from tkinter import filedialog, ttk
from tkinter.scrolledtext import ScrolledText

from excel_unmerge_fill import process_file


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
        self.events = queue.Queue()
        self.all_merges = tk.BooleanVar(root, value=False)
        self.status = tk.StringVar(root, value="请选择 Excel 文件。")
        root.title("Excel 合并行拆分")
        root.geometry("800x600")
        root.minsize(620, 460)
        root.protocol("WM_DELETE_WINDOW", self.close)

        frame = ttk.Frame(root, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="拆开合并行，将原值填充到每一行。",
                  font=("", 14, "bold")).pack(anchor="w")
        ttk.Label(frame, text="支持 .xlsx / .xlsm；结果另存到原文件目录，不覆盖原表。"
                  ).pack(anchor="w", pady=(6, 12))
        self.choose_button = ttk.Button(frame, text="选择 Excel 文件（可多选）",
                                        command=self.choose)
        self.choose_button.pack(anchor="w")
        self.file_list = ScrolledText(frame, height=5, wrap="none", state="disabled")
        self.file_list.pack(fill="x", pady=8)
        self.option = ttk.Checkbutton(frame, variable=self.all_merges,
                                     text="同时拆开横向合并（默认保留横向表头）")
        self.option.pack(anchor="w")
        self.start_button = ttk.Button(frame, text="开始处理", command=self.start,
                                       state="disabled")
        self.start_button.pack(anchor="w", pady=(12, 8))
        self.progress = ttk.Progressbar(frame, mode="indeterminate")
        self.progress.pack(fill="x")
        ttk.Label(frame, textvariable=self.status, wraplength=740).pack(
            anchor="w", pady=(8, 6))
        self.results = ScrolledText(frame, height=13, wrap="word", state="disabled")
        self.results.pack(fill="both", expand=True)

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
        self.choose_button.configure(state="disabled")
        self.start_button.configure(state="disabled")
        self.option.configure(state="disabled")
        self.replace_text(self.results, "")
        self.status.set("正在处理，请稍候……")
        self.progress.start(12)
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
                self.progress.stop()
                self.choose_button.configure(state="normal")
                self.start_button.configure(state="normal")
                self.option.configure(state="normal")
                self.status.set("处理完成：共 {} 个文件，失败 {} 个。详见下方结果。".format(
                    count, failures))
            else:
                _, file, output, stats, error = event
                lines = [str(file)]
                if error is not None:
                    lines.append("处理失败：" + error)
                else:
                    lines.extend("  {}：拆分 {} 个合并区域，填充 {} 个单元格".format(*stat)
                                 for stat in stats)
                    lines.append("已保存：" + str(output) if output else
                                 "没有符合条件的合并区域，未生成新文件。")
                self.results.configure(state="normal")
                self.results.insert("end", "\n".join(lines) + "\n\n")
                self.results.see("end")
                self.results.configure(state="disabled")
        if self.running:
            self.root.after(100, self.poll_results)

    def close(self):
        if self.running:
            self.status.set("文件仍在处理中，请完成后再关闭窗口。")
        else:
            self.root.destroy()


def self_test():
    """验证打包的 Tcl/Tk 与真实文件处理；只使用临时合成数据。"""
    root = tk.Tk()
    try:
        root.withdraw()
        root.update_idletasks()
    finally:
        root.destroy()

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


def main():
    parser = argparse.ArgumentParser(description="Excel 合并行拆分图形工具")
    parser.add_argument("--self-test", action="store_true", help="运行离线打包自检")
    args = parser.parse_args()
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
