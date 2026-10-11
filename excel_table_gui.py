"""Tk 子窗口：明确选择源表，后台读取表头和执行数据型合表/拆表。"""

import errno
import os
import queue
import threading
import time
import tkinter as tk
import zipfile
from pathlib import Path
from tkinter import filedialog, ttk
from tkinter.scrolledtext import ScrolledText

import excel_table_tools as tables


def _error_text(error):
    if isinstance(error, PermissionError):
        reason = "无法访问文件或目录。请关闭占用文件的程序，并检查读写权限后重试。"
    elif isinstance(error, FileNotFoundError):
        reason = "文件或目录已不存在，请重新选择后重试。"
    elif isinstance(error, zipfile.BadZipFile):
        reason = "工作簿可能损坏、加密或不符合标准，请用 Excel 另存为未加密的 .xlsx 后重试。"
    elif isinstance(error, MemoryError):
        reason = "可用内存不足，请关闭不需要的程序后重试。"
    elif isinstance(error, OSError) and error.errno == errno.ENOSPC:
        reason = "磁盘空间不足，请检查系统盘和结果目录所在磁盘的可用空间。"
    else:
        reason = str(error) or "任务未完成，请检查输入后重试。"
    detail = "{}: {}".format(type(error).__name__, error)
    notes = getattr(error, "__notes__", ())
    return "{}\n技术详情：{}{}".format(reason, detail, "\n" + "\n".join(notes) if notes else "")


def _work(job, operation, snapshot, events, cancel):
    """Only immutable task data and a queue cross the worker boundary, never Tk."""
    def progress(detail):
        events.put((job, "progress", dict(detail)))

    try:
        if cancel.is_set():
            raise tables.TableCancelled("任务已取消。")
        if operation == "inspect":
            result = tables.inspect_workbook(snapshot["path"])
        elif operation == "preview":
            result = tables.preview_sources(snapshot["sources"], progress=progress, cancel=cancel)
        else:
            result = tables.run_tables(progress=progress, cancel=cancel, **snapshot)
        # An execution return means its directory has been published. A late
        # cancellation request must not hide these successfully committed files.
        if operation != "run" and cancel.is_set():
            raise tables.TableCancelled("任务已取消。")
    except tables.TableCancelled as error:
        events.put((job, "cancelled", _error_text(error)))
    except Exception as error:
        events.put((job, "error", _error_text(error)))
    else:
        events.put((job, "result", result))


class TableToolsWindow:
    """One Toplevel in the parent's interpreter; close() never kills a worker."""

    def __init__(self, parent, on_close=None):
        self.window = tk.Toplevel(parent)
        self.window.title("合表与拆表")
        self.window.geometry("900x760")
        self.window.minsize(760, 700)
        self.window.protocol("WM_DELETE_WINDOW", self.close)
        self.on_close = on_close
        self.running = False
        self.closed = False
        self.sources = []
        self.preview_result = None
        self.result = None
        self.events = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker = None
        self.job = 0
        self.operation = None
        self.snapshot = None
        self.started_at = None
        self.last_detail = {}
        self._after_id = None
        self._sheet_names = ()
        self._inspected_path = None
        self._custom_output = False

        self.file = tk.StringVar(self.window)
        self.sheet = tk.StringVar(self.window)
        self.header_row = tk.StringVar(self.window, "1")
        self.output_dir = tk.StringVar(self.window)
        self.mode = tk.StringVar(self.window, "merge")
        self.rows_per_file = tk.StringVar(self.window, "100000")
        self.split_column = tk.StringVar(self.window)
        self.provenance = tk.BooleanVar(self.window, True)
        self.status = tk.StringVar(self.window, "选择文件后，请明确选择工作表和表头行，再添加到来源列表。")
        self.preview_status = tk.StringVar(self.window, "尚未读取表头。")
        self._build()
        self._refresh_controls()
        self._schedule_poll()

    def _build(self):
        frame = ttk.Frame(self.window, padding=14)
        frame.pack(fill="both", expand=True)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(6, weight=1)
        source_box = ttk.LabelFrame(frame, text="1. 选择来源（合表可添加多个工作表，拆表只用一个）", padding=10)
        source_box.grid(row=0, column=0, sticky="ew")
        source_box.columnconfigure(1, weight=1)
        self.choose_button = ttk.Button(source_box, text="选择文件…", command=self.choose_file)
        self.choose_button.grid(row=0, column=0, padx=(0, 8), pady=(0, 8))
        self.file_entry = ttk.Entry(source_box, textvariable=self.file, state="readonly")
        self.file_entry.grid(row=0, column=1, sticky="ew", pady=(0, 8))
        selectors = ttk.Frame(source_box)
        selectors.grid(row=1, column=0, columnspan=2, sticky="ew")
        selectors.columnconfigure(1, weight=1)
        ttk.Label(selectors, text="工作表").grid(row=0, column=0, padx=(0, 6))
        self.sheet_combo = ttk.Combobox(selectors, textvariable=self.sheet, state="readonly", width=18)
        self.sheet_combo.grid(row=0, column=1, sticky="ew")
        self.sheet_combo.bind("<<ComboboxSelected>>", self._refresh_controls)
        ttk.Label(selectors, text="表头行").grid(row=0, column=2, padx=(12, 6))
        self.header_spin = ttk.Spinbox(selectors, from_=1, to=1_048_576, textvariable=self.header_row,
                                      width=8, command=self._refresh_controls)
        self.header_spin.grid(row=0, column=3)
        self.header_spin.bind("<KeyRelease>", self._refresh_controls)
        self.add_button = ttk.Button(selectors, text="添加到来源", command=self.add_source)
        self.add_button.grid(row=0, column=4, padx=(10, 0))
        self.source_list = ttk.Treeview(source_box, columns=("path", "sheet", "header"),
                                       show="headings", height=4, selectmode="extended")
        for column, label, width in (("path", "文件（完整路径）", 400), ("sheet", "工作表", 140), ("header", "表头行", 65)):
            self.source_list.heading(column, text=label)
            self.source_list.column(column, width=width, minwidth=45, stretch=column != "header")
        self.source_list.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        vertical = ttk.Scrollbar(source_box, orient="vertical", command=self.source_list.yview)
        vertical.grid(row=2, column=2, sticky="ns", pady=(8, 0))
        scroll = ttk.Scrollbar(source_box, orient="horizontal", command=self.source_list.xview)
        scroll.grid(row=3, column=0, columnspan=2, sticky="ew")
        self.source_list.configure(xscrollcommand=scroll.set, yscrollcommand=vertical.set)
        self.source_list.bind("<<TreeviewSelect>>", self._refresh_controls)
        self.remove_button = ttk.Button(source_box, text="移除选中来源", command=self.remove_sources)
        self.remove_button.grid(row=4, column=0, columnspan=2, sticky="w", pady=(4, 0))

        options = ttk.LabelFrame(frame, text="2. 处理方式", padding=10)
        options.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        options.columnconfigure(0, weight=1)
        modes = ttk.Frame(options)
        modes.grid(row=0, column=0, sticky="w")
        self.mode_buttons = []
        for i, (value, label) in enumerate((("merge", "合并到一张表"), ("split_rows", "按行数拆表"), ("split_column", "按字段值拆表"))):
            button = ttk.Radiobutton(modes, text=label, value=value, variable=self.mode, command=self._options_changed)
            button.grid(row=0, column=i, padx=(0, 18))
            self.mode_buttons.append(button)
        parameters = ttk.Frame(options)
        parameters.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        parameters.columnconfigure(3, weight=1)
        ttk.Label(parameters, text="每份数据行数").grid(row=0, column=0, padx=(0, 6))
        self.rows_spin = ttk.Spinbox(parameters, from_=1, to=tables.MAX_DATA_ROWS,
                                    textvariable=self.rows_per_file, width=9, command=self._options_changed)
        self.rows_spin.grid(row=0, column=1)
        self.rows_spin.bind("<KeyRelease>", self._options_changed)
        ttk.Label(parameters, text="拆分字段").grid(row=0, column=2, padx=(16, 6))
        self.column_combo = ttk.Combobox(parameters, textvariable=self.split_column, state="readonly", width=18)
        self.column_combo.grid(row=0, column=3, sticky="ew")
        self.column_combo.bind("<<ComboboxSelected>>", self._options_changed)
        self.provenance_check = ttk.Checkbutton(options, text="添加来源信息（文件、工作表、原始行号）", variable=self.provenance)
        self.provenance_check.grid(row=2, column=0, sticky="w", pady=(8, 0))

        output = ttk.Frame(frame)
        output.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        output.columnconfigure(1, weight=1)
        ttk.Label(output, text="结果存放位置").grid(row=0, column=0, padx=(0, 8))
        self.output_entry = ttk.Entry(output, textvariable=self.output_dir, state="readonly")
        self.output_entry.grid(row=0, column=1, sticky="ew")
        self.output_button = ttk.Button(output, text="选择目录…", command=self.choose_output)
        self.output_button.grid(row=0, column=2, padx=(8, 0))
        notice = ttk.Label(frame, text="生成数据型.xlsx：保留值和数字格式，不复制完整样式或宏。所选表中的公式、表头或数据区合并会报错；请先用“拆分填充”处理合并区域。读取表头后，执行时仍会校验全部数据。", wraplength=820)
        notice.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        notice.bind("<Configure>", lambda event: notice.configure(wraplength=max(100, event.width)))
        preview_label = ttk.Label(frame, textvariable=self.preview_status, wraplength=820)
        preview_label.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        preview_label.bind("<Configure>", lambda event: preview_label.configure(wraplength=max(100, event.width)))

        actions = ttk.Frame(frame)
        actions.grid(row=5, column=0, sticky="ew", pady=(8, 0))
        self.preview_button = ttk.Button(actions, text="读取表头", command=self.preview)
        self.preview_button.pack(side="left", padx=(0, 8))
        self.run_button = ttk.Button(actions, text="开始处理", command=self.start)
        self.run_button.pack(side="left", padx=(0, 8))
        self.stop_button = ttk.Button(actions, text="停止任务", command=self.request_cancel)
        self.stop_button.pack(side="left", padx=(0, 8))
        self.open_button = ttk.Button(actions, text="打开最近成功结果", command=self.open_result)
        self.open_button.pack(side="left")
        self.log = ScrolledText(frame, height=4, wrap="word", state="disabled")
        self.log.grid(row=6, column=0, sticky="nsew", pady=(10, 0))
        self.status_label = ttk.Label(frame, textvariable=self.status, wraplength=820)
        self.status_label.grid(row=7, column=0, sticky="ew", pady=(8, 0))
        self.status_label.bind("<Configure>", lambda event: self.status_label.configure(wraplength=max(100, event.width)))

    @staticmethod
    def _enable(widget, enabled, readonly=False):
        widget.configure(state=("readonly" if readonly else "normal") if enabled else "disabled")

    def _run_problem(self):
        if not self.sources:
            return "请先添加来源。"
        if self.preview_result is None:
            return "请先读取表头，成功后才能开始处理。"
        mode = self.mode.get()
        if mode != "merge" and len(self.sources) != 1:
            return "拆表每次只能使用一个来源，请移除其他来源。"
        if mode == "split_rows":
            try:
                count = int(self.rows_per_file.get())
            except ValueError:
                count = 0
            if not 1 <= count <= tables.MAX_DATA_ROWS:
                return "每份数据行数须为 1 至 {:,} 的整数，不含表头。".format(tables.MAX_DATA_ROWS)
        if mode == "split_column" and self.split_column.get() not in self.preview_result["headers"]:
            return "请从已读取的表头中选择拆分字段。"
        if not self.output_dir.get().strip():
            return "请选择结果存放目录。"
        return None

    def _refresh_controls(self, event=None):
        if self.closed:
            return
        idle = not self.running
        for widget in (self.choose_button, self.output_button, self.provenance_check, *self.mode_buttons):
            self._enable(widget, idle)
        self._enable(self.header_spin, idle)
        self._enable(self.sheet_combo, idle and bool(self._sheet_names), readonly=True)
        self._enable(self.add_button, idle and self.sheet.get() in self._sheet_names)
        self._enable(self.remove_button, idle and bool(self.source_list.selection()))
        self._enable(self.rows_spin, idle and self.mode.get() == "split_rows")
        self._enable(self.column_combo, idle and self.mode.get() == "split_column" and self.preview_result is not None, readonly=True)
        self._enable(self.preview_button, idle and bool(self.sources))
        self._enable(self.run_button, idle and self._run_problem() is None)
        self._enable(self.stop_button, self.running and not self.cancel_event.is_set())
        self._enable(self.open_button, idle and self.result is not None)

    def _options_changed(self, event=None):
        self._refresh_controls()
        if not self.running:
            self.status.set(self._run_problem() or "配置已更新，可以开始处理。")

    def choose_file(self):
        if self.running:
            return
        path = filedialog.askopenfilename(parent=self.window, title="选择 Excel 文件", filetypes=[("Excel 工作簿", "*.xlsx *.xlsm")])
        if path:
            self.load_workbook(path)

    def load_workbook(self, path):
        if self.running or self.closed:
            return
        self.file.set(str(Path(path).absolute()))
        self.sheet.set("")
        self._sheet_names = ()
        self._inspected_path = None
        self.sheet_combo.configure(values=())
        self._launch("inspect", {"path": Path(self.file.get())})

    def choose_output(self):
        if self.running:
            return
        path = filedialog.askdirectory(parent=self.window, title="选择结果存放目录", mustexist=True,
                                       initialdir=self.output_dir.get() or None)
        if path:
            self._custom_output = True
            self.output_dir.set(path)
            self._options_changed()

    def add_source(self):
        if self.running:
            return
        if self._inspected_path is None or self.sheet.get() not in self._sheet_names:
            self.status.set("请先选择文件，并明确选择其中的工作表。")
            return
        try:
            header = int(self.header_row.get())
        except ValueError:
            header = 0
        if not 1 <= header <= 1_048_576:
            self.status.set("表头行须为 1 至 1048576 的整数。")
            return
        source = tables.TableSource(self._inspected_path, self.sheet.get(), header)
        if any(os.path.normcase(str(s.path)) == os.path.normcase(str(source.path)) and s.sheet == source.sheet for s in self.sources):
            self.status.set("该文件的工作表已在来源列表中；如需更改表头行，请先移除原项。")
            return
        self.sources.append(source)
        self._sources_changed()

    def remove_sources(self):
        if self.running:
            return
        selected = {int(item) for item in self.source_list.selection()}
        if selected:
            self.sources = [source for i, source in enumerate(self.sources) if i not in selected]
            self._sources_changed()

    def _sources_changed(self):
        if not self._custom_output:
            self.output_dir.set(str(self.sources[0].path.parent) if self.sources else "")
        self.preview_result = None
        self.split_column.set("")
        self.column_combo.configure(values=())
        self.preview_status.set("来源已更新，请读取表头。当前 {} 个来源。".format(len(self.sources)))
        self.source_list.delete(*self.source_list.get_children())
        for i, source in enumerate(self.sources):
            self.source_list.insert("", "end", iid=str(i), values=(str(source.path), source.sheet, source.header_row))
        self.status.set("已选择 {} 个来源。读取表头不会检查全部数据。".format(len(self.sources)))
        self._refresh_controls()

    def preview(self):
        if self.running or not self.sources:
            return
        self.preview_result = None
        self.split_column.set("")
        self.column_combo.configure(values=())
        self.preview_status.set("正在读取所选来源的表头…")
        self._launch("preview", {"sources": tuple(self.sources)})

    def start(self):
        if self.running:
            return
        problem = self._run_problem()
        if problem:
            self.status.set(problem)
            return
        snapshot = dict(sources=tuple(self.sources), output_dir=Path(self.output_dir.get()),
                        mode=self.mode.get(), rows_per_file=int(self.rows_per_file.get()) if self.mode.get() == "split_rows" else 100_000,
                        split_column=self.split_column.get() if self.mode.get() == "split_column" else None,
                        provenance=self.provenance.get())
        self._append_log("开始{}；{} 个来源；结果存放位置：{}".format(
            {"merge": "合表", "split_rows": "按行数拆表", "split_column": "按字段值拆表"}[snapshot["mode"]],
            len(snapshot["sources"]), snapshot["output_dir"]))
        self._launch("run", snapshot)

    def _launch(self, operation, snapshot):
        self.job += 1
        self.operation, self.snapshot = operation, snapshot
        self.running = True
        self.started_at = time.monotonic()
        self.last_detail = {}
        self.cancel_event = threading.Event()
        self._refresh_controls()
        self._refresh_status()
        try:
            self.worker = threading.Thread(target=_work, args=(self.job, operation, snapshot, self.events, self.cancel_event), daemon=True)
            self.worker.start()
        except (MemoryError, RuntimeError) as error:
            self.worker = None
            self._finish("error", _error_text(error))

    def request_cancel(self):
        if self.running:
            self.cancel_event.set()
            self._refresh_controls()
            self._refresh_status()

    def _refresh_status(self):
        if not self.running:
            return
        seconds = max(0, int(time.monotonic() - self.started_at))
        elapsed = "已用时 {:02d}:{:02d}".format(seconds // 60, seconds % 60)
        if self.cancel_event.is_set():
            self.status.set("已请求停止，等待后台结束并清理临时文件，请保留窗口。 · " + elapsed)
            return
        labels = dict(checking="检查表头", reading="读取", writing="写入", filling="写入", grouping="按字段分组",
                      indexing="建立分组索引", saving="保存", publishing="发布结果")
        default = {"inspect": "读取工作表列表", "preview": "读取表头", "run": "开始处理"}[self.operation]
        parts = [labels.get(self.last_detail.get("phase"), default)]
        if self.last_detail.get("sheet"):
            parts.append("工作表：" + str(self.last_detail["sheet"]))
        completed = self.last_detail.get("completed")
        if isinstance(completed, int) and completed >= 0:
            unit = {"rows": "行", "bytes": "字节"}.get(self.last_detail.get("unit"), "项")
            count = "已处理 {:,} {}".format(completed, unit)
            total = self.last_detail.get("total")
            if isinstance(total, int) and total >= completed:
                count += " / {:,} {}".format(total, unit)
            parts.append(count)
        parts.append(elapsed)
        self.status.set(" · ".join(parts))

    def _schedule_poll(self):
        self._after_id = self.window.after(100, self._poll)

    def _poll(self):
        self._after_id = None
        if self.closed:
            return
        while True:
            try:
                job, kind, value = self.events.get_nowait()
            except queue.Empty:
                break
            if job != self.job or not self.running:
                continue
            if kind == "progress":
                self.last_detail = value
            else:
                self._finish(kind, value)
        self._refresh_status()
        self._schedule_poll()

    def _finish(self, kind, value):
        self.running = False
        self.worker = None
        if kind == "result":
            if self.operation == "inspect":
                self._sheet_names = tuple(value)
                self._inspected_path = self.snapshot["path"]
                self.sheet_combo.configure(values=self._sheet_names)
                self.status.set("找到 {} 个工作表。请选择工作表和表头行，再添加到来源。".format(len(value)))
            elif self.operation == "preview":
                self.preview_result = value
                self.column_combo.configure(values=value["headers"])
                self.preview_status.set("已读取 {} 个来源、{} 列表头；完整数据将在执行时校验。".format(len(self.sources), len(value["headers"])))
                self.status.set(self._run_problem() or "表头读取成功，可以开始处理。")
            else:
                self.result = value
                self.status.set("处理完成：{:,} 行，{} 个结果文件。".format(value.rows, len(value.files)))
                self._append_log("{}\n结果目录：{}\n清单：{}".format(self.status.get(), value.directory, value.manifest))
        else:
            if self.operation == "preview":
                self.preview_status.set("表头未读取成功，请检查输入后重试。")
            prefix = "已取消" if kind == "cancelled" else "处理失败"
            self.status.set(prefix + "：未生成本次正式结果。详情见处理记录。")
            self._append_log(prefix + "\n" + value)
        self._refresh_controls()

    def _append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def open_result(self):
        if self.running or self.result is None:
            return
        try:
            os.startfile(str(self.result.directory))
        except OSError as error:
            self._append_log("无法打开结果目录\n" + _error_text(error))
            self.status.set("无法打开结果目录，记录中的完整路径仍可复制。")

    def close(self):
        if self.closed:
            return
        if self.running:
            self.request_cancel()
            return
        self.closed = True
        if self._after_id is not None:
            self.window.after_cancel(self._after_id)
            self._after_id = None
        self.window.destroy()
        callback, self.on_close = self.on_close, None
        if callback:
            callback()
