"""Tkinter user interface for the Excel semantic change tracker."""

from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl.utils import get_column_letter

from change_detector import build_change_report
from excel_reader import list_sheet_names, read_filtered_sheet, read_sheet_preview
from excel_writer import save_report
from models import MatchPlan, SheetPreview, WorkbookSnapshot
from semantic_matcher import DEFAULT_THRESHOLD, apply_review_decisions, match_records


@dataclass(frozen=True)
class AnalysisBundle:
    before: WorkbookSnapshot
    after: WorkbookSnapshot
    plan: MatchPlan
    before_content_column: int
    after_content_column: int
    ignored_column_indices: frozenset[int]


@dataclass(frozen=True)
class IgnoredColumnOption:
    index: int
    label: str
    disabled: bool


def build_ignored_column_options(
    before: SheetPreview,
    after: SheetPreview,
    before_content_column: int | None,
    after_content_column: int | None,
) -> tuple[IgnoredColumnOption, ...]:
    """Describe physical-column pairs and prevent either content column from being ignored."""
    before_by_index = {column.index: column for column in before.columns}
    after_by_index = {column.index: column for column in after.columns}
    maximum_columns = max(len(before.columns), len(after.columns))
    content_columns = {before_content_column, after_content_column}
    options = []
    for index in range(maximum_columns):
        letter = get_column_letter(index + 1)
        before_column = before_by_index.get(index)
        after_column = after_by_index.get(index)
        before_letter = before_column.letter if before_column and before_column.letter else letter
        after_letter = after_column.letter if after_column and after_column.letter else letter
        before_title = before_column.header if before_column and before_column.header else "(无标题)"
        after_title = after_column.header if after_column and after_column.header else "(无标题)"
        if before_column is None:
            before_title = "(无此列)"
        if after_column is None:
            after_title = "(无此列)"
        disabled = index in content_columns
        label = f"{before_letter}_before · {before_title}   ↔   {after_letter}_after · {after_title}"
        if disabled:
            label += "（内容列，不可忽略）"
        options.append(IgnoredColumnOption(index, label, disabled))
    return tuple(options)


class ChangeTrackerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel 语义变更跟踪工具")
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = min(1120, max(640, screen_width - 40))
        window_height = min(820, max(600, screen_height - 40))
        self.geometry(f"{window_width}x{window_height}")
        self.minsize(min(980, window_width), min(720, window_height))
        self._messages: queue.Queue = queue.Queue()
        self._previews: dict[str, SheetPreview | None] = {"before": None, "after": None}
        self._column_options: dict[str, dict[str, int]] = {"before": {}, "after": {}}
        self._sheet_generation = {"before": 0, "after": 0}
        self.path_vars = {}
        self.sheet_vars = {}
        self.sheet_boxes = {}
        self.file_info_vars = {}
        self.content_vars = {"before": tk.StringVar(), "after": tk.StringVar()}
        self.type_vars = {"before": tk.StringVar(), "after": tk.StringVar()}
        self.column_boxes = {"before": {}, "after": {}}
        self._ignored_column_options: tuple[IgnoredColumnOption, ...] = ()
        self._ignored_column_vars: dict[int, tk.BooleanVar] = {}
        self._ignored_column_disabled: dict[int, bool] = {}
        self._busy = False
        self._last_output: Path | None = None
        self._build_ui()
        self.after(100, self._poll_messages)

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)

        heading = ttk.Frame(self, padding=(18, 10, 18, 4))
        heading.grid(row=0, column=0, sticky="ew")
        heading.columnconfigure(0, weight=1)
        ttk.Label(heading, text="Excel 语义变更跟踪", font=("Microsoft YaHei UI", 17, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Label(heading, text="选择新旧版本、设置列和筛选类型，然后生成变更报告。", foreground="#555555").grid(row=1, column=0, sticky="w", pady=(4, 0))

        body = ttk.Frame(self, padding=(18, 4, 18, 4))
        body.grid(row=1, column=0, sticky="nsew")
        for column, weight in enumerate((11, 12, 10)):
            body.columnconfigure(column, weight=weight)
        body.rowconfigure(0, weight=1)

        files_box = ttk.LabelFrame(body, text="1. 选择 Excel 文件和工作表", padding=8)
        files_box.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        files_box.columnconfigure(0, weight=1)
        self._build_file_panel(files_box, "before", "旧版本（Before）", 0)
        self._build_file_panel(files_box, "after", "新版本（After）", 1)

        settings = ttk.LabelFrame(body, text="2. 指定内容列、Item Type 列和筛选值", padding=8)
        settings.grid(row=0, column=1, sticky="nsew", padx=(0, 8))
        settings.columnconfigure(1, weight=1)
        settings.rowconfigure(7, weight=1)
        self._build_column_settings(settings)

        actions = ttk.LabelFrame(body, text="3. 匹配和输出", padding=8)
        actions.grid(row=0, column=2, sticky="nsew")
        actions.columnconfigure(1, weight=1)
        self._build_match_settings(actions)

        footer = ttk.Frame(self, padding=(18, 4, 18, 8))
        footer.grid(row=2, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        self.status_var = tk.StringVar(value="请选择两份 Excel 文件。")
        ttk.Label(footer, textvariable=self.status_var, wraplength=700).grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=160)
        self.progress.grid(row=0, column=1, padx=(10, 12))
        self.open_result_button = ttk.Button(footer, text="打开结果文件", command=self._open_result, state="disabled")
        self.open_result_button.grid(row=0, column=2, padx=(0, 6))
        self.open_folder_button = ttk.Button(footer, text="打开输出文件夹", command=self._open_output_folder, state="disabled")
        self.open_folder_button.grid(row=0, column=3)

    def _build_file_panel(self, parent, side: str, title: str, column: int):
        frame = ttk.Frame(parent, padding=(0, 2, 0, 4 if column == 0 else 0))
        frame.grid(row=column, column=0, sticky="ew")
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text=title, font=("Microsoft YaHei UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.path_vars[side] = tk.StringVar()
        row = ttk.Frame(frame)
        row.grid(row=1, column=0, sticky="ew", pady=(5, 4))
        row.columnconfigure(0, weight=1)
        entry = ttk.Entry(row, textvariable=self.path_vars[side], state="readonly")
        entry.grid(row=0, column=0, sticky="ew")
        ttk.Button(row, text="浏览…", command=lambda s=side: self._choose_file(s)).grid(row=0, column=1, padx=(6, 0))
        sheet_row = ttk.Frame(frame)
        sheet_row.grid(row=2, column=0, sticky="ew", pady=(2, 3))
        sheet_row.columnconfigure(1, weight=1)
        ttk.Label(sheet_row, text="工作表").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.sheet_vars[side] = tk.StringVar()
        self.sheet_boxes[side] = ttk.Combobox(sheet_row, textvariable=self.sheet_vars[side], state="disabled", width=34)
        self.sheet_boxes[side].grid(row=0, column=1, sticky="ew")
        self.sheet_boxes[side].bind("<<ComboboxSelected>>", lambda _event, s=side: self._load_selected_sheet(s))
        self.file_info_vars[side] = tk.StringVar(value="尚未加载工作表")
        ttk.Label(frame, textvariable=self.file_info_vars[side], foreground="#555555", wraplength=460).grid(
            row=3, column=0, sticky="w"
        )

    def _build_column_settings(self, parent):
        for group, (side, version) in enumerate((("before", "旧版本"), ("after", "新版本"))):
            heading_row = group * 3
            ttk.Label(parent, text=version, font=("Microsoft YaHei UI", 9, "bold")).grid(
                row=heading_row, column=0, columnspan=2, sticky="w", pady=(4 if group else 0, 0)
            )
            content_row = heading_row + 1
            ttk.Label(parent, text="内容列").grid(row=content_row, column=0, sticky="w", pady=3)
            self.column_boxes[side]["content"] = ttk.Combobox(
                parent, textvariable=self.content_vars[side], state="disabled", width=22
            )
            self.column_boxes[side]["content"].grid(row=content_row, column=1, sticky="ew", padx=(6, 0), pady=3)
            self.column_boxes[side]["content"].bind(
                "<<ComboboxSelected>>", lambda _event: self._refresh_ignored_columns(preserve=True)
            )
            type_row = heading_row + 2
            ttk.Label(parent, text="Item Type 列").grid(row=type_row, column=0, sticky="w", pady=3)
            self.column_boxes[side]["type"] = ttk.Combobox(
                parent, textvariable=self.type_vars[side], state="disabled", width=22
            )
            self.column_boxes[side]["type"].grid(row=type_row, column=1, sticky="ew", padx=(6, 0), pady=3)
        ttk.Label(
            parent,
            text="列选项显示 Excel 字母、列编号和标题。字段报告默认按物理列号配对。",
            foreground="#555555",
            wraplength=340,
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(5, 0))

        ignore_box = ttk.LabelFrame(parent, text="忽略列设置", padding=8)
        ignore_box.grid(row=7, column=0, columnspan=2, sticky="nsew", pady=(6, 0))
        ignore_box.columnconfigure(0, weight=1)
        ignore_box.rowconfigure(1, weight=1)
        controls = ttk.Frame(ignore_box)
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 5))
        ttk.Button(controls, text="全选", command=self._select_all_ignored_columns).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(controls, text="清空选择", command=self._clear_ignored_columns).grid(row=0, column=1)
        ttk.Label(
            controls,
            text="内容列用于语义匹配，不能忽略；Item Type 列仍用于筛选。",
            foreground="#555555",
        ).grid(row=0, column=2, sticky="w", padx=(12, 0))
        list_frame = ttk.Frame(ignore_box)
        list_frame.grid(row=1, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self._ignored_canvas = tk.Canvas(list_frame, height=90, highlightthickness=0)
        self._ignored_canvas.grid(row=0, column=0, sticky="nsew")
        self._ignored_scrollbar = ttk.Scrollbar(
            list_frame, orient="vertical", command=self._ignored_canvas.yview
        )
        self._ignored_scrollbar.grid(row=0, column=1, sticky="ns")
        self._ignored_canvas.configure(yscrollcommand=self._ignored_scrollbar.set)
        self._ignored_items_frame = ttk.Frame(self._ignored_canvas)
        self._ignored_items_window = self._ignored_canvas.create_window(
            (0, 0), window=self._ignored_items_frame, anchor="nw"
        )
        self._ignored_items_frame.bind(
            "<Configure>",
            lambda _event: self._ignored_canvas.configure(scrollregion=self._ignored_canvas.bbox("all")),
        )
        self._ignored_canvas.bind(
            "<Configure>",
            lambda event: self._ignored_canvas.itemconfigure(self._ignored_items_window, width=event.width),
        )
        self._refresh_ignored_columns(preserve=False)

    def _refresh_ignored_columns(self, preserve: bool = True):
        selected = self._selected_ignored_column_indices() if preserve else set()
        for child in self._ignored_items_frame.winfo_children():
            child.destroy()
        self._ignored_column_vars = {}
        self._ignored_column_disabled = {}
        if self._previews["before"] is None or self._previews["after"] is None:
            self._ignored_column_options = ()
            ttk.Label(
                self._ignored_items_frame,
                text="请先加载新旧两侧工作表。",
                foreground="#555555",
            ).grid(row=0, column=0, sticky="w", padx=4, pady=3)
            return

        def content_index(side: str) -> int | None:
            value = self.content_vars[side].get()
            return self._column_options[side].get(value)

        self._ignored_column_options = build_ignored_column_options(
            self._previews["before"], self._previews["after"],
            content_index("before"), content_index("after"),
        )
        if not self._ignored_column_options:
            ttk.Label(
                self._ignored_items_frame,
                text="工作表没有可显示的列。",
                foreground="#555555",
            ).grid(row=0, column=0, sticky="w", padx=4, pady=3)
            return
        for row, option in enumerate(self._ignored_column_options):
            variable = tk.BooleanVar(master=self, value=option.index in selected and not option.disabled)
            self._ignored_column_vars[option.index] = variable
            self._ignored_column_disabled[option.index] = option.disabled
            checkbox = ttk.Checkbutton(
                self._ignored_items_frame,
                text=option.label,
                variable=variable,
                state="disabled" if option.disabled else "normal",
            )
            checkbox.grid(row=row, column=0, sticky="w", padx=4, pady=1)
        self._ignored_canvas.configure(scrollregion=self._ignored_canvas.bbox("all"))

    def _selected_ignored_column_indices(self) -> set[int]:
        return {
            index for index, variable in self._ignored_column_vars.items()
            if variable.get() and not self._ignored_column_disabled.get(index, True)
        }

    def _select_all_ignored_columns(self):
        for index, variable in self._ignored_column_vars.items():
            variable.set(not self._ignored_column_disabled.get(index, True))

    def _clear_ignored_columns(self):
        for variable in self._ignored_column_vars.values():
            variable.set(False)

    def _build_match_settings(self, parent):
        self.filter_var = tk.StringVar(value="Requirement")
        ttk.Label(parent, text="Item Type 筛选值").grid(row=0, column=0, sticky="w")
        self.filter_entry = ttk.Entry(parent, textvariable=self.filter_var, width=16)
        self.filter_entry.grid(row=0, column=1, columnspan=2, sticky="ew", padx=(6, 0))

        self.semantic_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(parent, text="启用本地语义匹配", variable=self.semantic_var).grid(
            row=1, column=0, columnspan=3, sticky="w", pady=(8, 0)
        )
        ttk.Label(parent, text="匹配阈值").grid(row=2, column=0, sticky="w", pady=(8, 0))
        self.threshold_var = tk.DoubleVar(value=DEFAULT_THRESHOLD)
        self.threshold_scale = ttk.Scale(parent, from_=0.50, to=0.95, variable=self.threshold_var, command=self._update_threshold_label)
        self.threshold_scale.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(2, 0))
        self.threshold_label = ttk.Label(parent, text=f"{DEFAULT_THRESHOLD:.2f}")
        self.threshold_label.grid(row=2, column=2, sticky="e", pady=(8, 0))
        ttk.Label(
            parent,
            text="阈值越高，自动匹配越保守；不确定候选可人工确认。",
            foreground="#555555",
            wraplength=220,
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(2, 0))

        ttk.Label(parent, text="输出文件夹").grid(row=5, column=0, columnspan=3, sticky="w", pady=(10, 0))
        self.output_dir_var = tk.StringVar()
        self.output_dir_entry = ttk.Entry(parent, textvariable=self.output_dir_var, state="readonly", width=14)
        self.output_dir_entry.grid(
            row=6, column=0, columnspan=2, sticky="ew", padx=(0, 4), pady=(3, 0)
        )
        ttk.Button(parent, text="选择…", command=self._choose_output_folder).grid(
            row=6, column=2, sticky="e", pady=(3, 0)
        )

        self.run_button = ttk.Button(parent, text="开始分析并生成报告", command=self._start_analysis, width=18)
        self.run_button.grid(row=7, column=0, columnspan=3, sticky="e", pady=(10, 0), ipady=1)
        self.counts_var = tk.StringVar(value="")
        ttk.Label(parent, textvariable=self.counts_var, font=("Microsoft YaHei UI", 10, "bold"), wraplength=300).grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(6, 0)
        )


    def _update_threshold_label(self, _value=None):
        self.threshold_label.configure(text=f"{self.threshold_var.get():.2f}")

    def _choose_file(self, side: str):
        path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 工作簿", "*.xlsx *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.path_vars[side].set(path)
        self._previews[side] = None
        self._column_options[side] = {}
        self.sheet_vars[side].set("")
        self.sheet_boxes[side]["state"] = "disabled"
        self.file_info_vars[side].set("正在加载工作簿…")
        for kind in ("content", "type"):
            self.column_boxes[side][kind]["values"] = ()
            self.column_boxes[side][kind]["state"] = "disabled"
            (self.content_vars if kind == "content" else self.type_vars)[side].set("")
        self._refresh_ignored_columns(preserve=False)
        self.output_dir_var.set(str(Path(path).parent))
        self._sheet_generation[side] += 1
        generation = self._sheet_generation[side]
        self.status_var.set(f"正在读取 {Path(path).name} 的工作表和列标题…")

        def read_metadata(_report):
            names = list_sheet_names(path)
            if not names:
                raise ValueError("工作簿中没有可读取的工作表。")
            preview = read_sheet_preview(path, names[0])
            return names, preview

        self._run_task(
            f"正在读取 {Path(path).name}…",
            read_metadata,
            lambda result: self._metadata_loaded(side, generation, result),
        )

    def _metadata_loaded(self, side: str, generation: int, result):
        if generation != self._sheet_generation[side]:
            return
        names, preview = result
        self.sheet_boxes[side]["values"] = names
        self.sheet_boxes[side]["state"] = "readonly"
        self.sheet_vars[side].set(preview.sheet_name)
        self._set_preview(side, preview)
        self.status_var.set("工作表已加载。请检查两侧列选择后开始分析。")

    def _load_selected_sheet(self, side: str):
        path = self.path_vars[side].get()
        sheet_name = self.sheet_vars[side].get()
        if not path or not sheet_name:
            return
        self._previews[side] = None
        self._column_options[side] = {}
        for kind in ("content", "type"):
            self.column_boxes[side][kind]["values"] = ()
            self.column_boxes[side][kind]["state"] = "disabled"
            (self.content_vars if kind == "content" else self.type_vars)[side].set("")
        self.file_info_vars[side].set(f"正在读取工作表 {sheet_name}…")
        self._refresh_ignored_columns(preserve=False)
        self._sheet_generation[side] += 1
        generation = self._sheet_generation[side]
        self._run_task(
            f"正在读取工作表 {sheet_name}…",
            lambda _report: read_sheet_preview(path, sheet_name),
            lambda preview: self._preview_loaded(side, generation, preview),
        )

    def _preview_loaded(self, side: str, generation: int, preview: SheetPreview):
        if generation != self._sheet_generation[side]:
            return
        self._set_preview(side, preview)
        self.status_var.set(f"已加载工作表 {preview.sheet_name}。")

    def _set_preview(self, side: str, preview: SheetPreview):
        self._previews[side] = preview
        self.file_info_vars[side].set(
            f"{Path(preview.path).name} · {preview.sheet_name} · {preview.data_row_count} 行数据 · {len(preview.columns)} 列"
        )
        options = {}
        for column in preview.columns:
            label = f"{column.letter}  |  {column.number}  |  {column.header or '(无标题)'}"
            options[label] = column.index
        self._column_options[side] = options
        for kind, variable in (("content", self.content_vars[side]), ("type", self.type_vars[side])):
            box = self.column_boxes[side][kind]
            box["values"] = tuple(options)
            box["state"] = "readonly" if options else "disabled"
            proposed = self._recommend_column(side, kind)
            if proposed is not None:
                label = next(label for label, index in options.items() if index == proposed)
                variable.set(label)
            elif variable.get() not in options:
                variable.set("")
        self._refresh_ignored_columns(preserve=False)

    def _recommend_column(self, side: str, kind: str) -> int | None:
        preview = self._previews[side]
        if preview is None or not preview.columns:
            return None
        headers = [(column.index, column.header.casefold()) for column in preview.columns]
        if kind == "type":
            for index, header in headers:
                if header.strip() in {"item type", "item_type", "itemtype"}:
                    return index
            return None
        keywords = ("objecttext", "object text", "description", "requirement", "content", "text")
        matching = [index for index, header in headers if any(keyword in header for keyword in keywords)]
        if not matching:
            return None
        other_side = "after" if side == "before" else "before"
        other = self._previews[other_side]
        if other:
            current_header = dict(headers).get(matching[0])
            same = next((column.index for column in other.columns if column.header.casefold() == current_header), None)
            if same is not None:
                return matching[0]
        return matching[0]

    def _choose_output_folder(self):
        current = self.output_dir_var.get() or str(Path.cwd())
        directory = filedialog.askdirectory(title="选择报告输出文件夹", initialdir=current)
        if directory:
            self.output_dir_var.set(directory)

    def _selected_column(self, side: str, kind: str) -> int:
        value = self.content_vars[side].get() if kind == "content" else self.type_vars[side].get()
        try:
            return self._column_options[side][value]
        except KeyError as exc:
            raise ValueError(("内容列" if kind == "content" else "Item Type 列") + f"未选择或不可用（{side}）。") from exc

    def _start_analysis(self):
        if self._busy:
            return
        before_path, after_path = self.path_vars["before"].get(), self.path_vars["after"].get()
        if not before_path or not after_path:
            messagebox.showwarning("需要两个文件", "请先选择旧版本和新版本 Excel 文件。", parent=self)
            return
        if not self._previews["before"] or not self._previews["after"]:
            messagebox.showwarning("工作表未就绪", "请等待两个工作表的列标题加载完成。", parent=self)
            return
        filter_value = self.filter_var.get().strip()
        if not filter_value:
            messagebox.showwarning("筛选值为空", "请输入 Item Type 筛选值。", parent=self)
            return
        try:
            before_content = self._selected_column("before", "content")
            after_content = self._selected_column("after", "content")
            before_type = self._selected_column("before", "type")
            after_type = self._selected_column("after", "type")
        except ValueError as exc:
            messagebox.showwarning("列选择不完整", str(exc), parent=self)
            return

        before_headers = [column.header for column in self._previews["before"].columns]
        after_headers = [column.header for column in self._previews["after"].columns]
        if before_headers != after_headers:
            proceed = messagebox.askyesno(
                "列结构不同",
                "新旧版本的列标题或列顺序不同。报告将按 Excel 物理列号（A 对 A、B 对 B）配对字段。\n\n确认列结构后继续吗？",
                parent=self,
            )
            if not proceed:
                return
        threshold = float(self.threshold_var.get())
        semantic_enabled = bool(self.semantic_var.get())
        ignored_column_indices = frozenset(self._selected_ignored_column_indices())
        before_sheet_name = self.sheet_vars["before"].get()
        after_sheet_name = self.sheet_vars["after"].get()

        def analyze(report_progress):
            report_progress("正在读取并筛选两个工作表…")
            before_snapshot = read_filtered_sheet(
                before_path, before_sheet_name, before_type, filter_value
            )
            after_snapshot = read_filtered_sheet(
                after_path, after_sheet_name, after_type, filter_value
            )
            if not before_snapshot.records and not after_snapshot.records:
                raise ValueError(f"两份工作表均未找到 Item Type = {filter_value!r} 的记录。请检查筛选值和列选择。")
            plan = match_records(
                before_snapshot.records,
                after_snapshot.records,
                before_content,
                before_content_column=before_content,
                after_content_column=after_content,
                threshold=threshold,
                semantic_enabled=semantic_enabled,
                before_headers=[column.header for column in before_snapshot.columns],
                after_headers=[column.header for column in after_snapshot.columns],
                progress_callback=report_progress,
            )
            return AnalysisBundle(
                before_snapshot, after_snapshot, plan, before_content, after_content, ignored_column_indices
            )

        self._run_task("开始筛选并匹配…", analyze, self._analysis_completed)

    def _analysis_completed(self, bundle: AnalysisBundle):
        plan = bundle.plan
        while plan.review_cases:
            case = plan.review_cases[0]
            decision = MatchReviewDialog(self, case, plan, bundle.after, bundle.before_content_column).show()
            if decision[0] == "skip":
                break
            try:
                plan = apply_review_decisions(plan, {case.before.row_number: decision[1]})
            except (ValueError, StopIteration) as exc:
                messagebox.showerror("复核失败", str(exc), parent=self)
                break

        report = build_change_report(
            bundle.before,
            bundle.after,
            plan,
            bundle.before_content_column,
            bundle.after_content_column,
            ignored_column_indices=bundle.ignored_column_indices,
        )
        output_dir_text = self.output_dir_var.get()
        output_dir = Path(output_dir_text) if output_dir_text else Path(bundle.after.path).parent
        output_path = output_dir / "Change_Tracking_Result.xlsx"
        source_paths = {Path(bundle.before.path).resolve(), Path(bundle.after.path).resolve()}
        if output_path.resolve() in source_paths:
            suffix = 1
            while (output_dir / f"Change_Tracking_Result_{suffix}.xlsx").resolve() in source_paths:
                suffix += 1
            output_path = output_dir / f"Change_Tracking_Result_{suffix}.xlsx"
        if output_path.exists() and not messagebox.askyesno(
            "覆盖现有报告", f"文件已存在：\n{output_path}\n\n是否覆盖？", parent=self
        ):
            self.status_var.set("已取消保存。")
            return

        def write_report(_report):
            return save_report(
                output_path,
                report,
                bundle.before,
                bundle.after,
                bundle.before_content_column,
                bundle.after_content_column,
            )

        self._run_task("正在生成 Excel 报告…", write_report, lambda path: self._report_saved(path, report))

    def _report_saved(self, path: Path, report):
        self._last_output = path
        self.open_result_button.configure(state="normal")
        self.open_folder_button.configure(state="normal")
        stats = report.stats
        self.counts_var.set(
            f"changed {stats.changed_count}  ·  new {stats.new_count}  ·  out {stats.out_count}  ·  still {stats.still_count}  ·  人工确认 {stats.manual_confirmed_count}  ·  待复核 {stats.pending_review_count}"
        )
        self.status_var.set(f"报告已保存：{path}")
        messagebox.showinfo("分析完成", f"报告已生成：\n{path}\n\n{self.counts_var.get()}", parent=self)

    def _open_result(self):
        if self._last_output and self._last_output.exists():
            os.startfile(str(self._last_output))

    def _open_output_folder(self):
        if self._last_output:
            os.startfile(str(self._last_output.parent))

    def _run_task(self, initial_status, work, on_success):
        self._busy = True
        self.run_button.configure(state="disabled")
        self.status_var.set(initial_status)
        self.progress.start(10)

        def report_progress(message):
            self._messages.put(("progress", message))

        def worker():
            try:
                result = work(report_progress)
            except Exception as exc:
                self._messages.put(("error", exc))
            else:
                self._messages.put(("success", (result, on_success)))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_messages(self):
        try:
            while True:
                kind, payload = self._messages.get_nowait()
                if kind == "progress":
                    self.status_var.set(payload)
                elif kind == "success":
                    result, callback = payload
                    self._stop_busy()
                    callback(result)
                elif kind == "error":
                    self._stop_busy()
                    self.status_var.set("操作未完成。请检查文件、列选择或语义模型下载状态。")
                    messagebox.showerror("无法完成", str(payload), parent=self)
        except queue.Empty:
            pass
        self.after(100, self._poll_messages)

    def _stop_busy(self):
        self._busy = False
        self.run_button.configure(state="normal")
        self.progress.stop()


class MatchReviewDialog:
    def __init__(self, parent, case, plan: MatchPlan, after: WorkbookSnapshot, before_content_column: int):
        self.parent = parent
        self.case = case
        self.plan = plan
        self.after = after
        self.before_content_column = before_content_column
        self.available_after_rows = {record.row_number for record in plan.unmatched_after}
        self.result = ("skip", None)

    def show(self):
        window = tk.Toplevel(self.parent)
        window.title("复核语义匹配候选")
        window.geometry("820x540")
        window.transient(self.parent)
        window.grab_set()
        window.columnconfigure(0, weight=1)
        window.rowconfigure(1, weight=1)
        ttk.Label(window, text=f"Before 原始行 {self.case.before.row_number}", font=("Microsoft YaHei UI", 11, "bold")).grid(row=0, column=0, sticky="w", padx=16, pady=(14, 6))
        content_value = self.case.before.values[self.before_content_column]
        content_text = str(content_value) if content_value not in (None, "") else "(内容为空)"
        text = tk.Text(window, height=6, wrap="word", state="normal")
        text.grid(row=1, column=0, sticky="nsew", padx=16)
        text.insert("1.0", content_text)
        text.configure(state="disabled")
        ttk.Label(window, text="选择 After 候选；也可以输入任意 After 原始行号重新指定匹配。", wraplength=780).grid(row=2, column=0, sticky="w", padx=16, pady=(10, 4))

        labels, row_for_label = [], {}
        for candidate in self.case.candidates:
            row = candidate.after
            excerpt = " ".join(str(value or "") for value in row.values)
            label = f"第 {row.row_number} 行  ·  分数 {candidate.score:.3f}  ·  {excerpt[:110]}"
            labels.append(label)
            row_for_label[label] = row.row_number
        selected = tk.StringVar(value=labels[0] if labels else "")
        combo = ttk.Combobox(window, textvariable=selected, values=labels, state="readonly")
        combo.grid(row=3, column=0, sticky="ew", padx=16, pady=(2, 8))
        ttk.Label(window, text="手动指定 After 行号").grid(row=4, column=0, sticky="w", padx=16)
        manual_row = tk.StringVar()
        ttk.Entry(window, textvariable=manual_row).grid(row=5, column=0, sticky="ew", padx=16, pady=(2, 10))

        buttons = ttk.Frame(window, padding=(16, 2, 16, 14))
        buttons.grid(row=6, column=0, sticky="ew")
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=1)
        buttons.columnconfigure(2, weight=1)

        def finish(kind, row_number):
            self.result = (kind, row_number)
            window.destroy()

        def confirm():
            if manual_row.get().strip():
                try:
                    row_number = int(manual_row.get().strip())
                except ValueError:
                    messagebox.showwarning("行号无效", "请输入数字形式的 Excel 行号。", parent=window)
                    return
                if row_number not in self.available_after_rows:
                    messagebox.showwarning("行号不可用", "该行不在 After 的可匹配筛选结果中，或已与其他记录配对。", parent=window)
                    return
                finish("confirm", row_number)
                return
            row_number = row_for_label.get(selected.get())
            if row_number is None:
                messagebox.showwarning("请选择候选", "请先选择一个候选或输入 After 行号。", parent=window)
                return
            finish("confirm", row_number)

        ttk.Button(buttons, text="确认选中匹配", command=confirm).grid(row=0, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(buttons, text="确认不匹配", command=lambda: finish("reject", None)).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(buttons, text="暂不确认", command=lambda: finish("skip", None)).grid(row=0, column=2, sticky="ew", padx=(6, 0))
        window.protocol("WM_DELETE_WINDOW", lambda: finish("skip", None))
        self.parent.wait_window(window)
        return self.result


def main():
    app = ChangeTrackerApp()
    app.mainloop()
