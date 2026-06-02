from __future__ import annotations

import queue
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import colorchooser, filedialog, messagebox, ttk

from PIL import Image, ImageSequence, ImageTk

from scatter_backend import (
    PlotConfig,
    compute_axis_limits,
    group_pairs_for_combined_axes,
    infer_plot_pairs,
    read_scatter_data,
    run_plotting,
)


APP_NAME = "汪桑的无敌绘图包哈哈哈哈"

LAYOUT_LABELS = {
    "自动判断": "auto",
    "两列一组(X1/Y1, X2/Y2)": "pairs",
    "公共X列(X, Y1, Y2)": "common-x",
}
PLOT_MODE_LABELS = {
    "每组单独出图": "separate",
    "合并勾选的数据列": "combined",
}
SERIES_TYPE_LABELS = {
    "散点图": "scatter",
    "点线图": "line_symbol",
    "折线图": "line",
}
SYMBOL_LABELS = {
    "无": "none",
    "圆形": "circle",
    "空心圆": "hollow_circle",
    "正方形": "square",
    "上三角": "triangle_up",
    "下三角": "triangle_down",
    "菱形": "diamond",
    "加号": "plus",
    "叉号": "cross",
    "星形": "star",
    "六边形": "hexagon",
    "左三角": "left_triangle",
    "右三角": "right_triangle",
    "五边形": "pentagon",
    "细加号": "thin_plus",
    "细叉号": "thin_cross",
    "星号": "asterisk",
    "横线": "horizontal_line",
    "竖线": "vertical_line",
}
SYMBOL_VALUES_TO_LABELS = {value: label for label, value in SYMBOL_LABELS.items()}
SYMBOL_CHOICES = tuple(SYMBOL_LABELS.keys())
LINE_STYLE_LABELS = {
    "实线": "solid",
    "虚线": "dash",
    "点线": "dot",
    "点划线": "dash_dot",
    "双点划线": "dash_dot_dot",
}


def resource_path(relative_path: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base_path / relative_path


class OriginScatterApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_NAME)
        icon_path = resource_path("assets/app_icon.ico")
        if icon_path.exists():
            self.iconbitmap(str(icon_path))

        self.geometry("1160x820")
        self.minsize(1080, 740)
        self.configure(bg="#f8f7fb")

        self.log_queue: queue.Queue[str] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.pair_style_rows: list[dict[str, object]] = []
        self._background_label: tk.Label | None = None
        self._background_source_frames: list[Image.Image] = []
        self._background_photo_frames: list[ImageTk.PhotoImage] = []
        self._background_durations: list[int] = []
        self._background_index = 0
        self._background_animation_job: str | None = None

        self.mode_var = tk.StringVar(value="single")
        self.input_file_var = tk.StringVar()
        self.batch_dir_var = tk.StringVar()
        self.layout_display_var = tk.StringVar(value="自动判断")
        self.plot_mode_display_var = tk.StringVar(value="每组单独出图")
        self.figures_dir_var = tk.StringVar(value=str(Path.cwd() / "figures"))
        self.origin_dir_var = tk.StringVar(value=str(Path.cwd() / "origin"))
        self.single_origin_project_var = tk.BooleanVar(value=False)
        self.axis_width_var = tk.StringVar(value="70")
        self.axis_height_var = tk.StringVar(value="68")

        self._setup_background()
        self._build_ui()
        self.plot_mode_display_var.trace_add("write", lambda *_: self._refresh_combine_state())
        self.after(100, self._drain_logs)

    def _setup_background(self) -> None:
        gif_path = resource_path("assets/app_background.gif")
        if not gif_path.exists():
            return
        try:
            gif = Image.open(gif_path)
            self._background_source_frames = [frame.convert("RGBA") for frame in ImageSequence.Iterator(gif)]
            self._background_durations = [
                max(40, int(frame.info.get("duration", gif.info.get("duration", 80))))
                for frame in ImageSequence.Iterator(gif)
            ]
        except Exception:
            return
        if not self._background_source_frames:
            return
        return

    def _resize_background_frames(self) -> None:
        if not self._background_source_frames or self._background_label is None:
            return
        width = 126
        height = 126
        resized: list[ImageTk.PhotoImage] = []
        for frame in self._background_source_frames:
            scale = max(width / frame.width, height / frame.height)
            new_size = (max(1, int(frame.width * scale)), max(1, int(frame.height * scale)))
            image = frame.resize(new_size, Image.Resampling.LANCZOS)
            left = max(0, (image.width - width) // 2)
            top = max(0, (image.height - height) // 2)
            image = image.crop((left, top, left + width, top + height))
            resized.append(ImageTk.PhotoImage(image))
        self._background_photo_frames = resized
        self._background_index = min(self._background_index, len(resized) - 1)
        if self._background_animation_job is None:
            self._animate_background()
        elif self._background_label is not None:
            self._background_label.configure(image=self._background_photo_frames[self._background_index])

    def _animate_background(self) -> None:
        if not self._background_photo_frames or self._background_label is None:
            return
        frame = self._background_photo_frames[self._background_index]
        self._background_label.configure(image=frame)
        duration = self._background_durations[self._background_index % len(self._background_durations)]
        self._background_index = (self._background_index + 1) % len(self._background_photo_frames)
        self._background_animation_job = self.after(duration, self._animate_background)

    def _build_ui(self) -> None:
        style = ttk.Style(self)
        style.configure("Panel.TFrame", background="#e9e7ef")
        style.configure("TLabel", background="#e9e7ef")
        style.configure("TLabelframe", background="#e9e7ef")
        style.configure("TLabelframe.Label", background="#e9e7ef")

        self.columnconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=0)
        self.rowconfigure(1, weight=0)
        self.rowconfigure(2, weight=0)
        self.rowconfigure(3, weight=0)
        self.rowconfigure(4, weight=0)
        self.rowconfigure(5, weight=0)
        self.rowconfigure(6, weight=1)

        header_frame = ttk.Frame(self, style="Panel.TFrame")
        header_frame.grid(row=0, column=0, columnspan=2, sticky="ew", padx=28, pady=(14, 6))
        header_frame.columnconfigure(0, weight=1)
        ttk.Label(
            header_frame,
            text=APP_NAME,
            font=("Microsoft YaHei UI", 14, "bold"),
        ).grid(row=0, column=0, sticky="w")
        if self._background_source_frames:
            self._background_label = tk.Label(header_frame, borderwidth=0, highlightthickness=0, bg="#f8f7fb")
            self._background_label.grid(row=0, column=1, sticky="e")
            self.after_idle(self._resize_background_frames)

        self._build_basic_tab(self, row_base=1)
        self._build_style_tab(self, row_base=3)

        log_frame = ttk.Frame(self, padding=(8, 4), style="Panel.TFrame")
        log_frame.grid(row=5, column=0, columnspan=2, sticky="ew", padx=28, pady=(0, 16))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=0)
        ttk.Label(log_frame, text="运行日志").grid(row=0, column=0, sticky="w", pady=(0, 4))
        self.log_text = tk.Text(log_frame, wrap="word", height=7, bg="#f8f7fb", relief="flat")
        self.log_text.grid(row=1, column=0, sticky="ew")
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _build_basic_tab(self, parent, row_base: int = 0) -> None:
        parent.columnconfigure(0, weight=1)
        parent.columnconfigure(1, weight=1)

        data_box = ttk.LabelFrame(parent, text="数据来源", padding=8)
        data_box.grid(row=row_base, column=0, columnspan=2, sticky="ew", padx=28, pady=(22, 10))
        data_box.columnconfigure(1, weight=1)
        data_box.columnconfigure(3, weight=1)

        ttk.Label(data_box, text="运行模式").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        mode_frame = ttk.Frame(data_box)
        mode_frame.grid(row=0, column=1, columnspan=3, sticky="w", pady=5)
        for text, value in (
            ("单个数据文件", "single"),
            ("批量数据文件夹", "batch"),
            ("Excel 多 Sheet", "excel"),
        ):
            ttk.Radiobutton(mode_frame, text=text, value=value, variable=self.mode_var).pack(side="left", padx=(0, 14))

        self._path_row(data_box, 1, "数据文件", self.input_file_var, self._choose_input_file)
        self._path_row(data_box, 2, "批量文件夹", self.batch_dir_var, self._choose_batch_dir)

        options_box = ttk.LabelFrame(parent, text="读取规则", padding=8)
        options_box.grid(row=row_base + 1, column=0, sticky="nsew", padx=(28, 8), pady=(0, 10))
        options_box.columnconfigure(1, weight=1)
        ttk.Label(options_box, text="数据结构").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Combobox(
            options_box,
            textvariable=self.layout_display_var,
            values=tuple(LAYOUT_LABELS.keys()),
            state="readonly",
            width=28,
        ).grid(row=0, column=1, sticky="w", pady=5)
        ttk.Label(options_box, text="X轴长度(%)").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(options_box, textvariable=self.axis_width_var, width=10).grid(row=1, column=1, sticky="w", pady=5)
        ttk.Label(options_box, text="Y轴长度(%)").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(options_box, textvariable=self.axis_height_var, width=10).grid(row=2, column=1, sticky="w", pady=5)
        output_box = ttk.LabelFrame(parent, text="输出位置", padding=8)
        output_box.grid(row=row_base + 1, column=1, sticky="nsew", padx=(8, 28), pady=(0, 10))
        output_box.columnconfigure(1, weight=1)
        output_box.columnconfigure(2, weight=1)
        self._path_row(output_box, 0, "EMF 输出文件夹", self.figures_dir_var, self._choose_figures_dir)
        self._path_row(output_box, 1, "Origin OPJU 输出文件夹", self.origin_dir_var, self._choose_origin_dir)
        ttk.Checkbutton(
            output_box,
            text="所有图页保存到同一个 Origin 项目",
            variable=self.single_origin_project_var,
        ).grid(row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))

    def _build_style_tab(self, parent, row_base: int = 0) -> None:
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        control_box = ttk.LabelFrame(parent, text="出图流程", padding=8)
        control_box.grid(row=row_base, column=0, columnspan=2, sticky="ew", padx=28, pady=(0, 10))
        control_box.columnconfigure(3, weight=1)
        ttk.Button(control_box, text="读取数据列", command=self._load_column_styles).grid(row=0, column=0, sticky="w", padx=(0, 10))
        ttk.Label(control_box, text="绘图方式").grid(row=0, column=1, sticky="w", padx=(0, 6))
        ttk.Combobox(
            control_box,
            textvariable=self.plot_mode_display_var,
            values=tuple(PLOT_MODE_LABELS.keys()),
            state="readonly",
            width=20,
        ).grid(row=0, column=2, sticky="w", padx=(0, 14))
        self.start_button_hint = ttk.Label(control_box, text="读取后逐列设置图形、坐标轴和拟合")
        self.start_button_hint.grid(row=0, column=3, sticky="w")
        self.preview_button = ttk.Button(control_box, text="预览出图", command=self._preview_plotting)
        self.preview_button.grid(row=0, column=4, sticky="e", padx=(0, 10))
        self.start_button = ttk.Button(control_box, text="开始绘图", command=self._start_plotting)
        self.start_button.grid(row=0, column=5, sticky="e")

        style_box = ttk.LabelFrame(parent, text="数据列设置", padding=8)
        style_box.grid(row=row_base + 1, column=0, columnspan=2, sticky="ew", padx=28, pady=(0, 10))
        style_box.columnconfigure(0, weight=1)
        style_box.rowconfigure(0, weight=1)

        scroll_frame = ttk.Frame(style_box)
        scroll_frame.grid(row=0, column=0, sticky="nsew")
        scroll_frame.columnconfigure(0, weight=1)
        scroll_frame.rowconfigure(0, weight=1)

        self.style_canvas = tk.Canvas(scroll_frame, height=290, highlightthickness=0)
        self.style_canvas.grid(row=0, column=0, sticky="nsew")
        self.style_scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=self.style_canvas.yview)
        self.style_scrollbar.grid(row=0, column=1, sticky="ns")
        self.style_x_scrollbar = ttk.Scrollbar(scroll_frame, orient="horizontal", command=self.style_canvas.xview)
        self.style_x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.style_canvas.configure(
            yscrollcommand=self.style_scrollbar.set,
            xscrollcommand=self.style_x_scrollbar.set,
        )

        self.style_rows_frame = ttk.Frame(self.style_canvas)
        self.style_canvas_window = self.style_canvas.create_window((0, 0), window=self.style_rows_frame, anchor="nw")
        self.style_rows_frame.bind("<Configure>", self._update_style_scroll_region)
        self.style_canvas.bind("<Configure>", self._resize_style_canvas_window)
        self.style_canvas.bind("<Enter>", self._bind_style_mousewheel)
        self.style_canvas.bind("<Leave>", self._unbind_style_mousewheel)

        for col in range(16):
            self.style_rows_frame.columnconfigure(col, weight=1 if col == 2 else 0)

    def _path_row(self, parent: ttk.Frame, row: int, label: str, variable: tk.StringVar, command) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 6), pady=3)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, columnspan=2, sticky="ew", pady=3)
        ttk.Button(parent, text="选择", command=command).grid(row=row, column=3, sticky="e", pady=3)

    def _choose_input_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=[
                ("数据文件", "*.xlsx *.xlsm *.xls *.csv *.txt *.dat"),
                ("所有文件", "*.*"),
            ],
        )
        if filename:
            self.input_file_var.set(filename)

    def _choose_batch_dir(self) -> None:
        directory = filedialog.askdirectory(title="选择批量数据文件夹")
        if directory:
            self.batch_dir_var.set(directory)

    def _choose_figures_dir(self) -> None:
        directory = filedialog.askdirectory(title="选择 EMF 输出文件夹")
        if directory:
            self.figures_dir_var.set(directory)

    def _choose_origin_dir(self) -> None:
        directory = filedialog.askdirectory(title="选择 Origin OPJU 输出文件夹")
        if directory:
            self.origin_dir_var.set(directory)

    def _update_style_scroll_region(self, _event=None) -> None:
        self.style_canvas.configure(scrollregion=self.style_canvas.bbox("all"))

    def _resize_style_canvas_window(self, event: tk.Event) -> None:
        requested = self.style_rows_frame.winfo_reqwidth()
        self.style_canvas.itemconfigure(self.style_canvas_window, width=max(event.width, requested))

    def _bind_style_mousewheel(self, _event=None) -> None:
        self.style_canvas.bind_all("<MouseWheel>", self._on_style_mousewheel)

    def _unbind_style_mousewheel(self, _event=None) -> None:
        self.style_canvas.unbind_all("<MouseWheel>")

    def _on_style_mousewheel(self, event: tk.Event) -> None:
        self.style_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _load_column_styles(self) -> None:
        try:
            data_path, sheet_name = self._preview_dataset()
            df = read_scatter_data(data_path, sheet_name, self._log)
            pairs = infer_plot_pairs(df, LAYOUT_LABELS[self.layout_display_var.get()], self._log)
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return

        self._clear_style_rows()
        for index, pair in enumerate(pairs):
            self._add_style_row(index=index, title=f"{pair.x_title} / {pair.y_title}")
        self._refresh_combine_state()
        self._log(f"已读取 {len(pairs)} 组数据列，请在表格中逐列选择出图类型和参数。")

    def _preview_dataset(self) -> tuple[Path, str | int | None]:
        mode = self.mode_var.get()
        if mode in {"single", "excel"}:
            input_file = self.input_file_var.get().strip()
            if not input_file:
                raise ValueError("请先选择数据文件。")
            return Path(input_file), 0

        batch_dir = self.batch_dir_var.get().strip()
        if not batch_dir:
            raise ValueError("请先选择批量数据文件夹。")
        files = []
        for pattern in ("*.xlsx", "*.xlsm", "*.xls", "*.csv", "*.txt", "*.dat"):
            files.extend(Path(batch_dir).glob(pattern))
        files = sorted(path for path in files if path.is_file() and not path.name.startswith("~$"))
        if not files:
            raise ValueError("批量文件夹中没有可读取的数据文件。")
        return files[0], None

    def _clear_style_rows(self) -> None:
        for child in self.style_rows_frame.winfo_children():
            child.destroy()
        self.pair_style_rows.clear()

    def _add_style_row(self, index: int, title: str) -> None:
        row = index + 1
        if index == 0:
            headers = (
                "组",
                "出图类型",
                "数据列",
                "合并",
                "点颜色",
                "点类型",
                "点大小",
                "线颜色",
                "线宽",
                "线型",
                "X对数",
                "Y倒序",
                "线性拟合",
            )
            for col, text in enumerate(headers):
                ttk.Label(self.style_rows_frame, text=text).grid(row=0, column=col, sticky="w", padx=(0, 8), pady=3)

        series_type_var = tk.StringVar(value="散点图")
        combine_var = tk.BooleanVar(value=False)
        marker_color_var = tk.StringVar(value="#1F77B4")
        marker_symbol_var = tk.StringVar(value="圆形")
        marker_size_var = tk.StringVar(value="8")
        hollow_fill_color_var = tk.StringVar(value="#FFFFFF")
        hollow_edge_color_var = tk.StringVar(value="#1F77B4")
        hollow_edge_width_var = tk.StringVar(value="1.5")
        hollow_dialog_opened_var = tk.BooleanVar(value=False)
        line_color_var = tk.StringVar(value="#1F77B4")
        line_width_var = tk.StringVar(value="2")
        line_style_var = tk.StringVar(value="实线")
        x_log_var = tk.BooleanVar(value=False)
        y_reverse_var = tk.BooleanVar(value=False)
        linear_fit_var = tk.BooleanVar(value=False)
        axis_width_row_var = tk.StringVar(value=self.axis_width_var.get() or "70")
        axis_height_row_var = tk.StringVar(value=self.axis_height_var.get() or "68")

        ttk.Label(self.style_rows_frame, text=str(row)).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
        type_box = ttk.Combobox(
            self.style_rows_frame,
            textvariable=series_type_var,
            values=tuple(SERIES_TYPE_LABELS.keys()),
            state="readonly",
            width=9,
        )
        type_box.grid(row=row, column=1, sticky="w", padx=(0, 8), pady=3)
        ttk.Label(self.style_rows_frame, text=title).grid(row=row, column=2, sticky="ew", padx=(0, 8), pady=3)

        combine_widget = ttk.Checkbutton(self.style_rows_frame, variable=combine_var)
        combine_widget.grid(row=row, column=3, sticky="w", padx=(0, 8), pady=3)

        marker_color_button = tk.Button(self.style_rows_frame, text="选择", width=8, bg=marker_color_var.get())
        marker_color_button.grid(row=row, column=4, sticky="w", padx=(0, 8), pady=3)
        marker_color_button.configure(
            command=lambda var=marker_color_var, button=marker_color_button, title="选择点颜色": self._choose_row_color(var, button, title)
        )
        marker_symbol_box = ttk.Combobox(
            self.style_rows_frame,
            textvariable=marker_symbol_var,
            values=SYMBOL_CHOICES,
            state="readonly",
            width=12,
        )
        marker_symbol_box.grid(row=row, column=5, sticky="w", padx=(0, 8), pady=3)
        marker_size_frame = ttk.Frame(self.style_rows_frame)
        marker_size_frame.grid(row=row, column=6, sticky="w", padx=(0, 8), pady=3)
        marker_size_entry = ttk.Entry(marker_size_frame, textvariable=marker_size_var, width=6)
        marker_size_entry.grid(row=0, column=0, sticky="w")
        hollow_settings_button = ttk.Button(marker_size_frame, text="设置", width=5)
        hollow_settings_button.grid(row=0, column=1, sticky="w", padx=(4, 0))

        line_color_button = tk.Button(self.style_rows_frame, text="选择", width=8, bg=line_color_var.get())
        line_color_button.grid(row=row, column=7, sticky="w", padx=(0, 8), pady=3)
        line_color_button.configure(
            command=lambda var=line_color_var, button=line_color_button, title="选择线条颜色": self._choose_row_color(var, button, title)
        )
        line_width_entry = ttk.Entry(self.style_rows_frame, textvariable=line_width_var, width=7)
        line_width_entry.grid(row=row, column=8, sticky="w", padx=(0, 8), pady=3)
        line_style_box = ttk.Combobox(
            self.style_rows_frame,
            textvariable=line_style_var,
            values=tuple(LINE_STYLE_LABELS.keys()),
            state="readonly",
            width=10,
        )
        line_style_box.grid(row=row, column=9, sticky="w", padx=(0, 8), pady=3)

        ttk.Checkbutton(self.style_rows_frame, variable=x_log_var).grid(row=row, column=10, sticky="w", padx=(0, 8), pady=3)
        ttk.Checkbutton(self.style_rows_frame, variable=y_reverse_var).grid(row=row, column=11, sticky="w", padx=(0, 8), pady=3)
        ttk.Checkbutton(self.style_rows_frame, variable=linear_fit_var).grid(row=row, column=12, sticky="w", pady=3)
        ttk.Entry(self.style_rows_frame, textvariable=axis_width_row_var, width=8).grid(row=row, column=13, sticky="w", padx=(8, 8), pady=3)
        ttk.Entry(self.style_rows_frame, textvariable=axis_height_row_var, width=8).grid(row=row, column=14, sticky="w", padx=(0, 8), pady=3)

        row_state = {
            "series_type": series_type_var,
            "combine": combine_var,
            "combine_widget": combine_widget,
            "marker_color": marker_color_var,
            "marker_symbol": marker_symbol_var,
            "marker_size": marker_size_var,
            "hollow_fill_color": hollow_fill_color_var,
            "hollow_edge_color": hollow_edge_color_var,
            "hollow_edge_width": hollow_edge_width_var,
            "hollow_dialog_opened": hollow_dialog_opened_var,
            "line_color": line_color_var,
            "line_width": line_width_var,
            "line_style": line_style_var,
            "x_log": x_log_var,
            "y_reverse": y_reverse_var,
            "linear_fit": linear_fit_var,
            "axis_width": axis_width_row_var,
            "axis_height": axis_height_row_var,
            "marker_button": marker_color_button,
            "hollow_button": hollow_settings_button,
            "line_button": line_color_button,
            "point_widgets": [marker_color_button, marker_symbol_box, marker_size_frame],
            "hollow_widgets": [hollow_settings_button],
            "line_widgets": [line_color_button, line_width_entry, line_style_box],
        }
        hollow_settings_button.configure(command=lambda state=row_state: self._open_hollow_marker_dialog(state))
        self.pair_style_rows.append(row_state)
        series_type_var.trace_add("write", lambda *_args, state=row_state: self._refresh_row_widgets(state))
        marker_symbol_var.trace_add("write", lambda *_args, state=row_state: self._on_marker_symbol_changed(state))
        self._refresh_row_widgets(row_state)

    def _refresh_row_widgets(self, row_state: dict[str, object]) -> None:
        plot_type = self._series_type_value(row_state["series_type"].get())
        show_points = plot_type in {"scatter", "line_symbol"}
        show_line = plot_type in {"line", "line_symbol"}
        show_hollow = show_points and self._marker_symbol_from_display(row_state["marker_symbol"].get()) == "hollow_circle"
        self._set_widgets_visible(row_state["point_widgets"], show_points)
        self._set_widgets_visible(row_state["hollow_widgets"], show_hollow)
        self._set_widgets_visible(row_state["line_widgets"], show_line)

    def _set_widgets_visible(self, widgets: list[tk.Widget], visible: bool) -> None:
        for widget in widgets:
            if visible:
                widget.grid()
            else:
                widget.grid_remove()

    def _on_marker_symbol_changed(self, row_state: dict[str, object]) -> None:
        self._refresh_row_widgets(row_state)
        is_hollow = self._marker_symbol_from_display(row_state["marker_symbol"].get()) == "hollow_circle"
        if is_hollow and not bool(row_state["hollow_dialog_opened"].get()):
            row_state["hollow_dialog_opened"].set(True)
            self.after(50, lambda state=row_state: self._open_hollow_marker_dialog(state))

    def _open_hollow_marker_dialog(self, row_state: dict[str, object]) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("空心圆设置")
        dialog.transient(self)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.configure(bg="#f8f7fb")

        ttk.Label(dialog, text="空心圆内部颜色").grid(row=0, column=0, sticky="w", padx=14, pady=(14, 6))
        fill_button = tk.Button(dialog, text="选择", width=10, bg=row_state["hollow_fill_color"].get())
        fill_button.grid(row=0, column=1, sticky="w", padx=14, pady=(14, 6))
        fill_button.configure(
            command=lambda: self._choose_row_color(row_state["hollow_fill_color"], fill_button, "选择空心圆内部颜色")
        )

        ttk.Label(dialog, text="空心圆边缘颜色").grid(row=1, column=0, sticky="w", padx=14, pady=6)
        edge_button = tk.Button(dialog, text="选择", width=10, bg=row_state["hollow_edge_color"].get())
        edge_button.grid(row=1, column=1, sticky="w", padx=14, pady=6)
        edge_button.configure(
            command=lambda: self._choose_row_color(row_state["hollow_edge_color"], edge_button, "选择空心圆边缘颜色")
        )

        ttk.Label(dialog, text="空心圆圆圈线宽").grid(row=2, column=0, sticky="w", padx=14, pady=6)
        ttk.Entry(dialog, textvariable=row_state["hollow_edge_width"], width=12).grid(row=2, column=1, sticky="w", padx=14, pady=6)

        ttk.Button(dialog, text="确定", command=dialog.destroy).grid(row=3, column=0, columnspan=2, sticky="e", padx=14, pady=(8, 14))

    def _refresh_combine_state(self) -> None:
        allow_combine = self._plot_mode_value() == "combined"
        for row in self.pair_style_rows:
            widget = row["combine_widget"]
            if allow_combine:
                widget.configure(state="normal")
            else:
                row["combine"].set(False)
                widget.configure(state="disabled")

    def _choose_row_color(self, color_var: tk.StringVar, button: tk.Button, title: str) -> None:
        _, hex_color = colorchooser.askcolor(title=title, color=color_var.get())
        if hex_color:
            color_var.set(hex_color)
            button.configure(bg=hex_color)

    def _preview_plotting(self) -> None:
        try:
            if not self.pair_style_rows:
                self._load_column_styles()
                if not self.pair_style_rows:
                    return
            config = self._collect_config()
            data_path, sheet_name = self._preview_dataset()
            df = read_scatter_data(data_path, sheet_name, lambda _message: None)
            pairs = infer_plot_pairs(df, LAYOUT_LABELS[self.layout_display_var.get()], lambda _message: None)
            if config.plot_mode == "combined" and config.combine_indices:
                combined_pairs = [pairs[index] for index in config.combine_indices if index < len(pairs)]
                separate_pairs = [pair for index, pair in enumerate(pairs) if index not in config.combine_indices]
                axis_groups = group_pairs_for_combined_axes(df, combined_pairs, config, lambda _message: None)
            else:
                separate_pairs = pairs
                axis_groups = []
        except Exception as exc:
            messagebox.showerror("预览失败", str(exc))
            return

        preview = tk.Toplevel(self)
        preview.title("出图预览")
        preview.geometry("860x620")
        preview.minsize(760, 520)
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(1, weight=1)

        canvas = tk.Canvas(preview, height=230, bg="#fbfafc", highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="ew", padx=14, pady=(14, 8))
        self._draw_preview_canvas(canvas, axis_groups, separate_pairs)

        text_frame = ttk.Frame(preview, padding=(14, 0, 14, 14))
        text_frame.grid(row=1, column=0, sticky="nsew")
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        preview_text = tk.Text(text_frame, wrap="word", bg="#ffffff", relief="solid", borderwidth=1)
        preview_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=preview_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        preview_text.configure(yscrollcommand=scrollbar.set)
        preview_text.insert("1.0", self._build_preview_text(data_path, df, pairs, axis_groups, separate_pairs, config))
        preview_text.configure(state="disabled")

    def _draw_preview_canvas(self, canvas: tk.Canvas, axis_groups, separate_pairs) -> None:
        canvas.delete("all")
        canvas.create_text(24, 20, text="预览布局", anchor="w", font=("Microsoft YaHei UI", 12, "bold"))
        if axis_groups:
            canvas.create_rectangle(70, 58, 560, 190, outline="#1f2937", width=2)
            canvas.create_text(315, 42, text="合并导出：同一绘图区，多坐标轴", font=("Microsoft YaHei UI", 10, "bold"))
            canvas.create_line(70, 190, 560, 190, fill="#1f2937", width=2)
            canvas.create_line(70, 58, 70, 190, fill="#1f2937", width=2)
            colors = ("#1f77b4", "#d73027", "#4daf4a", "#984ea3", "#ff7f00", "#636363")
            for index, group in enumerate(axis_groups):
                color = colors[index % len(colors)]
                y = 82 + index * 20
                canvas.create_line(92, y, 250, y + 42, fill=color, width=2)
                canvas.create_oval(88, y - 4, 96, y + 4, fill=color, outline=color)
                canvas.create_oval(246, y + 38, 254, y + 46, fill=color, outline=color)
                canvas.create_text(585, y, text=f"轴组 {index + 1}: {len(group)} 组数据", anchor="w", fill=color, font=("Microsoft YaHei UI", 9))
            canvas.create_text(315, 211, text="导出前自动执行 Fit Page to Layers", fill="#374151", font=("Microsoft YaHei UI", 9))
        else:
            canvas.create_rectangle(90, 70, 280, 180, outline="#1f77b4", width=2)
            canvas.create_text(185, 52, text="单独出图预览", font=("Microsoft YaHei UI", 10, "bold"))
        if separate_pairs:
            start_x = 610
            canvas.create_text(start_x, 42, text=f"单独导出：{len(separate_pairs)} 张", anchor="w", font=("Microsoft YaHei UI", 10, "bold"))
            for index, _pair in enumerate(separate_pairs[:5]):
                top = 66 + index * 28
                canvas.create_rectangle(start_x, top, start_x + 80, top + 20, outline="#64748b")
                canvas.create_text(start_x + 92, top + 10, text=f"图 {index + 1}", anchor="w", font=("Microsoft YaHei UI", 9))
            if len(separate_pairs) > 5:
                canvas.create_text(start_x, 208, text=f"... 还有 {len(separate_pairs) - 5} 张", anchor="w", fill="#64748b")

    def _build_preview_text(self, data_path: Path, df, pairs, axis_groups, separate_pairs, config: PlotConfig) -> str:
        lines: list[str] = []
        lines.append(f"数据文件：{data_path}")
        lines.append(f"数据行列：{df.shape[0]} 行 x {df.shape[1]} 列")
        lines.append("")
        if axis_groups:
            lines.append("合并图：同一绘图区，多坐标轴")
            for group_index, group in enumerate(axis_groups, start=1):
                x_log = any(self._config_bool_at(config.x_logs, pair.index) for pair in group)
                limits = compute_axis_limits(df, group, None, x_log=x_log)
                names = "；".join(f"{pair.x_title} / {pair.y_title}" for pair in group)
                lines.append(
                    f"  轴组 {group_index}: {names}\n"
                    f"    X范围 {limits.x_min:g} - {limits.x_max:g}，步长 {limits.x_step:g}"
                    f"{'，对数轴' if x_log else ''}\n"
                    f"    Y范围 {limits.y_min:g} - {limits.y_max:g}，步长 {limits.y_step:g}"
                )
            lines.append("")
        if separate_pairs:
            lines.append("单独导出的数据列：")
            for pair in separate_pairs:
                x_log = self._config_bool_at(config.x_logs, pair.index)
                limits = compute_axis_limits(df, [pair], None, x_log=x_log)
                plot_type = config.plot_types[pair.index] if pair.index < len(config.plot_types) else config.plot_type
                lines.append(
                    f"  图 {pair.index + 1}: {pair.x_title} / {pair.y_title}，类型 {plot_type}\n"
                    f"    X范围 {limits.x_min:g} - {limits.x_max:g}"
                    f"{'，对数轴' if x_log else ''}；Y范围 {limits.y_min:g} - {limits.y_max:g}"
                )
        if not axis_groups and not separate_pairs:
            lines.append("没有可预览的数据列。")
        return "\n".join(lines)

    @staticmethod
    def _config_bool_at(values: list[bool], index: int) -> bool:
        return bool(values[index]) if index < len(values) else False

    def _start_plotting(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("正在运行", "绘图任务正在后台运行，请等待完成。")
            return

        try:
            config = self._collect_config()
        except ValueError as exc:
            messagebox.showerror("参数错误", str(exc))
            return

        self.log_text.delete("1.0", "end")
        self.start_button.configure(state="disabled")
        self._log("开始绘图任务...")
        self.worker = threading.Thread(target=self._run_worker, args=(config,), daemon=True)
        self.worker.start()

    def _collect_config(self) -> PlotConfig:
        mode = self.mode_var.get()
        input_file = self.input_file_var.get().strip()
        batch_dir = self.batch_dir_var.get().strip()
        if mode in {"single", "excel"} and not input_file:
            raise ValueError("请选择数据文件。")
        if mode == "batch" and not batch_dir:
            raise ValueError("请选择批量数据文件夹。")

        marker_sizes: list[float] = []
        marker_colors: list[str] = []
        marker_symbols: list[str] = []
        hollow_marker_fill_colors: list[str] = []
        hollow_marker_edge_colors: list[str] = []
        hollow_marker_edge_widths: list[float] = []
        plot_types: list[str] = []
        line_colors: list[str] = []
        line_widths: list[float] = []
        line_styles: list[str] = []
        combine_indices: list[int] = []
        x_logs: list[bool] = []
        y_reverses: list[bool] = []
        linear_fits: list[bool] = []
        axis_width_percents: list[float] = []
        axis_height_percents: list[float] = []

        plot_mode = self._plot_mode_value()
        for index, row in enumerate(self.pair_style_rows, start=1):
            plot_type = self._series_type_value(row["series_type"].get())
            plot_types.append(plot_type)

            if plot_type in {"scatter", "line_symbol"}:
                try:
                    marker_size = float(row["marker_size"].get())
                except ValueError as exc:
                    raise ValueError(f"第 {index} 组的点大小必须是数字。") from exc
            else:
                marker_size = 0.0

            if plot_type in {"line", "line_symbol"}:
                try:
                    line_width = float(row["line_width"].get())
                except ValueError as exc:
                    raise ValueError(f"第 {index} 组的线宽必须是数字。") from exc
            else:
                line_width = 0.0

            if plot_mode == "combined" and bool(row["combine"].get()):
                combine_indices.append(index - 1)

            marker_colors.append(row["marker_color"].get().strip() or "#1F77B4")
            marker_symbol = self._marker_symbol_from_display(row["marker_symbol"].get())
            marker_symbols.append(marker_symbol)
            marker_sizes.append(marker_size)
            hollow_marker_fill_colors.append(row["hollow_fill_color"].get().strip() or "#FFFFFF")
            hollow_marker_edge_colors.append(row["hollow_edge_color"].get().strip() or marker_colors[-1])
            try:
                hollow_marker_edge_widths.append(float(row["hollow_edge_width"].get()))
            except ValueError as exc:
                raise ValueError(f"第 {index} 组的空心圆圆圈线宽必须是数字。") from exc
            line_colors.append(row["line_color"].get().strip() or marker_colors[-1])
            line_widths.append(line_width)
            line_styles.append(self._line_style_value(row["line_style"].get()))
            x_logs.append(bool(row["x_log"].get()))
            y_reverses.append(bool(row["y_reverse"].get()))
            linear_fits.append(bool(row["linear_fit"].get()))
            try:
                axis_width_percents.append(float(row["axis_width"].get()))
                axis_height_percents.append(float(row["axis_height"].get()))
            except ValueError as exc:
                raise ValueError(f"? {index} ??X????Y????????????????") from exc

        if plot_mode == "combined" and self.pair_style_rows and not combine_indices:
            raise ValueError("已选择合并模式，请至少勾选一组需要合并的数据列。")

        try:
            axis_width = float(self.axis_width_var.get())
            axis_height = float(self.axis_height_var.get())
        except ValueError as exc:
            raise ValueError("X????Y????????????????") from exc

        default_plot_type = plot_types[0] if plot_types else "scatter"
        return PlotConfig(
            mode=mode,
            input_file=input_file,
            batch_dir=batch_dir,
            layout=LAYOUT_LABELS[self.layout_display_var.get()],
            plot_mode="separate" if combine_indices else plot_mode,
            marker_size=8.0,
            marker_sizes=marker_sizes,
            marker_symbol="circle",
            marker_color="#1F77B4",
            marker_colors=marker_colors,
            marker_symbols=marker_symbols,
            hollow_marker_fill_color="#FFFFFF",
            hollow_marker_fill_colors=hollow_marker_fill_colors,
            hollow_marker_edge_color="#1F77B4",
            hollow_marker_edge_colors=hollow_marker_edge_colors,
            hollow_marker_edge_width=1.5,
            hollow_marker_edge_widths=hollow_marker_edge_widths,
            cycle_colors=False,
            plot_type=default_plot_type,
            plot_types=plot_types,
            line_color="#1F77B4",
            line_colors=line_colors,
            line_width=2.0,
            line_widths=line_widths,
            line_style="solid",
            line_styles=line_styles,
            combine_indices=combine_indices,
            x_log=False,
            x_logs=x_logs,
            y_reverse=False,
            y_reverses=y_reverses,
            linear_fit=False,
            linear_fits=linear_fits,
            single_origin_project=bool(self.single_origin_project_var.get()),
            axis_width_percent=axis_width,
            axis_height_percent=axis_height,
            axis_width_percents=axis_width_percents,
            axis_height_percents=axis_height_percents,
            figures_dir=self.figures_dir_var.get().strip(),
            origin_dir=self.origin_dir_var.get().strip(),
        )

    def _plot_mode_value(self) -> str:
        return PLOT_MODE_LABELS[self.plot_mode_display_var.get()]

    def _series_type_value(self, value: str) -> str:
        return SERIES_TYPE_LABELS.get(value.strip(), "scatter")

    def _line_style_value(self, value: str) -> str:
        return LINE_STYLE_LABELS.get(value.strip(), "solid")

    def _marker_symbol_from_display(self, value: str) -> str:
        value = value.strip()
        if value in SYMBOL_LABELS:
            return SYMBOL_LABELS[value]
        if value.startswith("Origin 编号 "):
            return value.replace("Origin 编号 ", "").strip()
        if value.isdigit():
            return value
        return value or "circle"

    def _run_worker(self, config: PlotConfig) -> None:
        try:
            run_plotting(config, log=self._log)
        except Exception as exc:
            self._log("")
            self._log(f"运行失败: {exc}")
            self._log(traceback.format_exc())
        finally:
            self.log_queue.put("__WORKER_DONE__")

    def _log(self, message: str) -> None:
        self.log_queue.put(str(message))

    def _drain_logs(self) -> None:
        while True:
            try:
                message = self.log_queue.get_nowait()
            except queue.Empty:
                break
            if message == "__WORKER_DONE__":
                self.start_button.configure(state="normal")
            else:
                self.log_text.insert("end", message + "\n")
                self.log_text.see("end")
        self.after(100, self._drain_logs)


def main() -> None:
    app = OriginScatterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
