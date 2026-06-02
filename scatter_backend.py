from __future__ import annotations

import argparse
import math
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Iterable, Optional, Union

import pandas as pd


LogFunc = Callable[[str], None]
MarkerSymbol = Union[str, int]


@dataclass
class PlotConfig:
    mode: str = "single"
    input_file: str = ""
    batch_dir: str = ""
    layout: str = "auto"
    plot_mode: str = "combined"
    marker_size: float = 8.0
    marker_sizes: list[float] = field(default_factory=list)
    marker_symbol: MarkerSymbol = "circle"
    marker_color: str = "#1F77B4"
    marker_colors: list[str] = field(default_factory=list)
    marker_symbols: list[MarkerSymbol] = field(default_factory=list)
    hollow_marker_fill_color: str = "#FFFFFF"
    hollow_marker_fill_colors: list[str] = field(default_factory=list)
    hollow_marker_edge_color: str = "#1F77B4"
    hollow_marker_edge_colors: list[str] = field(default_factory=list)
    hollow_marker_edge_width: float = 1.5
    hollow_marker_edge_widths: list[float] = field(default_factory=list)
    cycle_colors: bool = False
    plot_type: str = "scatter"
    plot_types: list[str] = field(default_factory=list)
    line_color: str = "#1F77B4"
    line_colors: list[str] = field(default_factory=list)
    line_width: float = 2.0
    line_widths: list[float] = field(default_factory=list)
    line_style: str = "solid"
    line_styles: list[str] = field(default_factory=list)
    combine_indices: list[int] = field(default_factory=list)
    x_log: bool = False
    x_logs: list[bool] = field(default_factory=list)
    y_reverse: bool = False
    y_reverses: list[bool] = field(default_factory=list)
    linear_fit: bool = False
    linear_fits: list[bool] = field(default_factory=list)
    single_origin_project: bool = False
    axis_width_percent: float = 70.0
    axis_height_percent: float = 68.0
    axis_width_percents: list[float] = field(default_factory=list)
    axis_height_percents: list[float] = field(default_factory=list)
    figures_dir: str = ""
    origin_dir: str = ""


@dataclass
class PlotPair:
    x_col: int
    y_col: int
    x_title: str
    y_title: str
    output_suffix: str
    index: int = 0


@dataclass
class AxisLimits:
    x_min: float
    x_max: float
    x_step: float
    x_decimals: int
    y_min: float
    y_max: float
    y_step: float
    y_decimals: int


@dataclass
class LinearFitResult:
    slope: float
    intercept: float
    r_squared: float
    p_value: float
    x_min: float
    x_max: float


@dataclass
class AxisProfile:
    x_key: str
    y_key: str
    x_log: bool
    y_reverse: bool
    x_order: Optional[int]
    y_order: Optional[int]


SUPPORTED_DATA_EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".dat")
COLORS_FOR_MULTIPLE_PLOTS = [
    "#1F77B4",
    "#D73027",
    "#4DAF4A",
    "#984EA3",
    "#FF7F00",
    "#A65628",
    "#F781BF",
    "#636363",
    "#66C2A5",
    "#FC8D62",
]
SYMBOL_MAP = {
    "none": 0,
    "square": 1,
    "circle": 2,
    "hollow_circle": 2,
    "triangle_up": 3,
    "triangle": 3,
    "triangle_down": 4,
    "diamond": 5,
    "plus": 6,
    "cross": 7,
    "star": 8,
    "hexagon": 9,
    "inverted_triangle": 4,
    "left_triangle": 10,
    "right_triangle": 11,
    "pentagon": 12,
    "thin_plus": 13,
    "thin_cross": 14,
    "asterisk": 15,
    "horizontal_line": 16,
    "vertical_line": 17,
}
PLOT_TYPE_MAP = {
    "scatter": "s",
    "line_symbol": "y",
    "line": "l",
}
LINE_STYLE_MAP = {
    "solid": 1,
    "dash": 2,
    "dot": 3,
    "dash_dot": 4,
    "dash_dot_dot": 5,
}

FONT_NAME = "Times New Roman"
FONT_SIZE = 20
AXIS_WIDTH = 2
TICK_LENGTH = 5
SHOW_TOP_RIGHT_AXES = True
SHOW_LEGEND = True
AXIS_PADDING_RATIO = 0.06
AUTO_TICK_TARGET_INTERVALS = 5


class OriginConnectionError(RuntimeError):
    pass


def run_plotting(config: PlotConfig, log: Optional[LogFunc] = None) -> tuple[list[str], list[tuple[str, str]]]:
    logger = log or print
    config = normalize_config(config)
    figures_dir = ensure_dir(config.figures_dir, "figures")
    origin_dir = ensure_dir(config.origin_dir, "origin")

    op = connect_origin(logger)
    success: list[str] = []
    failed: list[tuple[str, str]] = []

    logger(f"运行模式: {config.mode}")
    logger(f"数据结构: {config.layout}")
    logger(f"绘图方式: {config.plot_mode}")
    if config.marker_colors:
        logger(f"每组颜色列表: {', '.join(config.marker_colors)}")
    if config.marker_symbols:
        logger(f"每组散点类型列表: {', '.join(str(item) for item in config.marker_symbols)}")
    if config.marker_sizes:
        logger(f"每组散点大小列表: {', '.join(str(item) for item in config.marker_sizes)}")
    logger(f"默认X轴对数: {'开启' if config.x_log else '关闭'}")
    logger(f"默认Y轴倒序: {'开启' if config.y_reverse else '关闭'}")
    if config.x_logs:
        logger(f"每组X轴对数: {', '.join('开启' if item else '关闭' for item in config.x_logs)}")
    if config.y_reverses:
        logger(f"每组Y轴倒序: {', '.join('开启' if item else '关闭' for item in config.y_reverses)}")
    if config.combine_indices:
        logger(f"合并出图数据组: {', '.join(str(index + 1) for index in config.combine_indices)}")
    logger(f"默认线性拟合: {'开启' if config.linear_fit else '关闭'}")
    if config.linear_fits:
        fit_labels = ["拟合" if item else "不拟合" for item in config.linear_fits]
        logger(f"每组线性拟合设置: {', '.join(fit_labels)}")
    logger(f"EMF 输出文件夹: {figures_dir}")
    logger(f"OPJU 输出文件夹: {origin_dir}")
    if config.single_origin_project:
        logger("Origin 保存方式: 所有图页保存到同一个 OPJU 项目")
        op.new()
    logger(f"默认坐标轴长度: X轴 {config.axis_width_percent:g}%, Y轴 {config.axis_height_percent:g}%")
    if config.axis_width_percents or config.axis_height_percents:
        logger("已启用每组独立坐标轴长度设置。")

    for data_path, sheet_name, df in iter_datasets(config, logger):
        label = make_dataset_label(data_path, sheet_name)
        logger("")
        logger("=" * 70)
        logger(f"处理数据: {label}")
        try:
            pairs = infer_plot_pairs(df, config.layout, logger)
            combine_index_set = set(config.combine_indices)
            if combine_index_set:
                selected_pairs = [pair for index, pair in enumerate(pairs) if index in combine_index_set]
                if selected_pairs:
                    output_stem = sanitize_filename(f"{label}__combined", fallback="scatter_combined")
                    plot_combined_dataset(op, df, selected_pairs, output_stem, figures_dir, origin_dir, config, logger)
                    success.append(output_stem)
                for index, pair in enumerate(pairs):
                    if index in combine_index_set:
                        continue
                    output_stem = sanitize_filename(f"{label}__{pair.output_suffix}", fallback="scatter")
                    plot_one_pair(
                        op=op,
                        df=df,
                        pair=pair,
                        pair_index=index,
                        output_stem=output_stem,
                        figures_dir=figures_dir,
                        origin_dir=origin_dir,
                        config=config,
                        logger=logger,
                    )
                    success.append(output_stem)
            elif config.plot_mode == "combined":
                output_stem = sanitize_filename(label, fallback="scatter")
                plot_combined_dataset(op, df, pairs, output_stem, figures_dir, origin_dir, config, logger)
                success.append(output_stem)
            else:
                for index, pair in enumerate(pairs):
                    output_stem = sanitize_filename(f"{label}__{pair.output_suffix}", fallback="scatter")
                    plot_one_pair(
                        op=op,
                        df=df,
                        pair=pair,
                        pair_index=index,
                        output_stem=output_stem,
                        figures_dir=figures_dir,
                        origin_dir=origin_dir,
                        config=config,
                        logger=logger,
                    )
                    success.append(output_stem)
        except Exception as exc:
            failed.append((label, str(exc)))
            logger(f"失败: {label}")
            logger(f"错误: {exc}")

    if config.single_origin_project and success:
        project_path = save_single_origin_project(op, origin_dir, logger)
        logger(f"统一 Origin 项目: {project_path}")

    print_summary(success, failed, logger)
    return success, failed


def normalize_config(config: PlotConfig) -> PlotConfig:
    mode = config.mode.strip().lower()
    if mode == "sheets":
        mode = "excel"
    if mode not in {"auto", "single", "batch", "excel"}:
        raise ValueError("mode 必须是 auto、single、batch、excel 或 sheets")
    if config.layout not in {"auto", "pairs", "common-x"}:
        raise ValueError("layout 必须是 auto、pairs 或 common-x")
    if config.plot_mode not in {"combined", "separate"}:
        raise ValueError("plot_mode 必须是 combined 或 separate")
    plot_type = normalize_plot_type(config.plot_type)
    plot_types = [normalize_plot_type(item) for item in config.plot_types]
    line_style = normalize_line_style(config.line_style)
    line_styles = [normalize_line_style(item) for item in config.line_styles]
    axis_width_percent = clamp_float(float(config.axis_width_percent), 20.0, 90.0)
    axis_height_percent = clamp_float(float(config.axis_height_percent), 20.0, 85.0)
    axis_width_percents = [clamp_float(float(item), 20.0, 90.0) for item in config.axis_width_percents]
    axis_height_percents = [clamp_float(float(item), 20.0, 85.0) for item in config.axis_height_percents]

    if mode in {"single", "excel"} and not config.input_file:
        raise ValueError("single/excel 模式必须选择 input_file")
    if mode == "batch" and not config.batch_dir:
        raise ValueError("batch 模式必须选择 batch_dir")

    return PlotConfig(
        mode=mode,
        input_file=config.input_file,
        batch_dir=config.batch_dir,
        layout=config.layout,
        plot_mode=config.plot_mode,
        marker_size=float(config.marker_size),
        marker_sizes=[float(item) for item in config.marker_sizes],
        marker_symbol=config.marker_symbol,
        marker_color=config.marker_color,
        marker_colors=list(config.marker_colors),
        marker_symbols=list(config.marker_symbols),
        hollow_marker_fill_color=config.hollow_marker_fill_color,
        hollow_marker_fill_colors=list(config.hollow_marker_fill_colors),
        hollow_marker_edge_color=config.hollow_marker_edge_color,
        hollow_marker_edge_colors=list(config.hollow_marker_edge_colors),
        hollow_marker_edge_width=float(config.hollow_marker_edge_width),
        hollow_marker_edge_widths=[float(item) for item in config.hollow_marker_edge_widths],
        cycle_colors=bool(config.cycle_colors),
        plot_type=plot_type,
        plot_types=plot_types,
        line_color=config.line_color,
        line_colors=list(config.line_colors),
        line_width=float(config.line_width),
        line_widths=[float(item) for item in config.line_widths],
        line_style=line_style,
        line_styles=line_styles,
        combine_indices=[int(item) for item in config.combine_indices],
        x_log=bool(config.x_log),
        x_logs=[bool(item) for item in config.x_logs],
        y_reverse=bool(config.y_reverse),
        y_reverses=[bool(item) for item in config.y_reverses],
        linear_fit=bool(config.linear_fit),
        linear_fits=[bool(item) for item in config.linear_fits],
        single_origin_project=bool(config.single_origin_project),
        axis_width_percent=axis_width_percent,
        axis_height_percent=axis_height_percent,
        axis_width_percents=axis_width_percents,
        axis_height_percents=axis_height_percents,
        figures_dir=config.figures_dir,
        origin_dir=config.origin_dir,
    )


def connect_origin(logger: LogFunc):
    logger("正在连接 Origin 2023b ...")
    try:
        import originpro as op
    except ImportError as exc:
        raise OriginConnectionError(
            "无法导入 originpro。请先安装依赖: pip install -r requirements.txt"
        ) from exc

    try:
        if getattr(op, "oext", False):
            logger("检测到 Origin 已在运行，将使用当前 Origin 实例。")
        elif hasattr(op, "set_origin_instance"):
            logger("正在启动/连接 Origin，请稍候 ...")
            op.set_origin_instance()
        if hasattr(op, "set_show"):
            op.set_show(True)
    except Exception as exc:
        raise OriginConnectionError(
            "无法连接 Origin。请确认:\n"
            "1. 当前电脑已安装并可正常启动 Origin 2023b / OriginPro 2023b；\n"
            "2. 当前 Python 环境已安装 originpro；\n"
            "3. Origin 2023b 已正确注册为 COM/Automation 版本；\n"
            "4. 程序在 Windows 本地环境运行，而不是 WSL 或远程 Linux。"
        ) from exc
    return op


def iter_datasets(config: PlotConfig, logger: LogFunc) -> Iterable[tuple[Path, Optional[str], pd.DataFrame]]:
    mode = config.mode
    if mode == "auto":
        batch_files = find_data_files(Path(config.batch_dir) if config.batch_dir else Path.cwd() / "data" / "batch")
        mode = "batch" if batch_files else "single"
        logger(f"auto 模式选择为: {mode}")

    if mode == "single":
        data_path = Path(config.input_file) if config.input_file else find_default_single_file()
        yield data_path, None, read_scatter_data(data_path, None, logger)
    elif mode == "batch":
        batch_dir = Path(config.batch_dir)
        data_files = find_data_files(batch_dir)
        if not data_files:
            raise FileNotFoundError(f"未在批量目录中找到数据文件: {batch_dir}")
        logger(f"批量绘图: 找到 {len(data_files)} 个数据文件。")
        for data_path in data_files:
            yield data_path, None, read_scatter_data(data_path, None, logger)
    elif mode == "excel":
        workbook_path = Path(config.input_file)
        if workbook_path.suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
            raise ValueError("Excel 多 Sheet 模式仅适用于 .xlsx/.xlsm/.xls 文件")
        sheet_names = pd.ExcelFile(workbook_path).sheet_names
        if not sheet_names:
            raise ValueError(f"Excel 文件中没有可读取的 sheet: {workbook_path}")
        logger(f"Excel 多 Sheet 模式: 找到 {len(sheet_names)} 个 sheet。")
        for sheet_name in sheet_names:
            yield workbook_path, sheet_name, read_scatter_data(workbook_path, sheet_name, logger)


def ensure_dir(path_value: str, default_name: str) -> Path:
    path = Path(path_value).expanduser() if path_value else Path.cwd() / default_name
    path.mkdir(parents=True, exist_ok=True)
    return path.resolve()


def clamp_float(value: float, minimum: float, maximum: float) -> float:
    if not math.isfinite(value):
        return minimum
    return max(minimum, min(maximum, value))


def find_default_single_file() -> Path:
    data_dir = Path.cwd() / "data"
    preferred = ["1.xlsx", "1.xlsm", "1.xls", "1.csv", "1.txt", "1.dat"]
    for filename in preferred:
        candidate = data_dir / filename
        if candidate.exists():
            return candidate
    files = find_data_files(data_dir)
    if files:
        return files[0]
    return data_dir / "1.xlsx"


def find_data_files(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    files: list[Path] = []
    for ext in SUPPORTED_DATA_EXTENSIONS:
        files.extend(folder.glob(f"*{ext}"))
    return sorted(p for p in files if p.is_file() and not p.name.startswith("~$"))


def read_scatter_data(data_path: Path, sheet_name: Optional[Union[str, int]], logger: LogFunc) -> pd.DataFrame:
    if not data_path.exists():
        raise FileNotFoundError(f"未找到数据文件: {data_path}")

    suffix = data_path.suffix.lower()
    if suffix in (".xlsx", ".xlsm", ".xls"):
        effective_sheet = 0 if sheet_name is None else sheet_name
        logger(f"读取 Excel 数据: {data_path.name}, sheet={effective_sheet}")
        df = pd.read_excel(data_path, sheet_name=effective_sheet, header=0)
    elif suffix in (".csv", ".txt", ".dat"):
        logger(f"读取文本数据: {data_path.name}")
        try:
            df = pd.read_csv(data_path, sep=None, engine="python", header=0)
        except Exception:
            df = pd.read_csv(data_path, sep=r"\s+", engine="python", header=0)
        if df.shape[1] < 2:
            df = pd.read_csv(data_path, sep=r"\s+", engine="python", header=0)
    else:
        raise ValueError(f"不支持的数据文件格式: {data_path.suffix}")

    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    if df.shape[1] < 2:
        raise ValueError("数据至少需要 2 列: X/Y")
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(axis=0, how="all")
    if df.empty:
        raise ValueError("清理空值后没有可绘制的数据。")

    logger(f"数据规模: {df.shape[0]} 行 x {df.shape[1]} 列")
    return df


def infer_plot_pairs(df: pd.DataFrame, layout: str, logger: LogFunc) -> list[PlotPair]:
    n_cols = df.shape[1]
    if n_cols < 2:
        raise ValueError("数据至少需要 2 列。")
    if layout == "auto":
        layout = "pairs" if n_cols % 2 == 0 else "common-x"

    columns = list(df.columns)
    pairs: list[PlotPair] = []
    if layout == "pairs":
        if n_cols % 2 != 0:
            raise ValueError("pairs 布局要求偶数列。请改用 common-x，或整理为 X1/Y1、X2/Y2 ...")
        for i in range(0, n_cols, 2):
            idx = len(pairs) + 1
            x_title = clean_column_title(columns[i], f"X{idx}")
            y_title = clean_column_title(columns[i + 1], f"Y{idx}")
            suffix = sanitize_filename(f"{idx}_{y_title}", fallback=f"plot_{idx}")
            pairs.append(PlotPair(i, i + 1, x_title, y_title, suffix, idx - 1))
    elif layout == "common-x":
        x_title = clean_column_title(columns[0], "X")
        for i in range(1, n_cols):
            idx = len(pairs) + 1
            y_title = clean_column_title(columns[i], f"Y{idx}")
            suffix = sanitize_filename(f"{idx}_{y_title}", fallback=f"plot_{idx}")
            pairs.append(PlotPair(0, i, x_title, y_title, suffix, idx - 1))
    else:
        raise ValueError("layout 只能是 auto、pairs 或 common-x")

    logger(f"识别到二维散点图数量: {len(pairs)}")
    for idx, pair in enumerate(pairs, start=1):
        valid_points = df.iloc[:, [pair.x_col, pair.y_col]].dropna(how="any").shape[0]
        logger(f"  图 {idx}: X='{pair.x_title}', Y='{pair.y_title}', 有效点数={valid_points}")
    return pairs


def compute_axis_limits(
    df: pd.DataFrame,
    pairs: list[PlotPair],
    logger: Optional[LogFunc] = None,
    x_log: bool = False,
) -> AxisLimits:
    x_values = []
    y_values = []
    for pair in pairs:
        xy = df.iloc[:, [pair.x_col, pair.y_col]].dropna(how="any")
        if not xy.empty:
            pair_x_values = pd.to_numeric(xy.iloc[:, 0], errors="coerce").dropna()
            if x_log:
                pair_x_values = pair_x_values[pair_x_values > 0]
            x_values.extend(pair_x_values.tolist())
            y_values.extend(pd.to_numeric(xy.iloc[:, 1], errors="coerce").dropna().tolist())
    if not x_values or not y_values:
        raise ValueError("没有有效 X-Y 数据。")
    if x_log and any(float(value) <= 0 for value in x_values):
        raise ValueError("X 轴使用对数时，X 数据必须大于 0。")

    raw_x_min, raw_x_max = float(min(x_values)), float(max(x_values))
    raw_y_min, raw_y_max = float(min(y_values)), float(max(y_values))
    if raw_x_min == raw_x_max:
        raw_x_min -= 0.5
        raw_x_max += 0.5
    if raw_y_min == raw_y_max:
        raw_y_min -= 0.5
        raw_y_max += 0.5

    x_span = raw_x_max - raw_x_min
    y_span = raw_y_max - raw_y_min
    x_pad = x_span * AXIS_PADDING_RATIO
    y_pad = y_span * AXIS_PADDING_RATIO
    x_step = nice_number((x_span + 2 * x_pad) / AUTO_TICK_TARGET_INTERVALS)
    y_step = nice_number((y_span + 2 * y_pad) / AUTO_TICK_TARGET_INTERVALS)
    limits = AxisLimits(
        x_min=nice_floor(raw_x_min - x_pad, x_step),
        x_max=nice_ceil(raw_x_max + x_pad, x_step),
        x_step=x_step,
        x_decimals=infer_decimals(x_step),
        y_min=nice_floor(raw_y_min - y_pad, y_step),
        y_max=nice_ceil(raw_y_max + y_pad, y_step),
        y_step=y_step,
        y_decimals=infer_decimals(y_step),
    )

    if logger:
        logger(f"X 绘图范围: {limits.x_min:g} - {limits.x_max:g}, 主刻度步长 {limits.x_step:g}")
        logger(f"Y 绘图范围: {limits.y_min:g} - {limits.y_max:g}, 主刻度步长 {limits.y_step:g}")
    return limits



def group_pairs_for_combined_axes(
    df: pd.DataFrame,
    pairs: list[PlotPair],
    config: PlotConfig,
    logger: LogFunc,
) -> list[list[PlotPair]]:
    if len(pairs) <= 1:
        return [pairs] if pairs else []

    groups: list[list[PlotPair]] = []
    profiles: list[AxisProfile] = []
    for pair in pairs:
        profile = build_axis_profile(df, pair, config)
        placed = False
        for group_index, group_profile in enumerate(profiles):
            if axis_profiles_compatible(profile, group_profile):
                groups[group_index].append(pair)
                placed = True
                break
        if not placed:
            profiles.append(profile)
            groups.append([pair])

    if len(groups) > 1:
        logger("合并出图时检测到坐标轴标签、单位、对数/倒序设置或数值量级不一致。")
        logger("将保持所有数据合并在同一个绘图区，并自动创建 Origin Multi-Axis 多坐标轴图层。")
        for index, group in enumerate(groups, start=1):
            names = ", ".join(f"{pair.x_title} / {pair.y_title}" for pair in group)
            logger(f"  坐标轴组 {index}: {names}")
    return groups


def build_axis_profile(df: pd.DataFrame, pair: PlotPair, config: PlotConfig) -> AxisProfile:
    x_values, y_values = get_pair_numeric_values(df, pair, select_x_log(config, pair.index))
    return AxisProfile(
        x_key=axis_title_key(pair.x_title),
        y_key=axis_title_key(pair.y_title),
        x_log=select_x_log(config, pair.index),
        y_reverse=select_y_reverse(config, pair.index),
        x_order=value_order(x_values),
        y_order=value_order(y_values),
    )


def axis_profiles_compatible(left: AxisProfile, right: AxisProfile) -> bool:
    if left.x_key != right.x_key or left.y_key != right.y_key:
        return False
    if left.x_log != right.x_log or left.y_reverse != right.y_reverse:
        return False
    return orders_compatible(left.x_order, right.x_order) and orders_compatible(left.y_order, right.y_order)


def axis_title_key(title: str) -> str:
    text = normalize_axis_text(title)
    unit = extract_axis_unit(text)
    return unit or text


def normalize_axis_text(title: str) -> str:
    text = str(title).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def extract_axis_unit(title: str, normalize: bool = True) -> str:
    matches = re.findall(r"[\(\[（【]\s*([^\)\]）】]+?)\s*[\)\]）】]", title)
    if matches:
        unit = str(matches[-1]).strip()
        return normalize_axis_text(unit) if normalize else unit
    return ""


def get_pair_numeric_values(df: pd.DataFrame, pair: PlotPair, x_log: bool) -> tuple[list[float], list[float]]:
    xy = df.iloc[:, [pair.x_col, pair.y_col]].dropna(how="any")
    if xy.empty:
        return [], []
    x_series = pd.to_numeric(xy.iloc[:, 0], errors="coerce").dropna()
    if x_log:
        x_series = x_series[x_series > 0]
    y_series = pd.to_numeric(xy.iloc[:, 1], errors="coerce").dropna()
    return x_series.astype(float).tolist(), y_series.astype(float).tolist()


def value_order(values: list[float]) -> Optional[int]:
    finite_values = [abs(float(value)) for value in values if math.isfinite(float(value)) and float(value) != 0]
    if not finite_values:
        return None
    reference = max(finite_values)
    if reference <= 0:
        return None
    return int(math.floor(math.log10(reference)))


def orders_compatible(left: Optional[int], right: Optional[int]) -> bool:
    if left is None or right is None:
        return True
    return abs(left - right) <= 2

def plot_one_pair(
    op,
    df: pd.DataFrame,
    pair: PlotPair,
    pair_index: int,
    output_stem: str,
    figures_dir: Path,
    origin_dir: Path,
    config: PlotConfig,
    logger: LogFunc,
) -> tuple[Path, Path]:
    logger(f"创建单图: {output_stem}")
    if not config.single_origin_project:
        op.new()
    xy = df.iloc[:, [pair.x_col, pair.y_col]].dropna(how="any").copy()
    xy.columns = [pair.x_title, pair.y_title]
    fit = compute_linear_fit(xy, pair.x_title, pair.y_title, logger) if select_linear_fit(config, pair_index) else None
    wks = op.new_sheet("w", lname="Scatter_Data")
    wks.from_df(xy)
    try:
        wks.set_labels([clean_column_title(col, f"C{i + 1}") for i, col in enumerate(xy.columns)], "L")
    except Exception:
        pass

    gp = op.new_graph(template="scatter")
    gp.activate()
    lock_graph_element_scale(gp)
    gl = gp[0]
    gl.activate()
    sync_overlay_layer(gl, 0, config, pair_index)
    marker_color = select_marker_color(config, pair_index)
    plot_type = select_plot_type(config, pair_index)
    plot = gl.add_plot(wks, colx=0, coly=1, type=origin_plot_type(plot_type))
    apply_series_style(
        plot=plot,
        plot_type=plot_type,
        marker_size=select_marker_size(config, pair_index),
        marker_symbol=select_marker_symbol(config, pair_index),
        marker_color=marker_color,
        hollow_fill_color=select_hollow_marker_fill_color(config, pair_index),
        hollow_edge_color=select_hollow_marker_edge_color(config, pair_index, marker_color),
        hollow_edge_width=select_hollow_marker_edge_width(config, pair_index),
        line_color=select_line_color(config, pair_index, marker_color),
        line_width=select_line_width(config, pair_index),
        line_style=select_line_style(config, pair_index),
    )
    x_log = select_x_log(config, pair_index)
    y_reverse = select_y_reverse(config, pair_index)
    axis_limits = compute_axis_limits(df, [pair], logger, x_log=x_log)
    format_axes(gp, gl, pair.x_title, pair.y_title, axis_limits, x_log=x_log, y_reverse=y_reverse)
    apply_legend(gl, [pair.y_title])
    if fit:
        fit_wks = create_fit_worksheet(op, [(fit, marker_color, pair.x_title, pair.y_title)])
        fit_plot = gl.add_plot(fit_wks, colx=0, coly=1, type="l")
        apply_fit_plot_style(fit_plot, marker_color)
        add_fit_annotations(gl, [(fit, marker_color, pair.x_title, pair.y_title)], axis_limits)
    return export_outputs(gp, figures_dir, origin_dir, output_stem, logger, save_origin=not config.single_origin_project)


def plot_combined_dataset(
    op,
    df: pd.DataFrame,
    pairs: list[PlotPair],
    output_stem: str,
    figures_dir: Path,
    origin_dir: Path,
    config: PlotConfig,
    logger: LogFunc,
) -> tuple[Path, Path]:
    logger(f"创建合并图: {output_stem}")
    if not config.single_origin_project:
        op.new()
    wks = op.new_sheet("w", lname="Scatter_Data")
    wks.from_df(df)
    try:
        wks.set_labels([clean_column_title(col, f"C{i + 1}") for i, col in enumerate(df.columns)], "L")
    except Exception:
        pass

    axis_groups = group_pairs_for_combined_axes(df, pairs, config, logger)
    multi_axis_overlay = len(axis_groups) > 1
    if multi_axis_overlay:
        logger("检测到合并数据的坐标轴范围、单位或轴设置不一致，自动使用 Origin Multi-Axis 叠加坐标轴模式。")
        logger("所有数据仍合并在同一个绘图区内，各坐标轴分组保留独立范围、轴标题、单位和刻度样式。")
    gp = op.new_graph(template="scatter")
    gp.activate()
    lock_graph_element_scale(gp)

    for group_index, group_pairs in enumerate(axis_groups):
        if not group_pairs:
            continue
        if group_index == 0:
            gl = gp[0]
        else:
            gl = add_multi_axis_layer(gp) if multi_axis_overlay else gp.add_layer(0)
        if multi_axis_overlay:
            sync_multi_axis_layer(gl, group_index, len(axis_groups), config, group_pairs[0].index)
        else:
            sync_overlay_layer(gl, group_index, config, group_pairs[0].index)
        gl.activate()

        fit_specs: list[tuple[LinearFitResult, str, str, str]] = []
        for pair in group_pairs:
            index = pair.index
            marker_color = select_marker_color(config, index)
            plot_type = select_plot_type(config, index)
            plot = gl.add_plot(wks, colx=pair.x_col, coly=pair.y_col, type=origin_plot_type(plot_type))
            apply_series_style(
                plot=plot,
                plot_type=plot_type,
                marker_size=select_marker_size(config, index),
                marker_symbol=select_marker_symbol(config, index),
                marker_color=marker_color,
                hollow_fill_color=select_hollow_marker_fill_color(config, index),
                hollow_edge_color=select_hollow_marker_edge_color(config, index, marker_color),
                hollow_edge_width=select_hollow_marker_edge_width(config, index),
                line_color=select_line_color(config, index, marker_color),
                line_width=select_line_width(config, index),
                line_style=select_line_style(config, index),
            )
            if select_linear_fit(config, index):
                xy = df.iloc[:, [pair.x_col, pair.y_col]].dropna(how="any").copy()
                xy.columns = [pair.x_title, pair.y_title]
                fit = compute_linear_fit(xy, pair.x_title, pair.y_title, logger)
                if fit:
                    fit_specs.append((fit, marker_color, pair.x_title, pair.y_title))

        x_title = combined_axis_title([pair.x_title for pair in group_pairs], fallback="X")
        y_title = combined_axis_title([pair.y_title for pair in group_pairs], fallback="Y")
        x_log = select_combined_x_log(config, group_pairs, logger)
        y_reverse = select_combined_y_reverse(config, group_pairs, logger)
        axis_limits = compute_axis_limits(df, group_pairs, logger, x_log=x_log)
        y_axis_side = "left" if group_index == 0 else "right"
        x_axis_side = "bottom" if group_index == 0 else "top"
        format_axes(
            gp,
            gl,
            x_title,
            y_title,
            axis_limits,
            x_log=x_log,
            y_reverse=y_reverse,
            y_axis_side=y_axis_side,
            x_axis_side=x_axis_side,
            overlay_index=group_index,
        )
        apply_legend(gl, [pair.y_title for pair in group_pairs])
        if fit_specs:
            fit_wks = create_fit_worksheet(op, fit_specs)
            for fit_index, (_fit, color, _x_title, _y_title) in enumerate(fit_specs):
                x_col = fit_index * 2
                fit_plot = gl.add_plot(fit_wks, colx=x_col, coly=x_col + 1, type="l")
                apply_fit_plot_style(fit_plot, color)
            add_fit_annotations(gl, fit_specs, axis_limits)

    return export_outputs(gp, figures_dir, origin_dir, output_stem, logger, save_origin=not config.single_origin_project)


def sync_overlay_layer(gl, overlay_index: int, config: PlotConfig, pair_index: int = 0) -> None:
    try:
        left_margin = 15.0
        top_margin = 10.0
        right_margin = 12.0
        bottom_margin = 17.0
        width = min(select_axis_width_percent(config, pair_index), 100.0 - left_margin - right_margin)
        height = min(select_axis_height_percent(config, pair_index), 100.0 - top_margin - bottom_margin)
        width = max(25.0, width)
        height = max(25.0, height)

        gl.lt_exec("layer.unit = 1")
        gl.lt_exec(f"layer.left = {left_margin:.3f}")
        gl.lt_exec(f"layer.top = {top_margin:.3f}")
        gl.lt_exec(f"layer.width = {width:.3f}")
        gl.lt_exec(f"layer.height = {height:.3f}")
        gl.lt_exec("doc -uw")
    except Exception:
        pass


def sync_multi_axis_layer(
    gl,
    axis_index: int,
    axis_count: int,
    config: PlotConfig,
    pair_index: int = 0,
) -> None:
    """Overlay incompatible axes in one plot rectangle without crossing page bounds."""
    try:
        count = max(1, axis_count)
        right_margin = min(28.0, 10.0 + max(0, count - 1) * 5.0)
        top_margin = 12.0 if count > 1 else 10.0
        bottom_margin = 17.0
        left_margin = 15.0

        max_width = max(35.0, 100.0 - left_margin - right_margin)
        max_height = max(35.0, 100.0 - top_margin - bottom_margin)
        width = min(select_axis_width_percent(config, pair_index), max_width)
        height = min(select_axis_height_percent(config, pair_index), max_height)

        gl.lt_exec("layer.unit = 1")
        gl.lt_exec(f"layer.left = {left_margin:.3f}")
        gl.lt_exec(f"layer.top = {top_margin:.3f}")
        gl.lt_exec(f"layer.width = {width:.3f}")
        gl.lt_exec(f"layer.height = {height:.3f}")
        gl.lt_exec("doc -uw")
    except Exception:
        pass


def add_multi_axis_layer(gp):
    """Create an Origin multi-axis layer with independent X/Y scales."""
    gp.activate()
    try:
        gp[0].activate()
    except Exception:
        pass
    try:
        gp.lt_exec("layadd type:=txry offset:=1 activate:=1")
        lock_graph_element_scale(gp)
        return gp[len(gp) - 1]
    except Exception:
        gl = gp.add_layer(4)
        lock_graph_element_scale(gp)
        return gl


def lock_graph_element_scale(gp) -> None:
    try:
        gp.activate()
        gp.lt_exec("page -afu1")
        gp.lt_exec("doc -uw")
    except Exception:
        pass


def combined_axis_title(titles: list[str], fallback: str) -> str:
    cleaned: list[str] = []
    for title in titles:
        text = str(title).strip()
        if text and text not in cleaned:
            cleaned.append(text)
    if not cleaned:
        return fallback
    if len(cleaned) == 1:
        return cleaned[0]
    units = []
    for title in cleaned:
        unit = extract_axis_unit(title, normalize=False)
        if unit and unit not in units:
            units.append(unit)
    if len(units) == 1:
        names = [re.sub(r"\s*[\(\[（【]\s*[^\)\]）】]+?\s*[\)\]）】]\s*$", "", title).strip() for title in cleaned]
        return f"{' / '.join(name for name in names if name)} ({units[0]})"
    return " / ".join(cleaned)


def apply_series_style(
    plot,
    plot_type: str,
    marker_size: float,
    marker_symbol: MarkerSymbol,
    marker_color: str,
    hollow_fill_color: str,
    hollow_edge_color: str,
    hollow_edge_width: float,
    line_color: str,
    line_width: float,
    line_style: str,
) -> None:
    symbol_id = normalize_marker_symbol(marker_symbol)
    line_style_id = LINE_STYLE_MAP[normalize_line_style(line_style)]
    has_symbol = plot_type in {"scatter", "line_symbol"}
    has_line = plot_type in {"line", "line_symbol"}
    is_hollow_circle = normalize_marker_symbol_name(marker_symbol) == "hollow_circle"
    symbol_color = hollow_edge_color if is_hollow_circle else marker_color

    try:
        plot.color = line_color if has_line else symbol_color
    except Exception:
        pass
    for prop in ("symbol.size", "sym.size", "size"):
        if safe_plot_set(plot, "set_float", prop, float(marker_size) if has_symbol else 0.0):
            break
    for prop in ("symbol.kind", "symbol.type", "symbol.shape", "sym.kind", "sym.shape"):
        if safe_plot_set(plot, "set_int", prop, int(symbol_id) if has_symbol else 0):
            break
    for prop in ("line.width", "line.thickness", "width"):
        if safe_plot_set(plot, "set_float", prop, float(line_width) if has_line else 0.0):
            break
    for prop in ("line.style", "line.type", "linestyle"):
        if safe_plot_set(plot, "set_int", prop, int(line_style_id)):
            break
    if has_symbol and is_hollow_circle:
        apply_hollow_marker_style(plot, hollow_fill_color, hollow_edge_color, hollow_edge_width)
    try:
        plot.symbol_kind = int(symbol_id) if has_symbol else 0
        plot.symbol_size = float(marker_size) if has_symbol else 0.0
    except Exception:
        pass
    try:
        plot.set_cmd(
            f"-k {int(symbol_id) if has_symbol else 0}",
            f"-z {float(marker_size) if has_symbol else 0}",
            f"-l {1 if has_line else 0}",
            f"-w {float(line_width) if has_line else 0}",
            f"-d {int(line_style_id)}",
        )
    except Exception:
        pass


def apply_marker_style(gl, plot, marker_size: float, marker_symbol: MarkerSymbol, marker_color: str) -> None:
    apply_series_style(
        plot=plot,
        plot_type="scatter",
        marker_size=marker_size,
        marker_symbol=marker_symbol,
        marker_color=marker_color,
        hollow_fill_color="#FFFFFF",
        hollow_edge_color=marker_color,
        hollow_edge_width=1.5,
        line_color=marker_color,
        line_width=0,
        line_style="solid",
    )


def apply_hollow_marker_style(plot, fill_color: str, edge_color: str, edge_width: float) -> None:
    fill_value = origin_color_value(fill_color)
    edge_value = origin_color_value(edge_color)
    for method_name, prop, value in (
        ("set_int", "symbol.fill.color", fill_value),
        ("set_int", "symbol.fillcolor", fill_value),
        ("set_int", "symbol.interior.color", fill_value),
        ("set_int", "sym.fill.color", fill_value),
        ("set_int", "symbol.edge.color", edge_value),
        ("set_int", "symbol.edgecolor", edge_value),
        ("set_int", "symbol.border.color", edge_value),
        ("set_int", "sym.edge.color", edge_value),
        ("set_float", "symbol.edge.width", float(edge_width)),
        ("set_float", "symbol.border.width", float(edge_width)),
        ("set_float", "sym.edge.width", float(edge_width)),
    ):
        safe_plot_set(plot, method_name, prop, value)
    try:
        plot.set_cmd(
            f"-c {edge_value}",
            f"-cf {fill_value}",
        )
    except Exception:
        pass


def create_fit_worksheet(op, fit_specs: list[tuple[LinearFitResult, str, str, str]]):
    columns: dict[str, pd.Series] = {}
    for index, (fit, _color, x_title, y_title) in enumerate(fit_specs, start=1):
        x_key = f"FitX_{index}_{x_title}"
        y_key = f"FitY_{index}_{y_title}"
        slope_key = f"Slope_{index}_{y_title}"
        intercept_key = f"Intercept_{index}_{y_title}"
        r2_key = f"R2_{index}_{y_title}"
        p_key = f"PValue_{index}_{y_title}"
        columns[x_key] = pd.Series([fit.x_min, fit.x_max])
        columns[y_key] = pd.Series([
            fit.slope * fit.x_min + fit.intercept,
            fit.slope * fit.x_max + fit.intercept,
        ])
        columns[slope_key] = pd.Series([fit.slope, None])
        columns[intercept_key] = pd.Series([fit.intercept, None])
        columns[r2_key] = pd.Series([fit.r_squared, None])
        columns[p_key] = pd.Series([fit.p_value, None])
    fit_df = pd.DataFrame(columns)
    fit_wks = op.new_sheet("w", lname="Linear_Fit_Data")
    fit_wks.from_df(fit_df)
    try:
        fit_wks.set_labels([clean_column_title(col, f"C{i + 1}") for i, col in enumerate(fit_df.columns)], "L")
    except Exception:
        pass
    return fit_wks


def apply_fit_plot_style(plot, color: str) -> None:
    try:
        plot.color = color
    except Exception:
        pass
    for prop, value in (
        ("line.width", 2.0),
        ("symbol.size", 0.0),
    ):
        safe_plot_set(plot, "set_float", prop, value)
    for prop, value in (
        ("symbol.kind", 0),
        ("line.style", 1),
    ):
        safe_plot_set(plot, "set_int", prop, value)
    try:
        plot.set_cmd("-l 1", "-w 2", "-k 0", "-z 0")
    except Exception:
        pass


def add_fit_annotations(
    gl,
    fit_specs: list[tuple[LinearFitResult, str, str, str]],
    axis_limits: AxisLimits,
) -> None:
    x_span = axis_limits.x_max - axis_limits.x_min
    y_span = axis_limits.y_max - axis_limits.y_min
    x_label = axis_limits.x_min + x_span * 0.04
    y_label_top = axis_limits.y_max - y_span * 0.05
    y_label_step = y_span * 0.13 if len(fit_specs) > 1 else y_span * 0.10

    for index, (fit, color, x_title, y_title) in enumerate(fit_specs):
        label_y = y_label_top - index * y_label_step
        label = gl.add_label(fit_label_text(fit, x_title, y_title, len(fit_specs) > 1), x_label, label_y)
        if label:
            try:
                label.set_int("attach", 2)
                label.set_float("x1", x_label)
                label.set_float("y1", label_y)
                label.set_int("font", 0)
                label.set_int("fsize", 14)
                label.set_int("color", origin_color_value(color))
            except Exception:
                pass


def fit_label_text(fit: LinearFitResult, x_title: str, y_title: str, include_name: bool) -> str:
    sign = "+" if fit.intercept >= 0 else "-"
    equation = f"y = {fit.slope:.6g}x {sign} {abs(fit.intercept):.6g}"
    r2 = f"R2 = {fit.r_squared:.6g}"
    p_value = f"P = {format_p_value(fit.p_value)}"
    if include_name:
        return f"{y_title} vs {x_title}\n{equation}\n{r2}\n{p_value}"
    return f"{equation}\n{r2}\n{p_value}"


def format_p_value(value: float) -> str:
    if not math.isfinite(value):
        return "NA"
    if value < 1e-4:
        return f"{value:.3e}"
    return f"{value:.6g}"


def origin_color_value(color: str) -> int:
    text = str(color).strip()
    if re.fullmatch(r"#[0-9a-fA-F]{6}", text):
        red = int(text[1:3], 16)
        green = int(text[3:5], 16)
        blue = int(text[5:7], 16)
        return red + green * 256 + blue * 65536
    named = {
        "black": 1,
        "red": 2,
        "green": 3,
        "blue": 4,
        "cyan": 5,
        "magenta": 6,
        "yellow": 7,
        "white": 8,
    }
    return named.get(text.lower(), 1)


def compute_linear_fit(xy: pd.DataFrame, x_title: str, y_title: str, logger: LogFunc) -> Optional[LinearFitResult]:
    numeric = pd.DataFrame({
        "x": pd.to_numeric(xy.iloc[:, 0], errors="coerce"),
        "y": pd.to_numeric(xy.iloc[:, 1], errors="coerce"),
    }).dropna(how="any")
    if len(numeric) < 2:
        logger(f"线性拟合跳过: {y_title} 可用数据点少于 2 个。")
        return None

    x_values = [float(item) for item in numeric["x"].tolist()]
    y_values = [float(item) for item in numeric["y"].tolist()]
    n = len(x_values)
    x_mean = sum(x_values) / n
    y_mean = sum(y_values) / n
    ss_xx = sum((x - x_mean) ** 2 for x in x_values)
    if ss_xx == 0:
        logger(f"线性拟合跳过: {y_title} 的 X 数据没有变化。")
        return None

    ss_xy = sum((x - x_mean) * (y - y_mean) for x, y in zip(x_values, y_values))
    slope = ss_xy / ss_xx
    intercept = y_mean - slope * x_mean
    fitted = [slope * x + intercept for x in x_values]
    ss_res = sum((y - y_fit) ** 2 for y, y_fit in zip(y_values, fitted))
    ss_tot = sum((y - y_mean) ** 2 for y in y_values)
    r_squared = 1.0 if ss_tot == 0 else 1.0 - ss_res / ss_tot
    p_value = compute_slope_p_value(slope, ss_xx, ss_res, n)
    x_min = min(x_values)
    x_max = max(x_values)
    logger(
        f"???? [{y_title} vs {x_title}]: "
        f"n={n}, X??={x_min:.6g}-{x_max:.6g}, "
        f"y = {slope:.6g}x + {intercept:.6g}, R2 = {r_squared:.6g}, P = {format_p_value(p_value)}"
    )
    return LinearFitResult(slope=slope, intercept=intercept, r_squared=r_squared, p_value=p_value, x_min=x_min, x_max=x_max)


def compute_slope_p_value(slope: float, ss_xx: float, ss_res: float, n: int) -> float:
    if n <= 2 or ss_xx <= 0:
        return math.nan
    if ss_res <= 1e-24:
        return 0.0
    residual_variance = ss_res / (n - 2)
    slope_stderr = math.sqrt(residual_variance / ss_xx)
    if slope_stderr <= 0:
        return 0.0
    t_value = abs(slope / slope_stderr)
    degrees = n - 2
    return student_t_two_sided_p_value(t_value, degrees)


def student_t_two_sided_p_value(t_value: float, degrees: int) -> float:
    if degrees <= 0 or not math.isfinite(t_value):
        return math.nan
    x_value = degrees / (degrees + t_value * t_value)
    return max(0.0, min(1.0, regularized_incomplete_beta(x_value, degrees / 2.0, 0.5)))


def regularized_incomplete_beta(x_value: float, a_value: float, b_value: float) -> float:
    if x_value <= 0:
        return 0.0
    if x_value >= 1:
        return 1.0
    log_beta = (
        math.lgamma(a_value + b_value)
        - math.lgamma(a_value)
        - math.lgamma(b_value)
        + a_value * math.log(x_value)
        + b_value * math.log1p(-x_value)
    )
    beta_front = math.exp(log_beta)
    if x_value < (a_value + 1.0) / (a_value + b_value + 2.0):
        return beta_front * incomplete_beta_fraction(a_value, b_value, x_value) / a_value
    return 1.0 - beta_front * incomplete_beta_fraction(b_value, a_value, 1.0 - x_value) / b_value


def incomplete_beta_fraction(a_value: float, b_value: float, x_value: float) -> float:
    max_iterations = 200
    epsilon = 3e-12
    tiny = 1e-300
    qab = a_value + b_value
    qap = a_value + 1.0
    qam = a_value - 1.0
    c_value = 1.0
    d_value = 1.0 - qab * x_value / qap
    if abs(d_value) < tiny:
        d_value = tiny
    d_value = 1.0 / d_value
    h_value = d_value
    for m_value in range(1, max_iterations + 1):
        m2_value = 2 * m_value
        aa_value = m_value * (b_value - m_value) * x_value / ((qam + m2_value) * (a_value + m2_value))
        d_value = 1.0 + aa_value * d_value
        if abs(d_value) < tiny:
            d_value = tiny
        c_value = 1.0 + aa_value / c_value
        if abs(c_value) < tiny:
            c_value = tiny
        d_value = 1.0 / d_value
        h_value *= d_value * c_value
        aa_value = -(a_value + m_value) * (qab + m_value) * x_value / ((a_value + m2_value) * (qap + m2_value))
        d_value = 1.0 + aa_value * d_value
        if abs(d_value) < tiny:
            d_value = tiny
        c_value = 1.0 + aa_value / c_value
        if abs(c_value) < tiny:
            c_value = tiny
        d_value = 1.0 / d_value
        delta = d_value * c_value
        h_value *= delta
        if abs(delta - 1.0) < epsilon:
            break
    return h_value

def safe_plot_set(plot, method_name: str, prop: str, value) -> bool:
    try:
        getattr(plot, method_name)(prop, value)
        return True
    except Exception:
        return False


def format_axes(
    gp,
    gl,
    x_title: str,
    y_title: str,
    axis_limits: AxisLimits,
    x_log: bool = False,
    y_reverse: bool = False,
    y_axis_side: str = "left",
    x_axis_side: str = "bottom",
    overlay_index: int = 0,
) -> None:
    gp.activate()
    gl.activate()
    try:
        gl.set_xlim(axis_limits.x_min, axis_limits.x_max, axis_limits.x_step)
        if y_reverse:
            gl.set_ylim(axis_limits.y_max, axis_limits.y_min)
        else:
            gl.set_ylim(axis_limits.y_min, axis_limits.y_max)
    except Exception:
        pass

    lt = gl.lt_exec
    lt(f"layer.x.type = {2 if x_log else 1}")
    lt(f"layer.x.from = {axis_limits.x_min}")
    lt(f"layer.x.to = {axis_limits.x_max}")
    lt(f"layer.x.inc = {axis_limits.x_step}")
    if y_reverse:
        lt(f"layer.y.from = {axis_limits.y_max}")
        lt(f"layer.y.to = {axis_limits.y_min}")
    else:
        lt(f"layer.y.from = {axis_limits.y_min}")
        lt(f"layer.y.to = {axis_limits.y_max}")
    lt(f"layer.y.inc = {axis_limits.y_step}")

    gl.axis("x").title = plain_title_for_origin(x_title)
    gl.axis("y").title = plain_title_for_origin(y_title)
    x_label_cmd = "xt" if x_axis_side == "top" else "xb"
    y_label_cmd = "yr" if y_axis_side == "right" else "yl"
    lt(f'label -{x_label_cmd} "{origin_rich_text(x_title)}"')
    lt(f'label -{y_label_cmd} "{origin_rich_text(y_title)}"')

    if y_axis_side == "right":
        lt("layer.y.showAxes = 2")
        lt("layer.y2.showlabel = 1")
        lt("layer.y.showlabel = 0")
    elif SHOW_TOP_RIGHT_AXES:
        lt("layer.y.showAxes = 3")
        lt("layer.y2.showlabel = 0")
    else:
        lt("layer.y.showAxes = 1")

    if x_axis_side == "top":
        lt("layer.x.showAxes = 2")
        lt("layer.x2.showlabel = 1")
        lt("layer.x.showlabel = 0")
    elif SHOW_TOP_RIGHT_AXES:
        lt("layer.x.showAxes = 3")
        lt("layer.x2.showlabel = 0")
    else:
        lt("layer.x.showAxes = 1")

    for axis in ("x", "y", "x2", "y2"):
        lt(f"layer.{axis}.thickness = {AXIS_WIDTH}")
        lt(f"layer.{axis}.ticks = 1")
        lt(f"layer.{axis}.tickLength = {TICK_LENGTH}")
        lt(f"layer.{axis}.minorTicks = 0")
        lt(f'layer.{axis}.label.font = font("{FONT_NAME}")')
        lt(f"layer.{axis}.label.fSize = {FONT_SIZE}")
        lt(f"layer.{axis}.label.bold = 1")
    lt("layer.x.labelType = 1")
    lt("layer.x.labelSubtype = 1")
    lt(f"layer.x.decPlaces = {axis_limits.x_decimals}")
    lt("layer.y.labelType = 1")
    lt("layer.y.labelSubtype = 1")
    lt(f"layer.y.decPlaces = {axis_limits.y_decimals}")

    text_objects = [x_label_cmd, y_label_cmd]
    for obj in text_objects:
        lt(f'{obj}.font = font("{FONT_NAME}")')
        lt(f"{obj}.fsize = {FONT_SIZE}")
        lt(f"{obj}.bold = 1")
    if overlay_index:
        offset = min(30, overlay_index * 8)
        if y_axis_side == "right":
            lt(f"yr.x += {offset}")
        else:
            lt(f"yl.x -= {offset}")
        if x_axis_side == "top":
            lt(f"xt.y += {min(20, overlay_index * 5)}")

    if not SHOW_LEGEND:
        try:
            gl.label("Legend").remove()
        except Exception:
            pass
    lt("doc -uw")


def apply_legend(gl, entries: list[str]) -> None:
    if not SHOW_LEGEND or not entries:
        return
    try:
        gl.activate()
        legend_lines = []
        for index, entry in enumerate(entries, start=1):
            legend_lines.append(f"\\l({index}) {origin_mixed_font_text(entry)}")
        legend_text = "\\r\\n".join(legend_lines)
        escaped = legend_text.replace('"', "'")
        lt = gl.lt_exec
        lt("legend -r")
        lt(f'Legend.text$ = "{escaped}"')
        lt("Legend.fsize = 16")
        lt("Legend.left = 85")
        lt("Legend.top = 12")
        lt("Legend.border = 0")
        lt("doc -uw")
    except Exception:
        try:
            gl.label("Legend").remove()
        except Exception:
            pass


def export_outputs(
    gp,
    figures_dir: Path,
    origin_dir: Path,
    output_stem: str,
    logger: LogFunc,
    save_origin: bool = True,
) -> tuple[Path, Path]:
    output_stem = sanitize_filename(output_stem)
    emf_path = figures_dir / f"{output_stem}.emf"
    opj_path = origin_dir / f"{output_stem}.opju"

    figures_dir.mkdir(parents=True, exist_ok=True)
    origin_dir.mkdir(parents=True, exist_ok=True)
    lock_graph_element_scale(gp)
    fit_page_to_layers(gp, logger)
    lock_graph_element_scale(gp)
    logger(f"?? EMF: {emf_path}")
    gp.lt_exec(
        f'expgraph type:=emf filename:="{output_stem}" '
        f'path:="{to_origin_path(figures_dir)}" overwrite:=replace'
    )
    fit_page_to_layers(gp, logger, quiet=True)
    lock_graph_element_scale(gp)
    if save_origin:
        logger(f"?? Origin ??: {opj_path}")
        gp.lt_exec(f'save -i "{to_origin_path(opj_path)}"')
    return emf_path, opj_path


def fit_page_to_layers(gp, logger: LogFunc, quiet: bool = False) -> None:
    try:
        gp.activate()
        gp.lt_exec("page -fit -u -m 0 -d 0")
        gp.lt_exec("doc -uw")
        if not quiet:
            logger("已自动执行 Origin Fit Page to Layers，仅调整页面边界，不缩放点和文字。")
    except Exception:
        try:
            gp.activate()
            gp.lt_exec("pfit2l margin:=tight direction:=0 go:=1")
            gp.lt_exec("doc -uw")
            if not quiet:
                logger("已使用 Origin X-Function 调整导出页面。")
        except Exception as exc:
            if not quiet:
                logger(f"自动适配页面失败，可在 Origin 中手动 Fit Page to Layers：{exc}")


def save_single_origin_project(op, origin_dir: Path, logger: LogFunc) -> Path:
    origin_dir.mkdir(parents=True, exist_ok=True)
    opj_path = origin_dir / "All_Plots_In_One_Origin.opju"
    logger(f"???? Origin ??: {opj_path}")
    try:
        op.lt_exec(f'save -i "{to_origin_path(opj_path)}"')
    except Exception:
        try:
            page = op.find_graph()
            page.lt_exec(f'save -i "{to_origin_path(opj_path)}"')
        except Exception as exc:
            raise RuntimeError(f"?? Origin ??????: {exc}") from exc
    return opj_path

def select_marker_color(config: PlotConfig, index: int) -> str:
    if config.marker_colors and index < len(config.marker_colors):
        return config.marker_colors[index]
    if config.cycle_colors:
        return COLORS_FOR_MULTIPLE_PLOTS[index % len(COLORS_FOR_MULTIPLE_PLOTS)]
    return config.marker_color


def select_marker_size(config: PlotConfig, index: int) -> float:
    if config.marker_sizes and index < len(config.marker_sizes):
        return float(config.marker_sizes[index])
    return float(config.marker_size)


def select_marker_symbol(config: PlotConfig, index: int) -> MarkerSymbol:
    if config.marker_symbols and index < len(config.marker_symbols):
        return config.marker_symbols[index]
    return config.marker_symbol


def select_hollow_marker_fill_color(config: PlotConfig, index: int) -> str:
    if config.hollow_marker_fill_colors and index < len(config.hollow_marker_fill_colors):
        return config.hollow_marker_fill_colors[index]
    return config.hollow_marker_fill_color


def select_hollow_marker_edge_color(config: PlotConfig, index: int, fallback: str) -> str:
    if config.hollow_marker_edge_colors and index < len(config.hollow_marker_edge_colors):
        return config.hollow_marker_edge_colors[index]
    return config.hollow_marker_edge_color or fallback


def select_hollow_marker_edge_width(config: PlotConfig, index: int) -> float:
    if config.hollow_marker_edge_widths and index < len(config.hollow_marker_edge_widths):
        return float(config.hollow_marker_edge_widths[index])
    return float(config.hollow_marker_edge_width)


def select_plot_type(config: PlotConfig, index: int) -> str:
    if config.plot_types and index < len(config.plot_types):
        return normalize_plot_type(config.plot_types[index])
    return normalize_plot_type(config.plot_type)


def select_line_color(config: PlotConfig, index: int, fallback: str) -> str:
    if config.line_colors and index < len(config.line_colors):
        return config.line_colors[index]
    return config.line_color or fallback


def select_line_width(config: PlotConfig, index: int) -> float:
    if config.line_widths and index < len(config.line_widths):
        return float(config.line_widths[index])
    return float(config.line_width)


def select_line_style(config: PlotConfig, index: int) -> str:
    if config.line_styles and index < len(config.line_styles):
        return normalize_line_style(config.line_styles[index])
    return normalize_line_style(config.line_style)


def select_linear_fit(config: PlotConfig, index: int) -> bool:
    if config.linear_fits and index < len(config.linear_fits):
        return bool(config.linear_fits[index])
    return bool(config.linear_fit)


def select_x_log(config: PlotConfig, index: int) -> bool:
    if config.x_logs and index < len(config.x_logs):
        return bool(config.x_logs[index])
    return bool(config.x_log)


def select_y_reverse(config: PlotConfig, index: int) -> bool:
    if config.y_reverses and index < len(config.y_reverses):
        return bool(config.y_reverses[index])
    return bool(config.y_reverse)


def select_axis_width_percent(config: PlotConfig, index: int) -> float:
    if config.axis_width_percents and index < len(config.axis_width_percents):
        return float(config.axis_width_percents[index])
    return float(config.axis_width_percent)


def select_axis_height_percent(config: PlotConfig, index: int) -> float:
    if config.axis_height_percents and index < len(config.axis_height_percents):
        return float(config.axis_height_percents[index])
    return float(config.axis_height_percent)


def select_combined_x_log(config: PlotConfig, pairs: list[PlotPair], logger: LogFunc) -> bool:
    values = [select_x_log(config, pair.index) for pair in pairs]
    if len(set(values)) > 1:
        logger("提示: 合并图只有一套X轴，已按合并列表第一组数据的X轴对数设置执行。")
    return values[0] if values else bool(config.x_log)


def select_combined_y_reverse(config: PlotConfig, pairs: list[PlotPair], logger: LogFunc) -> bool:
    values = [select_y_reverse(config, pair.index) for pair in pairs]
    if len(set(values)) > 1:
        logger("提示: 合并图只有一套Y轴，已按合并列表第一组数据的Y轴倒序设置执行。")
    return values[0] if values else bool(config.y_reverse)


def normalize_plot_type(value: str) -> str:
    text = str(value).strip().lower().replace("-", "_").replace(" ", "_")
    aliases = {
        "scatter": "scatter",
        "s": "scatter",
        "point": "scatter",
        "line_symbol": "line_symbol",
        "line_symbols": "line_symbol",
        "point_line": "line_symbol",
        "y": "line_symbol",
        "line": "line",
        "l": "line",
    }
    if text not in aliases:
        raise ValueError("plot_type 必须是 scatter、line_symbol 或 line")
    return aliases[text]


def origin_plot_type(value: str) -> str:
    return PLOT_TYPE_MAP[normalize_plot_type(value)]


def normalize_line_style(value: str) -> str:
    text = str(value).strip().lower().replace("-", "_").replace(" ", "_")
    if text.isdigit():
        number = int(text)
        for name, style_id in LINE_STYLE_MAP.items():
            if style_id == number:
                return name
        return "solid"
    aliases = {
        "solid": "solid",
        "dash": "dash",
        "dashed": "dash",
        "dot": "dot",
        "dotted": "dot",
        "dash_dot": "dash_dot",
        "dashdot": "dash_dot",
        "dash_dot_dot": "dash_dot_dot",
    }
    if text not in aliases:
        raise ValueError("line_style 必须是 solid、dash、dot、dash_dot 或 dash_dot_dot")
    return aliases[text]


def normalize_marker_symbol(symbol: MarkerSymbol) -> int:
    if isinstance(symbol, int):
        return symbol
    text = normalize_marker_symbol_name(symbol)
    if text.isdigit():
        return int(text)
    if text not in SYMBOL_MAP:
        raise ValueError(f"未知散点类型: {symbol}. 可用: {', '.join(sorted(SYMBOL_MAP))}，也可输入 Origin symbol 编号。")
    return SYMBOL_MAP[text]


def normalize_marker_symbol_name(symbol: MarkerSymbol) -> str:
    if isinstance(symbol, int):
        return str(symbol)
    return str(symbol).strip().lower().replace("-", "_").replace(" ", "_")


def clean_column_title(value: object, fallback: str) -> str:
    text = str(value).strip()
    if not text or text.lower().startswith("unnamed"):
        return fallback
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\.\d+$", "", text)
    return text


def sanitize_filename(name: str, fallback: str = "scatter_plot") -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\r\n\t]+', "_", str(name)).strip(" ._")
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or fallback


def plain_title_for_origin(text: str) -> str:
    return str(text).strip()


def origin_rich_text(text: str) -> str:
    text = str(text).strip().replace('"', "'")
    text = re.sub(r"cm\s*\^\s*\{\s*-1\s*\}", r"cm\\+(-1)", text, flags=re.I)
    text = re.sub(r"cm\s*\^\s*-1", r"cm\\+(-1)", text, flags=re.I)
    text = re.sub(r"cm\s*-\s*1", r"cm\\+(-1)", text, flags=re.I)
    return rf"\f:{FONT_NAME}(\b({text}))"


def origin_mixed_font_text(text: str) -> str:
    text = str(text).strip().replace('"', "'")
    text = re.sub(r"cm\s*\^\s*\{\s*-1\s*\}", r"cm\\+(-1)", text, flags=re.I)
    text = re.sub(r"cm\s*\^\s*-1", r"cm\\+(-1)", text, flags=re.I)
    text = re.sub(r"cm\s*-\s*1", r"cm\\+(-1)", text, flags=re.I)
    if not text:
        return ""

    parts: list[str] = []
    current = ""
    current_is_chinese: Optional[bool] = None
    for char in text:
        is_chinese = bool(re.match(r"[\u3400-\u9fff]", char))
        if current and is_chinese != current_is_chinese:
            parts.append(font_wrapped_text(current, current_is_chinese))
            current = char
        else:
            current += char
        current_is_chinese = is_chinese
    if current:
        parts.append(font_wrapped_text(current, current_is_chinese))
    return "".join(parts)


def font_wrapped_text(text: str, is_chinese: Optional[bool]) -> str:
    font_name = "SimSun" if is_chinese else FONT_NAME
    return rf"\f:{font_name}({text})"


def nice_number(value: float, round_value: bool = True) -> float:
    if not math.isfinite(value) or value <= 0:
        return 1.0
    exponent = math.floor(math.log10(value))
    fraction = value / (10**exponent)
    if round_value:
        if fraction < 1.5:
            nice_fraction = 1
        elif fraction < 3:
            nice_fraction = 2
        elif fraction < 7:
            nice_fraction = 5
        else:
            nice_fraction = 10
    else:
        if fraction <= 1:
            nice_fraction = 1
        elif fraction <= 2:
            nice_fraction = 2
        elif fraction <= 5:
            nice_fraction = 5
        else:
            nice_fraction = 10
    return nice_fraction * (10**exponent)


def nice_floor(value: float, step: float) -> float:
    return math.floor(value / step) * step if math.isfinite(value) and step > 0 else value


def nice_ceil(value: float, step: float) -> float:
    return math.ceil(value / step) * step if math.isfinite(value) and step > 0 else value


def infer_decimals(step: float) -> int:
    if not math.isfinite(step) or step <= 0 or abs(step - round(step)) < 1e-9:
        return 0
    return min(4, max(0, int(math.ceil(-math.log10(step))) + 1))


def to_origin_path(path: Path) -> str:
    return str(path).replace("\\", "/")


def make_dataset_label(data_path: Path, sheet_name: Optional[str]) -> str:
    if sheet_name is None:
        return sanitize_filename(data_path.stem, fallback="scatter")
    return sanitize_filename(f"{data_path.stem}__{sheet_name}", fallback="scatter")


def print_summary(success: Iterable[str], failed: Iterable[tuple[str, str]], logger: LogFunc) -> None:
    success = list(success)
    failed = list(failed)
    logger("")
    logger("=" * 70)
    logger("绘图任务完成")
    logger(f"成功: {len(success)} 个")
    logger(f"失败: {len(failed)} 个")
    if success:
        logger("成功列表:")
        for item in success:
            logger(f"  - {item}")
    if failed:
        logger("失败列表:")
        for item, err in failed:
            logger(f"  - {item}: {err}")


def parse_args() -> PlotConfig:
    parser = argparse.ArgumentParser(description="调用 Origin 2023b 批量绘制二维散点图。")
    parser.add_argument("--mode", choices=("auto", "single", "batch", "excel", "sheets"), default="single")
    parser.add_argument("--input-file", default="")
    parser.add_argument("--batch-dir", default="")
    parser.add_argument("--layout", choices=("auto", "pairs", "common-x"), default="auto")
    parser.add_argument("--plot-mode", choices=("combined", "separate"), default="combined")
    parser.add_argument("--marker-size", type=float, default=8.0)
    parser.add_argument(
        "--marker-sizes",
        default="",
        help="按顺序使用的散点大小列表，用逗号/分号/空格分隔，例如 8,10,12。",
    )
    parser.add_argument("--marker-symbol", default="circle")
    parser.add_argument(
        "--marker-symbols",
        default="",
        help="按顺序使用的散点类型列表，用逗号/分号/空格分隔，例如 circle,square,triangle_up。",
    )
    parser.add_argument("--marker-color", default="#1F77B4")
    parser.add_argument(
        "--marker-colors",
        default="",
        help="按顺序使用的颜色列表，用逗号/分号/空格分隔，例如 #1F77B4,#D73027,black。",
    )
    parser.add_argument("--hollow-marker-fill-color", default="#FFFFFF", help="空心圆内部颜色。")
    parser.add_argument("--hollow-marker-fill-colors", default="", help="按顺序使用的空心圆内部颜色列表。")
    parser.add_argument("--hollow-marker-edge-color", default="#1F77B4", help="空心圆边缘颜色。")
    parser.add_argument("--hollow-marker-edge-colors", default="", help="按顺序使用的空心圆边缘颜色列表。")
    parser.add_argument("--hollow-marker-edge-width", type=float, default=1.5, help="空心圆圆圈线宽。")
    parser.add_argument("--hollow-marker-edge-widths", default="", help="按顺序使用的空心圆圆圈线宽列表。")
    parser.add_argument("--plot-type", default="scatter", choices=("scatter", "line_symbol", "line"))
    parser.add_argument("--plot-types", default="", help="按顺序使用的图形类型列表: scatter,line_symbol,line。")
    parser.add_argument("--line-color", default="#1F77B4")
    parser.add_argument("--line-colors", default="", help="按顺序使用的线条颜色列表。")
    parser.add_argument("--line-width", type=float, default=2.0)
    parser.add_argument("--line-widths", default="", help="按顺序使用的线条粗细列表。")
    parser.add_argument("--line-style", default="solid", choices=("solid", "dash", "dot", "dash_dot", "dash_dot_dot"))
    parser.add_argument("--line-styles", default="", help="按顺序使用的线条类型列表。")
    parser.add_argument("--combine-indices", default="", help="需要合并到同一张图的数据组序号，从 1 开始，例如 1,3,4。")
    parser.add_argument("--x-log", action="store_true", help="X 轴使用对数坐标。")
    parser.add_argument("--x-logs", default="", help="按顺序使用的每组 X 轴对数设置，例如 true,false,true。")
    parser.add_argument("--y-reverse", action="store_true", help="Y 轴倒序显示。")
    parser.add_argument("--y-reverses", default="", help="按顺序使用的每组 Y 轴倒序设置，例如 false,true,false。")
    parser.add_argument("--cycle-colors", action="store_true")
    parser.add_argument("--no-cycle-colors", dest="cycle_colors", action="store_false")
    parser.set_defaults(cycle_colors=False)
    parser.add_argument("--linear-fit", action="store_true", help="为每个散点图独立添加线性拟合线。")
    parser.add_argument("--no-linear-fit", dest="linear_fit", action="store_false")
    parser.set_defaults(linear_fit=False)
    parser.add_argument(
        "--linear-fits",
        default="",
        help="按顺序设置每组是否拟合，用逗号/分号/空格分隔，例如 1,0,1 或 true,false,true。",
    )
    parser.add_argument("--single-origin-project", action="store_true", help="将本次任务的所有图页保存到同一个 OPJU 项目。")
    parser.add_argument("--separate-origin-projects", dest="single_origin_project", action="store_false")
    parser.set_defaults(single_origin_project=False)
    parser.add_argument("--axis-width-percent", type=float, default=70.0, help="X轴绘图区长度，占页面宽度百分比，建议 20-90。")
    parser.add_argument("--axis-height-percent", type=float, default=68.0, help="Y轴绘图区长度，占页面高度百分比，建议 20-85。")
    parser.add_argument("--axis-width-percents", default="", help="按顺序使用的每组X轴长度百分比，例如 70,80,65。")
    parser.add_argument("--axis-height-percents", default="", help="按顺序使用的每组Y轴长度百分比，例如 68,60,75。")
    parser.add_argument("--figures-dir", default="")
    parser.add_argument("--origin-dir", default="")
    args = parser.parse_args()
    return PlotConfig(
        mode=args.mode,
        input_file=args.input_file,
        batch_dir=args.batch_dir,
        layout=args.layout,
        plot_mode=args.plot_mode,
        marker_size=args.marker_size,
        marker_sizes=parse_marker_sizes(args.marker_sizes),
        marker_symbol=args.marker_symbol,
        marker_color=args.marker_color,
        marker_colors=parse_marker_colors(args.marker_colors),
        marker_symbols=parse_marker_symbols(args.marker_symbols),
        hollow_marker_fill_color=args.hollow_marker_fill_color,
        hollow_marker_fill_colors=parse_marker_colors(args.hollow_marker_fill_colors),
        hollow_marker_edge_color=args.hollow_marker_edge_color,
        hollow_marker_edge_colors=parse_marker_colors(args.hollow_marker_edge_colors),
        hollow_marker_edge_width=args.hollow_marker_edge_width,
        hollow_marker_edge_widths=parse_marker_sizes(args.hollow_marker_edge_widths),
        cycle_colors=args.cycle_colors,
        plot_type=args.plot_type,
        plot_types=parse_marker_symbols(args.plot_types),
        line_color=args.line_color,
        line_colors=parse_marker_colors(args.line_colors),
        line_width=args.line_width,
        line_widths=parse_marker_sizes(args.line_widths),
        line_style=args.line_style,
        line_styles=parse_marker_symbols(args.line_styles),
        combine_indices=[index - 1 for index in parse_int_list(args.combine_indices)],
        x_log=args.x_log,
        x_logs=parse_bool_list(args.x_logs),
        y_reverse=args.y_reverse,
        y_reverses=parse_bool_list(args.y_reverses),
        linear_fit=args.linear_fit,
        linear_fits=parse_linear_fits(args.linear_fits),
        single_origin_project=args.single_origin_project,
        axis_width_percent=args.axis_width_percent,
        axis_height_percent=args.axis_height_percent,
        axis_width_percents=parse_marker_sizes(args.axis_width_percents),
        axis_height_percents=parse_marker_sizes(args.axis_height_percents),
        figures_dir=args.figures_dir,
        origin_dir=args.origin_dir,
    )


def parse_marker_sizes(value: str) -> list[float]:
    if not value:
        return []
    sizes: list[float] = []
    for item in re.split(r"[,;，；\s]+", value.strip()):
        if item:
            sizes.append(float(item))
    return sizes


def parse_marker_colors(value: str) -> list[str]:
    if not value:
        return []
    return [item for item in re.split(r"[,;，；\s]+", value.strip()) if item]


def parse_marker_symbols(value: str) -> list[str]:
    if not value:
        return []
    return [item for item in re.split(r"[,;，；\s]+", value.strip()) if item]


def parse_int_list(value: str) -> list[int]:
    if not value:
        return []
    values: list[int] = []
    for item in re.split(r"[,;，；\s]+", value.strip()):
        if item:
            values.append(int(item))
    return values


def parse_linear_fits(value: str) -> list[bool]:
    if not value:
        return []
    result: list[bool] = []
    for item in re.split(r"[,;，；\s]+", value.strip()):
        text = item.strip().lower()
        if not text:
            continue
        if text in {"1", "true", "yes", "y", "on", "fit", "拟合", "是"}:
            result.append(True)
        elif text in {"0", "false", "no", "n", "off", "none", "不拟合", "否"}:
            result.append(False)
        else:
            raise ValueError(f"无法识别线性拟合开关: {item}")
    return result


def parse_bool_list(value: str) -> list[bool]:
    return parse_linear_fits(value)


def main() -> None:
    try:
        run_plotting(parse_args())
    except Exception as exc:
        print(f"运行失败: {exc}")
        sys.exit(1)


if __name__ == "__main__":
    main()
