"""
Origin 2023b 二维散点图批量绘制脚本
===================================
适用：普通 Origin 2023b / OriginPro 2023b 均可。

功能：
    1. 调用 Origin 2023b 绘制二维散点图；
    2. 单文件绘图：读取 data/ 中的 1.xlsx、1.txt、1.csv 或任意数据文件；
    3. 批量绘图：读取 data/batch/ 下所有 xlsx/txt/csv 文件；
    4. 多 Sheet 绘图：读取一个 Excel 文件中的所有 sheet；
    5. 支持一个文件中包含多组二维散点数据；
    6. 散点大小、类型、颜色均可在脚本顶部或命令行中调整；
    7. X 轴和 Y 轴标题自动使用导入数据第一行表头内容，单位会随表头一起显示。

推荐数据格式一：两列一组重复排列
    第一行：X1标题/单位, Y1标题/单位, X2标题/单位, Y2标题/单位 ...
    后续行：X1数据,     Y1数据,     X2数据,     Y2数据 ...

推荐数据格式二：公共 X 列
    第一行：X标题/单位, Y1标题/单位, Y2标题/单位, Y3标题/单位 ...
    后续行：X数据,     Y1数据,     Y2数据,     Y3数据 ...

运行方式（Windows PowerShell）：
    cd "e:\\claude code\\painting"
    .\\.venv\\Scripts\\activate

    # 自动模式：data/batch/ 有数据文件则批量绘图，否则单文件绘图
    python scripts/origin_2d_scatter_batch.py

    # 单文件绘图
    python scripts/origin_2d_scatter_batch.py --mode single

    # 多数据文件批量绘图
    python scripts/origin_2d_scatter_batch.py --mode batch

    # 单 Excel 多 Sheet 批量绘图
    python scripts/origin_2d_scatter_batch.py --mode sheets --input-file data/1.xlsx

    # 调整散点大小、类型、颜色
    python scripts/origin_2d_scatter_batch.py --marker-size 9 --marker-symbol circle --marker-color "#D73027"

依赖：
    pip install originpro pandas openpyxl
"""

from __future__ import annotations

import argparse
import math
import os
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Union

try:
    import originpro as op
except ImportError:
    print("错误：未找到 originpro 包。")
    print("请运行: pip install originpro")
    print("如果已安装但仍报错，请检查 VSCode/Claude Code 使用的 Python 解释器是否正确。")
    sys.exit(1)

try:
    import pandas as pd
except ImportError:
    print("错误：未找到 pandas/openpyxl 包。")
    print("请运行: pip install pandas openpyxl")
    sys.exit(1)


# ============================================================
# 绘图参数：散点大小、类型、颜色在这里调整
# ============================================================

# 散点大小。Origin 中通常 6–12 比较适合论文图；数值越大点越大。
MARKER_SIZE = 8.0

# 散点类型。可填英文名，也可直接填 Origin 的 symbol 编号。
# 常用英文名：circle, square, triangle_up, triangle_down, diamond, plus, cross, star
MARKER_SYMBOL: Union[str, int] = "circle"

# 散点颜色。支持十六进制颜色，也支持 Origin/英文常见颜色名。
# 示例："#1F77B4", "#D73027", "red", "blue", "black"
MARKER_COLOR = "#1F77B4"

# 如果一个数据文件中有多组曲线，可选择是否自动循环颜色。
# False：所有散点图都使用 MARKER_COLOR。
# True：同一文件内不同散点图依次使用 COLORS_FOR_MULTIPLE_PLOTS。
AUTO_CYCLE_COLORS = False
COLORS_FOR_MULTIPLE_PLOTS = [
    "#1F77B4", "#D73027", "#4DAF4A", "#984EA3", "#FF7F00",
    "#A65628", "#F781BF", "#636363", "#66C2A5", "#FC8D62",
]

# 字体与坐标轴样式
FONT_NAME = "Times New Roman"
FONT_SIZE = 20
AXIS_WIDTH = 2
TICK_LENGTH = 5
SHOW_TOP_RIGHT_AXES = True
SHOW_LEGEND = False

# 自动坐标范围与留白
AUTO_X_RANGE = True
AUTO_Y_RANGE = True
AXIS_PADDING_RATIO = 0.06
AUTO_TICK_TARGET_INTERVALS = 5

# 手动坐标范围；当 AUTO_X_RANGE/AUTO_Y_RANGE=False 时生效。
MANUAL_X_MIN, MANUAL_X_MAX, MANUAL_X_STEP = 0.0, 1.0, 0.2
MANUAL_Y_MIN, MANUAL_Y_MAX, MANUAL_Y_STEP = 0.0, 1.0, 0.2

# 支持的数据文件格式
SUPPORTED_DATA_EXTENSIONS = (".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".dat")
PREFERRED_SINGLE_DATA_NAMES = (
    "1.xlsx", "1.xlsm", "1.xls", "1.csv", "1.txt", "1.dat",
    "scatter_data.xlsx", "scatter_data.csv", "scatter_data.txt",
    "data.xlsx", "data.csv", "data.txt",
)

# Origin 常见 symbol 类型映射。不同 Origin 模板中显示可能略有差异；
# 如需完全匹配某个符号，可直接把 MARKER_SYMBOL 改成 Origin 中对应的数字编号。
SYMBOL_MAP = {
    "square": 1,
    "circle": 2,
    "triangle_up": 3,
    "triangle": 3,
    "triangle_down": 4,
    "diamond": 5,
    "plus": 6,
    "cross": 7,
    "star": 8,
    "none": 0,
}


@dataclass
class ProjectPaths:
    project_dir: Path
    data_dir: Path
    batch_data_dir: Path
    figures_dir: Path
    origin_dir: Path
    batch_figures_dir: Path
    batch_origin_dir: Path


@dataclass
class PlotPair:
    x_col: int
    y_col: int
    x_title: str
    y_title: str
    output_suffix: str


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


def lt(gl, command: str) -> None:
    """执行 LabTalk 命令。"""
    gl.lt_exec(command)


def sanitize_filename(name: str, fallback: str = "scatter_plot") -> str:
    """清理 Windows 文件名非法字符，避免 Origin 导出 EMF/OPJU 时失败。"""
    cleaned = re.sub(r'[<>:"/\\|?*\r\n\t]+', "_", str(name)).strip(" ._")
    cleaned = re.sub(r"\s+", "_", cleaned)
    return cleaned or fallback


def clean_column_title(value: object, fallback: str) -> str:
    """
    将第一行表头转换为轴标题。
    pandas 会把重复表头自动加成 .1、.2，这里会去掉这些后缀。
    """
    text = str(value).strip()
    if not text or text.lower().startswith("unnamed"):
        return fallback
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"\.\d+$", "", text)
    return text


def plain_title_for_origin(text: str) -> str:
    """用于 originpro API 的普通标题文本。"""
    text = str(text).strip()
    # 常见单位写法转换成 Unicode 上标，API 标题通常能直接显示。
    text = re.sub(r"cm\s*\^\s*\{?\s*-?1\s*\}?", "cm⁻¹", text, flags=re.I)
    text = re.sub(r"cm\s*-\s*1", "cm⁻¹", text, flags=re.I)
    text = re.sub(r"Å\s*\^\s*\{?\s*-?1\s*\}?", "Å⁻¹", text)
    text = re.sub(r"A\s*\^\s*\{?\s*-?1\s*\}?", "Å⁻¹", text)
    return text


def origin_rich_text(text: str, bold: bool = True) -> str:
    """
    转换为 Origin 富文本。
    重点处理 cm-1 / cm^-1 / cm^{-1}，使 -1 以上标显示。
    """
    text = str(text).strip()
    text = text.replace('"', "'")

    # cm-1 / cm^-1 / cm^{-1} -> cm\+(-1)
    text = re.sub(r"cm\s*\^\s*\{\s*-1\s*\}", r"cm\\+(-1)", text, flags=re.I)
    text = re.sub(r"cm\s*\^\s*-1", r"cm\\+(-1)", text, flags=re.I)
    text = re.sub(r"cm\s*-\s*1", r"cm\\+(-1)", text, flags=re.I)
    text = text.replace("cm⁻¹", r"cm\+(-1)")

    # Å-1 / A^-1 -> Å\+(-1)
    text = re.sub(r"Å\s*\^\s*\{\s*-1\s*\}", r"Å\\+(-1)", text)
    text = re.sub(r"Å\s*\^\s*-1", r"Å\\+(-1)", text)
    text = re.sub(r"Å\s*-\s*1", r"Å\\+(-1)", text)
    text = text.replace("Å⁻¹", r"Å\+(-1)")

    if bold:
        return rf"\f:{FONT_NAME}(\b({text}))"
    return rf"\f:{FONT_NAME}({text})"


def nice_number(value: float, round_value: bool = True) -> float:
    """把数值转换成 1/2/5/10 × 10^n 形式。"""
    if not math.isfinite(value) or value <= 0:
        return 1.0
    exponent = math.floor(math.log10(value))
    fraction = value / (10 ** exponent)
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
    return nice_fraction * (10 ** exponent)


def nice_floor(value: float, step: float) -> float:
    if not math.isfinite(value) or step <= 0:
        return value
    return math.floor(value / step) * step


def nice_ceil(value: float, step: float) -> float:
    if not math.isfinite(value) or step <= 0:
        return value
    return math.ceil(value / step) * step


def infer_decimals(step: float) -> int:
    if not math.isfinite(step) or step <= 0:
        return 0
    if abs(step - round(step)) < 1e-9:
        return 0
    return min(4, max(0, int(math.ceil(-math.log10(step))) + 1))


def get_project_paths() -> ProjectPaths:
    env_project_dir = os.environ.get("ORIGIN_SCATTER_PROJECT_DIR")
    if env_project_dir:
        project_dir = Path(env_project_dir).resolve()
    else:
        script_dir = Path(__file__).resolve().parent
        project_dir = script_dir.parent

    data_dir = project_dir / "data"
    batch_data_dir = data_dir / "batch"
    figures_dir = project_dir / "figures"
    origin_dir = project_dir / "origin"
    batch_figures_dir = figures_dir / "batch"
    batch_origin_dir = origin_dir / "batch"

    for path in (data_dir, batch_data_dir, figures_dir, origin_dir, batch_figures_dir, batch_origin_dir):
        path.mkdir(parents=True, exist_ok=True)

    return ProjectPaths(
        project_dir=project_dir,
        data_dir=data_dir,
        batch_data_dir=batch_data_dir,
        figures_dir=figures_dir,
        origin_dir=origin_dir,
        batch_figures_dir=batch_figures_dir,
        batch_origin_dir=batch_origin_dir,
    )


def find_data_files(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    files: list[Path] = []
    for ext in SUPPORTED_DATA_EXTENSIONS:
        files.extend(folder.glob(f"*{ext}"))
    return sorted(p for p in files if p.is_file() and not p.name.startswith("~$"))


def get_single_data_path(paths: ProjectPaths) -> Path:
    paths.data_dir.mkdir(parents=True, exist_ok=True)

    for filename in PREFERRED_SINGLE_DATA_NAMES:
        candidate = paths.data_dir / filename
        if candidate.exists():
            print(f"自动选择数据文件: {candidate.name}")
            return candidate

    candidates = [p for p in find_data_files(paths.data_dir) if p.parent == paths.data_dir]
    if candidates:
        print(f"自动选择数据文件: {candidates[0].name}")
        return candidates[0]

    return paths.data_dir / "1.xlsx"


def read_scatter_data(data_path: Path, sheet_name: Optional[Union[str, int]] = 0) -> pd.DataFrame:
    """
    读取二维散点数据。
    第一行自动作为表头，因此坐标轴标题会来自第一行内容。
    """
    if not data_path.exists():
        raise FileNotFoundError(
            f"未找到数据文件：{data_path}\n"
            f"单文件绘图请把数据放到 data/，文件名可为 1.xlsx、1.txt、1.csv 或任意支持格式；"
            f"批量绘图请放到 data/batch/。"
        )

    suffix = data_path.suffix.lower()
    effective_sheet = 0 if sheet_name is None else sheet_name

    if suffix in (".xlsx", ".xlsm", ".xls"):
        label = f"{data_path.name}" if effective_sheet == 0 else f"{data_path.name} / sheet={effective_sheet}"
        print(f"读取 Excel 数据: {label}")
        df = pd.read_excel(data_path, sheet_name=effective_sheet, header=0)
        if isinstance(df, dict):
            if not df:
                raise ValueError(f"Excel 文件中没有可读取的 sheet：{data_path}")
            first_sheet_name = next(iter(df))
            print(f"  提示：读取结果为多个 sheet，自动使用第一个 sheet：{first_sheet_name}")
            df = df[first_sheet_name]

    elif suffix in (".csv", ".txt", ".dat"):
        print(f"读取文本数据: {data_path.name}")
        try:
            df = pd.read_csv(data_path, sep=None, engine="python", header=0)
        except Exception:
            df = pd.read_csv(data_path, sep=r"\s+", engine="python", header=0)
        if df.shape[1] < 2:
            df = pd.read_csv(data_path, sep=r"\s+", engine="python", header=0)
    else:
        raise ValueError(
            f"不支持的数据文件格式：{data_path.suffix}\n"
            f"当前支持：{', '.join(SUPPORTED_DATA_EXTENSIONS)}"
        )

    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    if df.shape[1] < 2:
        raise ValueError("数据至少需要 2 列：X/Y。")

    # 保留第一行表头作为标题，同时将数据区转为数值。
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(axis=0, how="all")
    if df.empty:
        raise ValueError("清理空值后没有可绘制的数据。")

    print(f"  数据规模: {df.shape[0]} 行 × {df.shape[1]} 列")
    return df


def infer_plot_pairs(df: pd.DataFrame, layout: str = "auto") -> list[PlotPair]:
    """
    识别二维散点图组合。

    layout=pairs：X1/Y1、X2/Y2、X3/Y3……
    layout=common-x：第一列为公共 X，其余列分别作为 Y。
    layout=auto：偶数列默认 pairs；奇数列默认 common-x。
    """
    n_cols = df.shape[1]
    if n_cols < 2:
        raise ValueError("数据至少需要 2 列。")

    if layout == "auto":
        layout = "pairs" if n_cols % 2 == 0 else "common-x"

    pairs: list[PlotPair] = []
    columns = list(df.columns)

    if layout == "pairs":
        if n_cols % 2 != 0:
            raise ValueError(
                f"当前为 pairs 布局，但数据共有 {n_cols} 列，不是偶数列。"
                "请改成 X1/Y1、X2/Y2……，或使用 --layout common-x。"
            )
        for i in range(0, n_cols, 2):
            pair_idx = len(pairs) + 1
            x_title = clean_column_title(columns[i], f"X{pair_idx}")
            y_title = clean_column_title(columns[i + 1], f"Y{pair_idx}")
            suffix = sanitize_filename(f"{pair_idx}_{y_title}", fallback=f"plot_{pair_idx}")
            pairs.append(PlotPair(i, i + 1, x_title, y_title, suffix))

    elif layout == "common-x":
        x_title = clean_column_title(columns[0], "X")
        for i in range(1, n_cols):
            pair_idx = len(pairs) + 1
            y_title = clean_column_title(columns[i], f"Y{pair_idx}")
            suffix = sanitize_filename(f"{pair_idx}_{y_title}", fallback=f"plot_{pair_idx}")
            pairs.append(PlotPair(0, i, x_title, y_title, suffix))
    else:
        raise ValueError("--layout 只能为 auto、pairs 或 common-x。")

    print(f"  识别到二维散点图数量: {len(pairs)}")
    for idx, pair in enumerate(pairs, start=1):
        valid_points = df.iloc[:, [pair.x_col, pair.y_col]].dropna(how="any").shape[0]
        print(f"    图 {idx}: X='{pair.x_title}', Y='{pair.y_title}', 有效点数={valid_points}")

    return pairs


def compute_axis_limits(df: pd.DataFrame, pair: PlotPair) -> AxisLimits:
    xy = df.iloc[:, [pair.x_col, pair.y_col]].dropna(how="any")
    if xy.empty:
        raise ValueError(f"{pair.x_title} / {pair.y_title} 没有有效 X-Y 数据。")

    x = pd.to_numeric(xy.iloc[:, 0], errors="coerce").dropna()
    y = pd.to_numeric(xy.iloc[:, 1], errors="coerce").dropna()
    if x.empty or y.empty:
        raise ValueError(f"{pair.x_title} / {pair.y_title} 没有有效数值。")

    raw_x_min, raw_x_max = float(x.min()), float(x.max())
    raw_y_min, raw_y_max = float(y.min()), float(y.max())

    if raw_x_min == raw_x_max:
        raw_x_min -= 0.5
        raw_x_max += 0.5
    if raw_y_min == raw_y_max:
        raw_y_min -= 0.5
        raw_y_max += 0.5

    if AUTO_X_RANGE:
        x_span = raw_x_max - raw_x_min
        x_pad = x_span * AXIS_PADDING_RATIO
        x_step = nice_number((x_span + 2 * x_pad) / AUTO_TICK_TARGET_INTERVALS, round_value=True)
        x_min = nice_floor(raw_x_min - x_pad, x_step)
        x_max = nice_ceil(raw_x_max + x_pad, x_step)
    else:
        x_min, x_max, x_step = MANUAL_X_MIN, MANUAL_X_MAX, MANUAL_X_STEP

    if AUTO_Y_RANGE:
        y_span = raw_y_max - raw_y_min
        y_pad = y_span * AXIS_PADDING_RATIO
        y_step = nice_number((y_span + 2 * y_pad) / AUTO_TICK_TARGET_INTERVALS, round_value=True)
        y_min = nice_floor(raw_y_min - y_pad, y_step)
        y_max = nice_ceil(raw_y_max + y_pad, y_step)
    else:
        y_min, y_max, y_step = MANUAL_Y_MIN, MANUAL_Y_MAX, MANUAL_Y_STEP

    x_decimals = infer_decimals(x_step)
    y_decimals = infer_decimals(y_step)

    print("  自动坐标范围:")
    print(f"    X 原始范围: {raw_x_min:g} – {raw_x_max:g}")
    print(f"    X 绘图范围: {x_min:g} – {x_max:g}, 主刻度步长: {x_step:g}")
    print(f"    Y 原始范围: {raw_y_min:g} – {raw_y_max:g}")
    print(f"    Y 绘图范围: {y_min:g} – {y_max:g}, 主刻度步长: {y_step:g}")

    return AxisLimits(
        x_min=x_min,
        x_max=x_max,
        x_step=x_step,
        x_decimals=x_decimals,
        y_min=y_min,
        y_max=y_max,
        y_step=y_step,
        y_decimals=y_decimals,
    )


def normalize_marker_symbol(symbol: Union[str, int]) -> int:
    if isinstance(symbol, int):
        return symbol
    symbol_text = str(symbol).strip().lower().replace("-", "_").replace(" ", "_")
    if symbol_text.isdigit():
        return int(symbol_text)
    if symbol_text not in SYMBOL_MAP:
        raise ValueError(
            f"未知散点类型：{symbol}\n"
            f"可选英文名：{', '.join(sorted(SYMBOL_MAP))}；也可以直接输入 Origin symbol 编号。"
        )
    return SYMBOL_MAP[symbol_text]


def connect_origin() -> None:
    print("\n连接 Origin 2023b ...")
    try:
        if op.oext:
            print("  Origin 已在运行，将使用当前 Origin 实例。")
        else:
            print("  正在启动 Origin，请稍候 ...")
            op.set_origin_instance()
    except Exception as exc:
        raise RuntimeError(
            "无法连接 Origin。请确认：\n"
            "1. 当前是在 Windows 本地环境运行，不是在 WSL/远程 Linux；\n"
            "2. 本机已安装并注册 Origin 2023b / OriginPro 2023b；\n"
            "3. 当前 Python 环境已安装 originpro；\n"
            "4. 如果电脑上有多个 Origin 版本，请先让 Origin 2023b 成为默认 COM/Automation 版本。"
        ) from exc


def write_pair_to_origin(df: pd.DataFrame, pair: PlotPair):
    xy = df.iloc[:, [pair.x_col, pair.y_col]].dropna(how="any").copy()
    xy.columns = [pair.x_title, pair.y_title]

    wks = op.new_sheet("w", lname="Scatter_Data")
    wks.from_df(xy)
    wks.set_labels([pair.x_title, pair.y_title], "L")
    print(f"  工作表已创建: {wks.name}, {wks.shape[0]} 行 × {wks.shape[1]} 列")
    return wks


def safe_plot_set(plot, method_name: str, prop: str, value) -> bool:
    """兼容不同 originpro 版本，设置失败时不中断。"""
    try:
        method = getattr(plot, method_name)
        method(prop, value)
        return True
    except Exception:
        return False


def apply_marker_style(gl, plot, marker_size: float, marker_symbol: Union[str, int], marker_color: str) -> None:
    """设置散点大小、类型和颜色。"""
    symbol_id = normalize_marker_symbol(marker_symbol)

    # originpro API：颜色通常最稳定。
    try:
        plot.color = marker_color
    except Exception:
        pass

    # 尝试多组常见属性名，兼容不同 Origin/originpro 版本。
    for prop in ("symbol.size", "sym.size", "size"):
        if safe_plot_set(plot, "set_float", prop, float(marker_size)):
            break

    for prop in ("symbol.kind", "symbol.type", "symbol.shape", "sym.kind", "sym.shape"):
        if safe_plot_set(plot, "set_int", prop, int(symbol_id)):
            break

    # LabTalk 兜底：%C 通常代表当前激活数据图。
    # -k：symbol shape；-z：symbol size；-c：颜色。
    # 对十六进制颜色，plot.color 已经设置；LabTalk 颜色兜底主要用于 red/blue/black 等名称。
    try:
        lt(gl, f"set %C -k {int(symbol_id)}")
        lt(gl, f"set %C -z {float(marker_size)}")
        if not str(marker_color).strip().startswith("#"):
            lt(gl, f"set %C -c color({marker_color})")
    except Exception:
        pass

    # 确保为纯散点，不连线。
    try:
        lt(gl, "set %C -l 0")
    except Exception:
        pass


def create_scatter_plot(wks, pair: PlotPair, marker_size: float, marker_symbol: Union[str, int], marker_color: str):
    print("创建二维散点图 ...")
    gp = op.new_graph(template="scatter")
    gp.activate()
    gl = gp[0]
    gl.activate()

    plot = gl.add_plot(wks, colx=0, coly=1, type="s")
    apply_marker_style(gl, plot, marker_size, marker_symbol, marker_color)

    try:
        gl.rescale()
    except Exception:
        pass

    return gp, gl


def format_axes(gp, gl, pair: PlotPair, axis_limits: AxisLimits) -> None:
    gp.activate()
    gl.activate()

    gl.set_xlim(axis_limits.x_min, axis_limits.x_max, axis_limits.x_step)
    gl.set_ylim(axis_limits.y_min, axis_limits.y_max)

    # 双保险：通过 LabTalk 设置范围、刻度和数值标签小数位。
    lt(gl, "layer.x.type = 1")
    lt(gl, f"layer.x.from = {axis_limits.x_min}")
    lt(gl, f"layer.x.to = {axis_limits.x_max}")
    lt(gl, f"layer.x.inc = {axis_limits.x_step}")
    lt(gl, f"layer.y.from = {axis_limits.y_min}")
    lt(gl, f"layer.y.to = {axis_limits.y_max}")
    lt(gl, f"layer.y.inc = {axis_limits.y_step}")

    x_title_plain = plain_title_for_origin(pair.x_title)
    y_title_plain = plain_title_for_origin(pair.y_title)
    x_title_lt = origin_rich_text(pair.x_title, bold=True)
    y_title_lt = origin_rich_text(pair.y_title, bold=True)

    gl.axis("x").title = x_title_plain
    gl.axis("y").title = y_title_plain
    lt(gl, f'label -xb "{x_title_lt}"')
    lt(gl, f'label -yl "{y_title_lt}"')
    lt(gl, f'xb.text$ = "{x_title_lt}"')
    lt(gl, f'yl.text$ = "{y_title_lt}"')

    if SHOW_TOP_RIGHT_AXES:
        lt(gl, "layer.x.showAxes = 3")
        lt(gl, "layer.y.showAxes = 3")
        lt(gl, "layer.x.showLabels = 1")
        lt(gl, "layer.y.showLabels = 1")
        lt(gl, "layer.x2.showlabel = 0")
        lt(gl, "layer.y2.showlabel = 0")
    else:
        lt(gl, "layer.x.showAxes = 1")
        lt(gl, "layer.y.showAxes = 1")

    lt(gl, f"layer.x.thickness = {AXIS_WIDTH}")
    lt(gl, f"layer.y.thickness = {AXIS_WIDTH}")
    lt(gl, f"layer.x2.thickness = {AXIS_WIDTH}")
    lt(gl, f"layer.y2.thickness = {AXIS_WIDTH}")

    lt(gl, "layer.x.ticks = 1")
    lt(gl, "layer.y.ticks = 1")
    lt(gl, f"layer.x.tickLength = {TICK_LENGTH}")
    lt(gl, f"layer.y.tickLength = {TICK_LENGTH}")
    lt(gl, "layer.x.minorTicks = 0")
    lt(gl, "layer.y.minorTicks = 0")

    lt(gl, "layer.x.labelType = 1")
    lt(gl, "layer.x.labelSubtype = 1")
    lt(gl, f"layer.x.decPlaces = {axis_limits.x_decimals}")
    lt(gl, "layer.y.labelType = 1")
    lt(gl, "layer.y.labelSubtype = 1")
    lt(gl, f"layer.y.decPlaces = {axis_limits.y_decimals}")

    for axis in ("x", "y"):
        lt(gl, f'layer.{axis}.label.font = font("{FONT_NAME}")')
        lt(gl, f"layer.{axis}.label.fSize = {FONT_SIZE}")
        lt(gl, f"layer.{axis}.label.bold = 1")

    for obj in ("xb", "XB", "yl", "YL"):
        lt(gl, f'{obj}.font = font("{FONT_NAME}")')
        lt(gl, f"{obj}.fsize = {FONT_SIZE}")
        lt(gl, f"{obj}.bold = 1")

    if not SHOW_LEGEND:
        try:
            gl.label("Legend").remove()
        except Exception:
            pass

    lt(gl, "doc -uw")
    print("  坐标轴、标题、散点样式设置完成")


def export_outputs(gp, figures_dir: Path, origin_dir: Path, output_stem: str):
    output_stem = sanitize_filename(output_stem)
    emf_path = figures_dir / f"{output_stem}.emf"
    opj_path = origin_dir / f"{output_stem}.opju"

    print(f"导出 EMF 矢量图: {emf_path}")
    gp.lt_exec(
        f'expgraph type:=emf '
        f'filename:="{output_stem}" '
        f'path:="{str(figures_dir).replace(chr(92), chr(47))}" '
        f'overwrite:=replace'
    )
    print("  EMF 导出完成")

    print(f"保存 Origin 工程: {opj_path}")
    gp.lt_exec(f'save -i "{str(opj_path).replace(chr(92), chr(47))}"')
    print("  OPJU 保存完成")

    return emf_path, opj_path


def plot_one_pair(
    df: pd.DataFrame,
    pair: PlotPair,
    output_stem: str,
    figures_dir: Path,
    origin_dir: Path,
    marker_size: float,
    marker_symbol: Union[str, int],
    marker_color: str,
):
    """绘制一个二维散点图并导出。"""
    op.new()
    axis_limits = compute_axis_limits(df, pair)
    wks = write_pair_to_origin(df, pair)
    gp, gl = create_scatter_plot(wks, pair, marker_size, marker_symbol, marker_color)
    format_axes(gp, gl, pair, axis_limits)
    return export_outputs(gp, figures_dir, origin_dir, output_stem)


def plot_dataset(
    df: pd.DataFrame,
    output_prefix: str,
    figures_dir: Path,
    origin_dir: Path,
    layout: str,
    marker_size: float,
    marker_symbol: Union[str, int],
    marker_color: str,
) -> tuple[list[str], list[tuple[str, str]]]:
    pairs = infer_plot_pairs(df, layout=layout)
    success: list[str] = []
    failed: list[tuple[str, str]] = []

    for i, pair in enumerate(pairs):
        print("\n" + "-" * 60)
        output_stem = sanitize_filename(f"{output_prefix}__{pair.output_suffix}")
        color = COLORS_FOR_MULTIPLE_PLOTS[i % len(COLORS_FOR_MULTIPLE_PLOTS)] if AUTO_CYCLE_COLORS else marker_color
        try:
            plot_one_pair(
                df=df,
                pair=pair,
                output_stem=output_stem,
                figures_dir=figures_dir,
                origin_dir=origin_dir,
                marker_size=marker_size,
                marker_symbol=marker_symbol,
                marker_color=color,
            )
            success.append(output_stem)
        except Exception as exc:
            failed.append((output_stem, str(exc)))
            print(f"失败: {output_stem}")
            print(f"错误: {exc}")

    return success, failed


def plot_single(paths: ProjectPaths, input_file: Optional[Path], layout: str, marker_size: float, marker_symbol: Union[str, int], marker_color: str):
    data_path = input_file if input_file is not None else get_single_data_path(paths)
    df = read_scatter_data(data_path)
    output_prefix = sanitize_filename(data_path.stem, fallback="scatter")
    return plot_dataset(df, output_prefix, paths.figures_dir, paths.origin_dir, layout, marker_size, marker_symbol, marker_color)


def plot_batch_files(paths: ProjectPaths, batch_dir: Optional[Path], layout: str, marker_size: float, marker_symbol: Union[str, int], marker_color: str):
    data_dir = batch_dir if batch_dir is not None else paths.batch_data_dir
    data_files = find_data_files(data_dir)
    if not data_files:
        raise FileNotFoundError(f"未在批量目录中找到数据文件：{data_dir}")

    all_success: list[str] = []
    all_failed: list[tuple[str, str]] = []

    print(f"\n批量绘图：共找到 {len(data_files)} 个数据文件。")
    for data_path in data_files:
        print("\n" + "=" * 70)
        print(f"处理文件: {data_path.name}")
        try:
            df = read_scatter_data(data_path)
            output_prefix = sanitize_filename(data_path.stem, fallback="scatter")
            success, failed = plot_dataset(
                df=df,
                output_prefix=output_prefix,
                figures_dir=paths.batch_figures_dir,
                origin_dir=paths.batch_origin_dir,
                layout=layout,
                marker_size=marker_size,
                marker_symbol=marker_symbol,
                marker_color=marker_color,
            )
            all_success.extend(success)
            all_failed.extend(failed)
        except Exception as exc:
            all_failed.append((data_path.name, str(exc)))
            print(f"失败: {data_path.name}")
            print(f"错误: {exc}")

    return all_success, all_failed


def plot_workbook_sheets(paths: ProjectPaths, input_file: Optional[Path], layout: str, marker_size: float, marker_symbol: Union[str, int], marker_color: str):
    workbook_path = input_file if input_file is not None else get_single_data_path(paths)
    if not workbook_path.exists():
        raise FileNotFoundError(f"未找到多 Sheet 绘图数据文件：{workbook_path}")
    if workbook_path.suffix.lower() not in (".xlsx", ".xlsm", ".xls"):
        raise ValueError("--mode sheets 仅适用于 Excel 文件；TXT/CSV 文件请使用 --mode single 或 --mode batch。")

    sheet_names = pd.ExcelFile(workbook_path).sheet_names
    if not sheet_names:
        raise ValueError(f"数据文件中没有 sheet：{workbook_path}")

    all_success: list[str] = []
    all_failed: list[tuple[str, str]] = []

    print(f"\n多 Sheet 绘图：{workbook_path.name} 共找到 {len(sheet_names)} 个 sheet。")
    for sheet_name in sheet_names:
        print("\n" + "=" * 70)
        print(f"处理 Sheet: {sheet_name}")
        try:
            df = read_scatter_data(workbook_path, sheet_name=sheet_name)
            output_prefix = sanitize_filename(f"{workbook_path.stem}__{sheet_name}", fallback="scatter")
            success, failed = plot_dataset(
                df=df,
                output_prefix=output_prefix,
                figures_dir=paths.batch_figures_dir,
                origin_dir=paths.batch_origin_dir,
                layout=layout,
                marker_size=marker_size,
                marker_symbol=marker_symbol,
                marker_color=marker_color,
            )
            all_success.extend(success)
            all_failed.extend(failed)
        except Exception as exc:
            all_failed.append((sheet_name, str(exc)))
            print(f"失败: {sheet_name}")
            print(f"错误: {exc}")

    return all_success, all_failed


def print_summary(success: Iterable[str], failed: Iterable[tuple[str, str]]) -> None:
    success = list(success)
    failed = list(failed)

    print("\n" + "=" * 70)
    print("绘图任务完成")
    print(f"成功: {len(success)} 个")
    print(f"失败: {len(failed)} 个")

    if success:
        print("\n成功列表：")
        for item in success:
            print(f"  - {item}")

    if failed:
        print("\n失败列表：")
        for item, err in failed:
            print(f"  - {item}: {err}")

    print("=" * 70)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="调用 Origin 2023b 批量绘制二维散点图，坐标轴标题自动来自第一行表头。"
    )
    parser.add_argument(
        "--mode",
        choices=("auto", "single", "batch", "sheets"),
        default="auto",
        help="绘图模式：auto=自动判断；single=单文件；batch=批量文件；sheets=单 Excel 多 sheet。",
    )
    parser.add_argument(
        "--layout",
        choices=("auto", "pairs", "common-x"),
        default="auto",
        help="数据布局：pairs=X1/Y1、X2/Y2；common-x=第一列公共 X；auto=偶数列 pairs，奇数列 common-x。",
    )
    parser.add_argument(
        "--input-file",
        type=Path,
        default=None,
        help="指定单文件或多 Sheet 绘图的数据文件；未提供时按 data/ 自动扫描。",
    )
    parser.add_argument(
        "--batch-dir",
        type=Path,
        default=None,
        help="指定批量绘图的数据目录；未提供时使用 data/batch/。",
    )
    parser.add_argument(
        "--marker-size",
        type=float,
        default=MARKER_SIZE,
        help=f"散点大小，默认 {MARKER_SIZE}。",
    )
    parser.add_argument(
        "--marker-symbol",
        type=str,
        default=str(MARKER_SYMBOL),
        help=(
            "散点类型，可用 circle/square/triangle_up/triangle_down/diamond/plus/cross/star，"
            "也可输入 Origin symbol 编号。"
        ),
    )
    parser.add_argument(
        "--marker-color",
        type=str,
        default=MARKER_COLOR,
        help=f"散点颜色，支持 '#RRGGBB' 或 red/blue/black 等名称，默认 {MARKER_COLOR}。",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    paths = get_project_paths()
    input_file = args.input_file.resolve() if args.input_file is not None else None
    batch_dir = args.batch_dir.resolve() if args.batch_dir is not None else None

    connect_origin()

    mode = args.mode
    if mode == "auto":
        batch_files = find_data_files(batch_dir or paths.batch_data_dir)
        mode = "batch" if batch_files else "single"
        print(f"自动模式选择: {mode}")

    try:
        if mode == "single":
            success, failed = plot_single(
                paths=paths,
                input_file=input_file,
                layout=args.layout,
                marker_size=args.marker_size,
                marker_symbol=args.marker_symbol,
                marker_color=args.marker_color,
            )
        elif mode == "batch":
            success, failed = plot_batch_files(
                paths=paths,
                batch_dir=batch_dir,
                layout=args.layout,
                marker_size=args.marker_size,
                marker_symbol=args.marker_symbol,
                marker_color=args.marker_color,
            )
        elif mode == "sheets":
            success, failed = plot_workbook_sheets(
                paths=paths,
                input_file=input_file,
                layout=args.layout,
                marker_size=args.marker_size,
                marker_symbol=args.marker_symbol,
                marker_color=args.marker_color,
            )
        else:
            raise ValueError(f"未知模式: {mode}")
    finally:
        # 保持 Origin 打开，方便检查图形和工程。
        # 如果想运行后自动关闭 Origin，可取消下一行注释。
        # op.exit()
        pass

    print_summary(success, failed)


if __name__ == "__main__":
    main()
