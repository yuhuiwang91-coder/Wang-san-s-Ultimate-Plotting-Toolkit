# OriginScatterGUI - 汪桑的无敌绘图包

基于 Origin 2023b 的二维散点图批量绘制工具，提供图形化界面和命令行两种使用方式。

## 功能特点

- **图形界面（GUI）**：基于 tkinter，直观配置绘图参数
- **批量绘图**：支持单文件、批量文件夹、Excel 多 Sheet 三种模式
- **多种出图类型**：散点图、点线图、折线图
- **丰富的散点样式**：16 种散点符号，支持空心圆等特殊样式
- **线性拟合**：支持对每组数据自动进行线性拟合并标注
- **合并出图**：支持多组数据合并到同一绘图区（含多坐标轴）
- **对数坐标 / 倒序轴**：支持 X 轴对数坐标和 Y 轴倒序
- **导出格式**：EMF 矢量图 + Origin OPJU 工程文件

## 依赖环境

- Windows 操作系统
- Python 3.10+
- Origin 2023b / OriginPro 2023b
- 依赖包：`originpro`, `pandas`, `openpyxl`, `pillow`

## 安装

```bash
pip install -r requirements.txt
```

## 使用方式

### 图形界面

```bash
python app.py
```

### 命令行

```bash
# 自动模式
python scatter_backend.py --mode auto --input-file data/1.xlsx

# 批量模式
python scatter_backend.py --mode batch --batch-dir data/batch/

# Excel 多 Sheet 模式
python scatter_backend.py --mode excel --input-file data/1.xlsx

# 指定绘图样式
python scatter_backend.py --marker-size 9 --marker-color "#D73027" --marker-symbol circle
```

## 数据格式

### 格式一：两列一组（pairs）
```
X1标题, Y1标题, X2标题, Y2标题, ...
X1数据, Y1数据, X2数据, Y2数据, ...
```

### 格式二：公共 X 列（common-x）
```
X标题, Y1标题, Y2标题, Y3标题, ...
X数据, Y1数据, Y2数据, Y3数据, ...
```

## 打包为 EXE

```bash
# 使用 PyInstaller 打包
pyinstaller OriginScatterGUI.spec

# 或使用构建脚本
build_exe.bat
```

## 项目结构

```
plot painting/
├── app.py                    # GUI 主程序
├── scatter_backend.py        # 后端绘制引擎
├── origin_2d_scatter_batch.py # 命令行批量脚本
├── requirements.txt          # 依赖列表
├── OriginScatterGUI.spec     # PyInstaller 打包配置
├── assets/                   # 图标和资源
├── build_exe.bat             # 构建脚本
├── build_installer.bat       # 安装包构建脚本
├── build_inno_installer.bat  # Inno Setup 安装包
└── installer_assets/         # 安装包资源
```

## 许可证

MIT License
