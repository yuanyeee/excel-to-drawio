# Excel to Draw.io Converter

[🇯🇵 日本語](README.md) | [🇬🇧 English](README_en.md) | [🇨🇳 简体中文](README_zh.md)

一个把 Excel（.xlsx / .xlsm）工作表转换为 Draw.io（.drawio）格式的工具，
同时提供桌面 GUI 和命令行（CLI）两种使用方式。

## 特性

- 操作简单：选文件、选工作表、点转换
- 支持多工作表批量转换
- 转换绘图形状：矩形、椭圆、菱形、连接器等
- 还原单元格：填充色、边框、文本标签、合并单元格
- 保留样式：填充色、线条颜色、字体样式
- 图片以 base64 数据 URI 内嵌
- 每个转换使用独立的 Theme 对象（无全局可变状态）

## 运行环境

- Python 3.8+
- 无需第三方依赖（仅用标准库）
- GUI 需要 tkinter（Python 自带）

## 安装

```bash
pip install .
```

安装后可获得 `excel-to-drawio`（CLI）和 `excel-to-drawio-gui`（GUI）两个命令。
也可以不安装直接运行 `python -m excel_to_drawio`。

## 使用方法

### 命令行（CLI）

```bash
# 转换所有工作表（输出文件名自动生成）
python -m excel_to_drawio input.xlsx

# 指定输出文件
python -m excel_to_drawio input.xlsx -o output.drawio

# 只转换指定工作表
python -m excel_to_drawio input.xlsx -s "Sheet1" "Sheet2"

# 列出工作表
python -m excel_to_drawio input.xlsx -l
```

| 选项 | 说明 |
|---|---|
| `input`（必填） | 输入 Excel 文件（.xlsx / .xlsm） |
| `-o`, `--output` | 输出路径（默认 `<输入文件名>.drawio`） |
| `-s`, `--sheets` | 要转换的工作表名（默认全部） |
| `-l`, `--list` | 列出工作表后退出 |
| `--no-images` / `--no-borders` / `--no-fills` / `--no-labels` / `--no-shapes` | 关闭某一渲染 |
| `--no-merge-fills` | 关闭同色填充合并 |
| `--skip-hidden` | 跳过隐藏行列 |
| `--no-page-mode` | 输出 `page="0"`（默认 `page="1"`） |
| `--scale` | 缩放系数（默认 1.0） |

### 桌面 GUI

```bash
python -m excel_to_drawio.desktop_app
# 或:
excel-to-drawio-gui
```

如果想免安装直接启动 GUI，也可以使用各操作系统对应的启动脚本，
可直接在仓库根目录下运行。

| 操作系统 | 运行方式 |
|---|---|
| Windows | 双击 `run_gui.bat` |
| macOS | 双击 `run_gui.command`（首次运行可能需要右键→打开） |
| Linux | 在终端执行 `./run_gui.sh` |

每个脚本都会检测 Python 和 tkinter 是否可用，缺失时会给出安装提示。

换行符已在 `.gitattributes` 中固定：`run_gui.bat` 始终为 CRLF，`.sh` 与 `.command`
始终为 LF，无论是 clone 还是下载 ZIP 都一致。

## 项目结构

```
excel-to-drawio/
├── excel_to_drawio/          # Python 包
│   ├── __init__.py           # 公开 API
│   ├── __main__.py           # CLI 入口
│   ├── desktop_app.py        # tkinter GUI
│   ├── config.py             # ConvertConfig
│   ├── constants.py          # OOXML 命名空间与查找表
│   ├── colors.py             # Theme 与颜色解析
│   ├── grid.py               # 单元格坐标辅助
│   ├── ooxml.py              # 底层 OOXML 读取
│   ├── geometry.py           # DrawingML 几何辅助
│   ├── builder.py            # Drawio XML 构建器
│   ├── styles.py             # 单元格样式 / 填充 / 边框 / 标签
│   ├── images.py             # 图片提取
│   ├── connectors.py         # 连接器渲染
│   ├── shapes.py             # 形状渲染
│   └── convert.py            # 转换编排
├── run_gui.bat               # GUI 启动脚本 (Windows)
├── run_gui.command           # GUI 启动脚本 (macOS)
├── run_gui.sh                # GUI 启动脚本 (Linux)
├── pyproject.toml            # 打包配置
├── LICENSE                   # MIT
└── README.md                 # 本文件
```

## 许可证

MIT — 见 [LICENSE](LICENSE)。
