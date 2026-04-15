# Excel to Draw.io Converter

[🇯🇵 日本語](README.md) | [🇬🇧 English](README_en.md) | [🇨🇳 简体中文](README_zh.md)

将 Excel 文件 (.xlsx / .xlsm) 的工作表内容转换为 Draw.io (.drawio) 格式的工具。
提供桌面 GUI 和命令行（CLI）两种使用方式。

## 特性

- **操作简单**：只需选择文件、选择工作表并点击转换按钮。
- **支持多工作表**：可一次性批量转换多个工作表。
- **图形转换**：将 Excel 中的图形（矩形、椭圆、菱形、连接线等）转换为 Draw.io 格式。
- **单元格信息转换**：还原单元格的填充色、边框、文本标签和合并单元格。
- **保持样式**：保留原有的填充色、线条颜色和字体样式。

## 支持的文件格式

| 格式 | 说明 | 支持 |
|------|------|------|
| .xlsx | Excel 2007 及以上版本 | ✅ |
| .xlsm | 包含宏的 Excel 文件 | ✅ |

## 安装

```powershell
cd excel-to-drawio
pip install -r requirements.txt
```

依赖库: `openpyxl`

## 使用方法

### 桌面 GUI

在 `ExcelToDrawIO/` 文件夹内运行以下命令：

```powershell
cd ExcelToDrawIO
python desktop_app.py
```

1. 点击 **"Browse..."** 选择 Excel 文件。
2. 从加载的 **工作表列表中选择要转换的表**（支持多选）。
3. **"Output"** 中会自动设置输出路径（可通过 "Save As..." 修改）。
4. 点击 **"Convert"** 执行转换。
5. 完成后将保存为 `.drawio` 文件。

详细的操作说明请参考 [ExcelToDrawIO/docs/desktop_app_usage_zh.md](ExcelToDrawIO/docs/desktop_app_usage_zh.md)。

### 命令行（CLI）

在 `ExcelToDrawIO/` 文件夹内运行以下命令：

```bash
# 转换所有工作表（自动生成输出文件名）
python excel_to_drawio.py input.xlsx

# 指定输出文件
python excel_to_drawio.py input.xlsx -o output.drawio

# 仅转换指定工作表
python excel_to_drawio.py input.xlsx -s "Sheet1" "Sheet2"

# 显示工作表列表
python excel_to_drawio.py input.xlsx -l
```

| 选项 | 说明 |
|---|---|
| `input` (必需) | 输入的 Excel 文件 (.xlsx / .xlsm) |
| `-o`, `--output` | 输出文件路径 (省略时默认为 `输入文件名.drawio`) |
| `-s`, `--sheets` | 指定要转换的工作表名称 (省略时转换所有工作表) |
| `-l`, `--list` | 显示工作表列表并退出 |

## 项目结构

```
excel-to-drawio/
├── ExcelToDrawIO/
│   ├── excel_to_drawio.py   # 转换引擎核心 (也可作为 CLI 运行)
│   ├── desktop_app.py       # 桌面 GUI 应用 (tkinter)
│   └── docs/
│       └── desktop_app_usage.md  # GUI 操作手册
├── requirements.txt          # 依赖库
├── .gitignore
└── README.md
```

## 运行环境

- Python 3.8+
- Windows / Mac / Linux
- tkinter (Python 标准库)

## 常见问题解答

### 出现 tkinter 错误
请确认 Python 是否已正确安装。请从官网安装 Python 3.8 或更高版本。

### 无法打开 Excel 文件
请检查文件是否被其他程序占用，关闭后重试。

### 转换结果无法在 Draw.io 中打开
请从 [Draw.io 官网](https://www.drawio.com/) 下载最新版本后重试。

## 许可证

MIT License
