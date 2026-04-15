# Excel to Draw.io Desktop Tool 运行说明

[🇯🇵 日本語](desktop_app_usage.md) | [🇬🇧 English](desktop_app_usage_en.md) | [🇨🇳 简体中文](desktop_app_usage_zh.md)

## 概述

此工具是一个 Python 桌面应用程序，用于将 Excel 文件中的指定工作表转换为 Draw.io 格式。

- 支持的格式: `.xlsx`, `.xlsm`
- UI: `tkinter`
- 操作: 选择文件 → 选择工作表 → 选择选项 → 转换

## 运行前提

- Windows
- 已安装 Python 3
- 此文件夹中必须存在以下文件：
  - `desktop_app.py`
  - `excel_to_drawio.py`

## 启动方法

在工作目录中运行以下命令：

```powershell
python .\desktop_app.py
```

## 使用方法

1. 点击 `Browse...` 选择要转换的 Excel 文件。
2. 从加载的工作表列表中选择要转换的目标工作表。
3. `Output` 中会自动填充输出的 `.drawio` 文件名。
4. 如有需要，点击 `Save As...` 更改保存路径。
5. 点击 `Convert` 进行转换。
6. 转换成功后，将在目标路径生成一个 `.drawio` 文件。

## 主要选项

- Include images: 启用/禁用图像嵌入
- Include borders: 启用/禁用边框绘制
- Merge same-color fills: 将相邻同色填充单元格合并绘制
- Skip hidden rows/cols: 排除隐藏行/列

## 界面说明

- `Excel File`
  - 输入文件路径
- `Sheets`
  - 工作簿中的工作表列表
  - 工作表过多时可滚动
- `Output`
  - 输出的 `.drawio` 文件路径
- `Convert`
  - 执行转换
- 底部日志栏
  - 显示加载结果、转换结果和错误详情

## 注意事项

- 此版本不支持 `.xls` 格式。
- 转换是同步执行的。处理较大的文件可能需要一些时间。
- 取决于工作表，Excel 特有的复杂图表形状可能无法在 Draw.io 中完美匹配。
- 如果输入文件或输出路径存在访问权限问题，将显示错误对话框。

## 常见操作

### 自动决定输出路径
选择工作表后，将自动设置为在同一文件夹中保存为 `工作表名称.drawio`。

### 另存为其他名称
使用 `Save As...` 指定您想要的保存位置和文件名。

### 发生错误时
请检查屏幕底部的日志栏和弹出的提示信息。
