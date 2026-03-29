# Excel to draw.io Converter

Excel（`.xlsx` / `.xlsm`）の内容を draw.io（`.drawio`）へ変換するツールです。  
GUI（デスクトップ）とCLIの両方に対応しています。

## 主な仕様（最新）

- 変換エンジンは **legacy優先**（既定）
- legacyで図形が取得できず**空出力になる場合は自動でpipelineへフォールバック**
- 複数シートを1つの `.drawio` に複数ページとして出力
- `--include-cells` 有効時はセル由来図形（塗り/枠線/結合セル）を出力
- ノイズ低減のため「罫線のみ・塗りなし・値なし」セルは除外

---

## インストール

```bash
pip install -r requirements.txt
```

---

## 使い方

### 1) GUI（推奨）

```bash
python gui_tkinter.py
```

または

```bash
python desktop_app.py
```

手順:
1. Excelファイルを選択
2. 変換対象シートを選択
3. `Convert to draw.io` 実行
4. 保存先を指定

> 現在のGUIにエンジン選択UIはありません（内部で legacy優先 + 必要時pipelineフォールバック）。

### 2) CLI

```bash
python main.py input.xlsx -o output.drawio --sheets "Sheet1" "Sheet2" --include-cells
```

オプション:
- `-o, --output`: 出力 `.drawio` パス
- `-s, --sheets`: 変換するシート名（複数指定可）
- `--include-cells`: セル由来図形も含める
- `-v, --verbose`: 詳細ログ

---

## Python API

```python
from converter.excel_to_drawio import convert_excel_to_drawio

result = convert_excel_to_drawio(
    input_path="input.xlsx",
    output_path="output.drawio",
    sheet_names=["Sheet1"],
    include_cells=True,
    engine="legacy",   # "legacy" or "pipeline"
)
```

`engine` を指定しない場合は `legacy` が使われ、空結果時に `pipeline` へ自動フォールバックします。

---

## プロジェクト構成

```text
excel-to-drawio/
├── desktop_app.py
├── gui_tkinter.py
├── main.py
├── converter/
│   ├── excel_to_drawio.py
│   ├── excel_reader.py
│   ├── drawio_writer.py
│   ├── shape_mapper.py
│   └── cell_border.py
├── tests/
└── requirements.txt
```

---

## トラブルシューティング

- `python -m py_compile converter/excel_to_drawio.py` が失敗する場合  
  → ファイルが壊れている可能性があります。最新ブランチから `converter/excel_to_drawio.py` を取り直してください。
- GUIが起動しない場合  
  → `python -m py_compile gui_tkinter.py` で構文エラーを確認してください。

---

## ライセンス

MIT
