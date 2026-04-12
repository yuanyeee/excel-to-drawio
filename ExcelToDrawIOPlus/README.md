# Excel to Draw.io Converter Plus

Excel (.xlsx / .xlsm) を Draw.io (.drawio) 形式に変換する**汎用ツール**です。
既存の `ExcelToDrawIO` をベースに、画像対応・罫線強化・汎用化を行った改良版です。

## 既存ツールとの違い

| 機能 | ExcelToDrawIO (既存) | ExcelToDrawIOPlus (本ツール) |
|------|---------------------|------------------------------|
| 画像 (pic) | スキップ | base64埋め込みで再現 |
| 走査範囲 | 固定 (270行x230列) | シートデータから自動検出 |
| 罫線スタイル | thin/medium/thick のみ | dashed/dotted/double/hair も対応 |
| ハイパーリンク | 非対応 | セルリンクを保持 |
| 非表示行/列 | 処理なし | スキップオプション |
| 設定管理 | グローバル変数 | ConvertConfig dataclass |
| テキスト属性 | 基本のみ | 回転・下線・取消線も対応 |
| 用途 | 特定フロー図向け | 汎用 Excel シート変換 |

## インストール

依存ライブラリは不要です（Python 標準ライブラリのみ使用）。

```bash
cd ExcelToDrawIOPlus
python excel_to_drawio_plus.py --help
```

## 使い方

### コマンドライン (CLI)

```bash
# 全シートを変換
python excel_to_drawio_plus.py input.xlsx

# 出力ファイルを指定
python excel_to_drawio_plus.py input.xlsx -o output.drawio

# 特定のシートのみ変換
python excel_to_drawio_plus.py input.xlsx -s "Sheet1" "Sheet2"

# シート一覧を表示
python excel_to_drawio_plus.py input.xlsx -l

# 画像を除外して変換
python excel_to_drawio_plus.py input.xlsx --no-images

# 罫線を除外して変換
python excel_to_drawio_plus.py input.xlsx --no-borders

# 非表示行/列をスキップ
python excel_to_drawio_plus.py input.xlsx --skip-hidden

# 塗りつぶし結合を無効化（セル単位で描画）
python excel_to_drawio_plus.py input.xlsx --no-merge-fills
```

| オプション | 説明 |
|---|---|
| `input` (必須) | 入力 Excel ファイル (.xlsx / .xlsm) |
| `-o`, `--output` | 出力ファイルパス (省略時は自動生成) |
| `-s`, `--sheets` | 変換するシート名を指定 (省略時は全シート) |
| `-l`, `--list` | シート一覧を表示して終了 |
| `--no-images` | 画像の埋め込みを無効化 |
| `--no-borders` | 罫線描画を無効化 |
| `--no-fills` | 塗りつぶし描画を無効化 |
| `--no-merge-fills` | 隣接同色セルの結合を無効化 |
| `--skip-hidden` | 非表示行/列をスキップ |
| `--scale` | 描画スケール (デフォルト: 1.0) |

### デスクトップ GUI

```bash
cd ExcelToDrawIOPlus
python desktop_app_plus.py
```

1. **「Browse...」** をクリックして Excel ファイルを選択
2. **Options** パネルで変換オプションを設定
   - Images: 画像を埋め込むか
   - Shapes: 描画図形を変換するか
   - Fills: セル塗りつぶしを描画するか
   - Borders: 罫線を描画するか
   - Labels: テキストラベルを描画するか
   - Merge fills: 隣接同色セルを結合するか
   - Skip hidden: 非表示行/列をスキップするか
3. **シート一覧** から変換対象を選択（複数選択可）
4. **「Output」** に出力先を設定
5. **「Convert」** をクリックして変換実行

## 対応ファイル形式

| 形式 | 説明 | 対応 |
|------|------|------|
| .xlsx | Excel 2007以降 | OK |
| .xlsm | Excel マクロ付き | OK |

## 動作環境

- Python 3.8+
- Windows / Mac / Linux
- tkinter (GUI 使用時のみ、Python 標準ライブラリ)

## ファイル構成

```
ExcelToDrawIOPlus/
├── excel_to_drawio_plus.py   # 変換エンジン + CLI
├── desktop_app_plus.py       # GUI アプリ (tkinter)
└── README.md                 # 本ドキュメント
```

## ライセンス

MIT License
