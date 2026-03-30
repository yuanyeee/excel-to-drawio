# Excel to Draw.io Converter

Excel ファイル(.xlsx / .xlsm)のシート内容を Draw.io (.drawio) 形式に変換するツールです。
デスクトップ GUI とコマンドライン (CLI) の両方で利用できます。

## 特徴

- **シンプルな操作** : ファイルを選んで、シートを選択して、変換ボタンを押すだけ
- **複数シート対応** : 複数シートを一括で変換可能
- **描画図形変換** : Excel上の図形(矩形、楕円、菱形、コネクターなど)を Draw.io 形式に変換
- **セル情報変換** : セルの塗りつぶし色、罫線、テキストラベル、結合セルを再現
- **スタイル保持** : 塗りつぶし色、線色、フォントスタイルを保持

## 対応ファイル形式

| 形式 | 説明 | 対応 |
|------|------|------|
| .xlsx | Excel 2007以降 | ✅ |
| .xlsm | Excel マクロ付き | ✅ |

## インストール

```powershell
cd excel-to-drawio
pip install -r requirements.txt
```

依存ライブラリ: `openpyxl`

## 使い方

### デスクトップ GUI

`ExcelToDrawIO/` フォルダ内で以下を実行します。

```powershell
cd ExcelToDrawIO
python desktop_app.py
```

1. **「Browse...」** をクリックして Excel ファイルを選択
2. 読み込まれた **シート一覧から変換対象を選択**（複数選択可）
3. **「Output」** に出力先が自動設定される（「Save As...」で変更可）
4. **「Convert」** をクリックして変換実行
5. 完了すると `.drawio` ファイルが保存される

詳細な操作説明は [ExcelToDrawIO/docs/desktop_app_usage.md](ExcelToDrawIO/docs/desktop_app_usage.md) を参照してください。

### コマンドライン (CLI)

`ExcelToDrawIO/` フォルダ内で以下を実行します。

```bash
# 全シートを変換（出力ファイル名は自動生成）
python excel_to_drawio.py input.xlsx

# 出力ファイルを指定
python excel_to_drawio.py input.xlsx -o output.drawio

# 特定のシートのみ変換
python excel_to_drawio.py input.xlsx -s "Sheet1" "Sheet2"

# シート一覧を表示
python excel_to_drawio.py input.xlsx -l
```

| オプション | 説明 |
|---|---|
| `input` (必須) | 入力 Excel ファイル (.xlsx / .xlsm) |
| `-o`, `--output` | 出力ファイルパス (省略時は `入力ファイル名.drawio`) |
| `-s`, `--sheets` | 変換するシート名を指定 (省略時は全シート) |
| `-l`, `--list` | シート一覧を表示して終了 |

## プロジェクト構成

```
excel-to-drawio/
├── ExcelToDrawIO/
│   ├── excel_to_drawio.py   # 変換エンジン本体 (CLI としても実行可能)
│   ├── desktop_app.py       # デスクトップ GUI アプリ (tkinter)
│   └── docs/
│       └── desktop_app_usage.md  # GUI 操作マニュアル
├── requirements.txt          # 依存ライブラリ
├── .gitignore
└── README.md
```

## 動作環境

- Python 3.8+
- Windows / Mac / Linux
- tkinter (Python 標準ライブラリ)

## トラブルシューティング

### tkinter のエラーが出る
Python が正しくインストールされているか確認してください。公式サイトから Python 3.8 以降をインストールしてください。

### Excel ファイルが開けない
ファイルが別のプログラムで開かれていないか確認してください。閉じてから再度試してください。

### 変換結果が Draw.io で開けない
[Draw.io 公式サイト](https://www.drawio.com/) から最新版をダウンロードしてお試しください。

## ライセンス

MIT License
