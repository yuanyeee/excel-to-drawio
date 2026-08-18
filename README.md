# Excel to Draw.io Converter

[🇯🇵 日本語](README.md) | [🇬🇧 English](README_en.md) | [🇨🇳 简体中文](README_zh.md)

Excel ファイル(.xlsx / .xlsm)のシート内容を Draw.io (.drawio) 形式に変換するツールです。
デスクトップ GUI とコマンドライン (CLI) の両方で利用できます。

## 特徴

- **シンプルな操作**: ファイルを選んで、シートを選択して、変換ボタンを押すだけ
- **複数シート対応**: 複数シートを一括で変換可能
- **描画図形変換**: 矩形、楕円、菱形、コネクターなどを Draw.io 形式に変換
- **セル情報変換**: 塗りつぶし色、罫線、テキストラベル、結合セルを再現
- **スタイル保持**: 塗りつぶし色、線色、フォントスタイルを保持
- **画像埋め込み**: base64 データ URI で画像を再現
- **テーマ分離**: 変換ごとの Theme オブジェクト（グローバル状態を持たない）

## 動作環境

- Python 3.8+
- 依存ライブラリ不要（Python 標準ライブラリのみ）
- GUI には tkinter（Python 標準ライブラリ）が必要

## インストール

```bash
pip install .
```

インストールすると `excel-to-drawio` (CLI) と `excel-to-drawio-gui` (GUI) の
コマンドが使えます。インストールせずに `python -m excel_to_drawio` でも実行できます。

## 使い方

### コマンドライン (CLI)

```bash
# 全シートを変換（出力ファイル名は自動生成）
python -m excel_to_drawio input.xlsx

# 出力ファイルを指定
python -m excel_to_drawio input.xlsx -o output.drawio

# 特定のシートのみ変換
python -m excel_to_drawio input.xlsx -s "Sheet1" "Sheet2"

# シート一覧を表示
python -m excel_to_drawio input.xlsx -l
```

| オプション | 説明 |
|---|---|
| `input` (必須) | 入力 Excel ファイル (.xlsx / .xlsm) |
| `-o`, `--output` | 出力ファイルパス (省略時は `<入力ファイル名>.drawio`) |
| `-s`, `--sheets` | 変換するシート名 (省略時は全シート) |
| `-l`, `--list` | シート一覧を表示して終了 |
| `--no-images` / `--no-borders` / `--no-fills` / `--no-labels` / `--no-shapes` | 各描画を無効化 |
| `--no-merge-fills` | 同色塗りの結合を無効化 |
| `--skip-hidden` | 非表示行/列をスキップ |
| `--no-page-mode` | `page="0"` を出力 (既定は `page="1"`) |
| `--scale` | 拡大率 (既定: 1.0) |

### デスクトップ GUI

```bash
python -m excel_to_drawio.desktop_app
# または:
excel-to-drawio-gui
```

## プロジェクト構成

```
excel-to-drawio/
├── excel_to_drawio/          # Python パッケージ
│   ├── __init__.py           # 公開 API
│   ├── __main__.py           # CLI エントリポイント
│   ├── desktop_app.py        # tkinter GUI
│   ├── config.py             # ConvertConfig
│   ├── constants.py          # OOXML 名前空間・ルックアップ表
│   ├── colors.py             # Theme と色解決
│   ├── grid.py               # セル座標ヘルパー
│   ├── ooxml.py              # 低レベル OOXML 読み込み
│   ├── geometry.py           # DrawingML 幾何ヘルパー
│   ├── builder.py            # Drawio XML ビルダー
│   ├── styles.py             # セルスタイル / 塗り / 罫線 / ラベル
│   ├── images.py             # 画像抽出
│   ├── connectors.py         # コネクター描画
│   ├── shapes.py             # 図形描画
│   └── convert.py            # 変換オーケストレーション
├── pyproject.toml            # パッケージング設定
├── LICENSE                   # MIT
└── README.md                 # このファイル
```

## ライセンス

MIT — [LICENSE](LICENSE) を参照してください。
