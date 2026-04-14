# Excel to Draw.io Converter

Excel (.xlsx / .xlsm) を Draw.io (.drawio) 形式に変換する**汎用ツール**です。
画像 (PNG/JPG/SVG/EMF/WMF フォールバック) ・全種類の罫線・テキスト自動フィット・
ハイパーリンク・非表示行/列スキップなどに対応しています。

> 新機能の検証は隣接フォルダ `ExcelToDrawIOPlus/` で行います。安定確認後に本フォルダへ反映する運用です。

## 主な機能ハイライト

- **画像埋め込み**
  - PNG / JPG / GIF / BMP / SVG / WebP をそのまま base64 で埋め込み
  - Office の「アイコン挿入」で追加された SVG は `<a14:svgBlip>` 拡張経由で優先採用
  - EMF / WMF / TIFF など drawio が描画できない形式は、同名の PNG / JPG / SVG フォールバックを自動検索
  - どうしても描画不可の場合は破線プレースホルダを配置してレイアウトを維持
- **セル描画**
  - 塗りつぶしは隣接同色セルをまとめて 1 矩形に結合
  - 罫線は thin/medium/thick 以外に dashed/dotted/double/hair/dashDotDot 等もサポート
  - 同色の塗り領域内では内側の擬似罫線を抑制し、Excel と同じ見た目を再現
- **テキストラベル**
  - セル幅を超える長文は必要に応じてフォントサイズを自動縮小
  - 結合セルでないテキストは右隣の空セル (同色塗り or 塗りなし) へ自動延伸
  - 2 桁数字や `HH:MM` のような短いラベルは中央寄せで表示
  - `rotation` / `underline` / `strikethrough` / `fontFamily` 等の属性を保持
- **グリッド / 座標系**
  - 実データ範囲を自動検出しグリッドを動的サイズ化 (500 行の上限なし)
  - 列幅は OOXML 仕様 `pixels = Truncate(((256*w + Truncate(128/MDW))/256)*MDW)` に準拠
  - `<sheetFormatPr defaultColWidth>` があれば尊重、未指定時は Calibri 11 デフォルトの 64px に一致
- **その他**
  - セルのハイパーリンクを保持
  - 非表示行/列のスキップオプション
  - シート単位 / 複数シート一括変換
  - GUI / CLI 両対応

## インストール

依存ライブラリは不要です（Python 標準ライブラリのみ使用）。

```bash
cd ExcelToDrawIO
python excel_to_drawio.py --help
```

## 使い方

### コマンドライン (CLI)

```bash
# 全シートを変換
python excel_to_drawio.py input.xlsx

# 出力ファイルを指定
python excel_to_drawio.py input.xlsx -o output.drawio

# 特定のシートのみ変換
python excel_to_drawio.py input.xlsx -s "Sheet1" "Sheet2"

# シート一覧を表示
python excel_to_drawio.py input.xlsx -l

# 画像を除外して変換
python excel_to_drawio.py input.xlsx --no-images

# 罫線を除外して変換
python excel_to_drawio.py input.xlsx --no-borders

# 非表示行/列をスキップ
python excel_to_drawio.py input.xlsx --skip-hidden

# 塗りつぶし結合を無効化（セル単位で描画）
python excel_to_drawio.py input.xlsx --no-merge-fills
```

| オプション | 説明 |
|---|---|
| `input` (必須) | 入力 Excel ファイル (.xlsx / .xlsm) |
| `-o`, `--output` | 出力ファイルパス (省略時は自動生成) |
| `-s`, `--sheets` | 変換するシート名を指定 (省略時は全シート) |
| `-l`, `--list` | シート一覧を表示して終了 |
| `--no-images` | 画像の埋め込みを無効化 |
| `--no-shapes` | 描画図形 (drawing) の変換を無効化 |
| `--no-borders` | 罫線描画を無効化 |
| `--no-fills` | 塗りつぶし描画を無効化 |
| `--no-labels` | テキストラベル描画を無効化 |
| `--no-merge-fills` | 隣接同色セルの結合を無効化 |
| `--skip-hidden` | 非表示行/列をスキップ |
| `--scale` | 描画スケール (デフォルト: 1.0) |

### デスクトップ GUI

```bash
cd ExcelToDrawIO
python desktop_app.py
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

### 画像形式の扱い

| 形式 | 対応 | 備考 |
|---|---|---|
| PNG / JPG / GIF / BMP / WebP | そのまま埋め込み | data URI |
| SVG | そのまま埋め込み | `<a14:svgBlip>` 拡張も認識 |
| EMF / WMF / TIFF | 同名 PNG/JPG/SVG フォールバックを検索 | 見つからない場合はプレースホルダ |
| OLE オブジェクト | 非対応 | (外部ライブラリが必要) |

## 動作環境

- Python 3.8+
- Windows / Mac / Linux
- tkinter (GUI 使用時のみ、Python 標準ライブラリ)

## 既知の制限

- 純粋な **OLE 埋め込みオブジェクト** (`xl/embeddings/*.bin`) は抽出不可 (Python 標準ライブラリでは復号できないため)
- EMF/WMF は **ラスタ/SVG フォールバックが同梱されていない場合** プレースホルダ表示となる
- 列幅/行高はワークブックのデフォルトフォントを Calibri 11 (MDW=7) と仮定して計算
  (`ConvertConfig.char_width` で変更可)

## ファイル構成

```
ExcelToDrawIO/
├── excel_to_drawio.py   # 変換エンジン + CLI
├── desktop_app.py       # GUI アプリ (tkinter)
└── README.md                 # 本ドキュメント
```

## ライセンス

MIT License
