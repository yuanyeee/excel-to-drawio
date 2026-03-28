# Excel to draw.io Converter

Excelの図形、フロー Chart、ER図をdraw.io形式に変換するデスクトップアプリケーションです。

## 特徴

- **シンプルな操作**: ファイルを選んで、シートを選択して、変換ボタン押すだけ
- **Windows標準UI使用**: tkinterで構築、indowsのファイルダイアログをそのまま使用
- **サイズ調整可能**: ウィンドウのサイズを自由に変更可能
- **スクロール対応**: 多数シートがあってもスクロールして選択可能
- **形状変換**: 矩形、楕円、菱形、矢印、コネクター 등을 支持
- **セル図形対応**: 結合セルや枠線で描いた図形も変換可能
- **スタイル保存**: 塗りつぶし色、線色、テキスト样式를 保存

## 対応ファイル形式

| 形式 | 説明 | 対応 |
|------|------|------|
| .xlsx | Excel 2007以降 | ✅ |
| .xls | Excel 97-2003 | ✅ |
| .xlsm | Excel マクロ付き | ✅ |

## インストール

```powershell
cd excel-to-drawio
pip install -r requirements.txt
```

## 使い方

### デスクトップGUI（推奨）

```powershell
python gui_tkinter.py
```

画面が起動したら：

1. **「Browse Excel File」**をクリックしてExcelファイルを選択
2. **変換したいシートにチェックを入れる**
3. **必要に応じてオプションを設定**:
   - ☑ Include connectors/lines（線を含む）
   - ☑ Include cell background colors（背景色を含む）
4. **「Convert to draw.io」**ボタンをクリック
5. **保存ファイル名（.drawio）と保存先を指定して完了**

### コマンドライン

```powershell
python main.py 入力ファイル.xlsx --sheets "シート1" "シート2"
```

オプション:
- `--sheets`: 変換するシート名を指定（省略すると全シート）
- `-o 出力ファイル`: 出力先を指定

例：
```powershell
python main.py diagram.xlsx --sheets "Sheet1" "Sheet2" -o output.drawio
```

## プロジェクト構成

```
excel-to-drawio/
├── gui_tkinter.py      # デスクトップGUIアプリ
├── main.py             # コマンドライン版
├── converter/          # 変換エンジン
│   ├── __init__.py
│   ├── excel_reader.py     # Excel読み込み
│   ├── shape_mapper.py     # shapeタイプ変換
│   ├── drawio_writer.py    # draw.io XML出力
│   └── cell_border.py     # セル枠線処理
├── tests/              # テスト
├── requirements.txt    # 依存ライブラリ
├── README.md
└── SPEC.md           # 仕様書
```

## 動作環境

- Python 3.8+
- Windows / Mac / Linux
- tkinter（Python標準ライブラリ）

## トラブルシューティング

### Q: tkinterのエラーが出る
A: Pythonが正しくインストールされているか確認してください。公式サイトからPython 3.8以降をインストールしてください。

### Q: Excelファイルが開けない
A: ファイルが別のプログラムで開かれていないか確認してください。閉じてから再度試してください。

### Q: 変換结果がdraw.ioで開けない
A: draw.io官方网站から最新版をダウンロードしてください。

## ライセンス

MIT License
