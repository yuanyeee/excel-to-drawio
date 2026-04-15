# Excel to Draw.io Desktop Tool 実行説明

[🇯🇵 日本語](desktop_app_usage.md) | [🇬🇧 English](desktop_app_usage_en.md) | [🇨🇳 简体中文](desktop_app_usage_zh.md)

## 概要

このツールは、Excel ファイルの指定シートを Draw.io 形式へ変換する Python デスクトップアプリです。

- 対応形式: `.xlsx`, `.xlsm`
- UI: `tkinter`
- 操作: ファイル選択 → シート選択 → オプション選択 → 変換

## 実行前提

- Windows
- Python 3 がインストール済み
- このフォルダ内に以下のファイルが存在すること
  - `desktop_app.py`
  - `excel_to_drawio.py`

## 起動方法

作業フォルダで以下を実行します。

```powershell
python .\desktop_app.py
```

## 使い方

1. `Browse...` を押して、変換したい Excel ファイルを選択します。
2. 読み込まれたシート一覧から、変換対象のシートを選択します。
3. `Output` に出力先の `.drawio` ファイル名が自動で入ります。
4. 必要なら `Save As...` で保存先を変更します。
5. `Convert` を押します。
6. 正常終了すると、保存先に `.drawio` ファイルが作成されます。

## 主なオプション

- Include images: 画像の埋め込み有効/無効
- Include borders: 罫線描画の有効/無効
- Merge same-color fills: 同色塗りセルの結合描画
- Skip hidden rows/cols: 非表示行/列を除外

## 画面項目

- `Excel File`
  - 入力ファイルのパス
- `Sheets`
  - ブック内シート一覧
  - シート数が多い場合はスクロール可能
- `Output`
  - 出力する `.drawio` ファイルパス
- `Convert`
  - 変換実行
- 下部ログ欄
  - 読み込み結果や変換結果、エラー内容を表示

## 注意事項

- `.xls` はこの版では未対応です。
- 変換は同期実行です。大きいファイルでは完了まで少し待つ場合があります。
- シートにより、Excel 固有の複雑な図形表現は Draw.io 側で完全一致しないことがあります。
- 入力ファイルや出力先にアクセス権限がない場合、エラーダイアログが表示されます。

## よくある操作

### 出力先を自動決定したい

シートを選択すると、同じフォルダに `シート名.drawio` という名前で自動設定されます。

### 別の名前で保存したい

`Save As...` で任意の保存先とファイル名を指定してください。

### エラーが出たとき

画面下のログ欄とポップアップメッセージを確認してください。

