#!/bin/bash
# Excel to Draw.io Converter - GUI launcher (macOS / Linux)
set -e
cd "$(dirname "$0")"

PYCMD=""
if command -v python3 >/dev/null 2>&1; then
    PYCMD=python3
elif command -v python >/dev/null 2>&1; then
    PYCMD=python
else
    echo "[ERROR] Python が見つかりません。"
    echo "        macOS: https://www.python.org/downloads/ からインストール、または 'brew install python'"
    echo "        Linux: 'sudo apt install python3' (Debian/Ubuntu系) などお使いのディストリのパッケージマネージャーで"
    read -r -p "Press Enter to exit..."
    exit 1
fi

if ! "$PYCMD" -c "import tkinter" >/dev/null 2>&1; then
    echo "[ERROR] tkinter が見つかりません。"
    echo "        macOS: 'brew install python-tk' を実行してください"
    echo "        Linux : 'sudo apt install python3-tk' (Debian/Ubuntu系) などで tkinter を追加してください"
    read -r -p "Press Enter to exit..."
    exit 1
fi

echo "Excel to Draw.io Converter を起動しています..."
set +e
"$PYCMD" -m excel_to_drawio.desktop_app
STATUS=$?
set -e
if [ "$STATUS" -ne 0 ]; then
    echo
    echo "[ERROR] GUI の起動に失敗しました。上記のエラーを確認してください。"
    read -r -p "Press Enter to exit..."
fi
