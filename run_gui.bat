@echo off
chcp 65001 >nul
rem Excel to Draw.io Converter - GUI launcher (Windows)
setlocal
cd /d "%~dp0"

set PYCMD=
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set PYCMD=python
) else (
    where py >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        set PYCMD=py
    )
)

if "%PYCMD%"=="" (
    echo [ERROR] Python が見つかりません。
    echo         https://www.python.org/downloads/ からインストールしてください。
    echo         インストール時は "Add python.exe to PATH" にチェックを入れてください。
    pause
    exit /b 1
)

%PYCMD% -c "import tkinter" >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] tkinter が見つかりません。
    echo         Python インストーラーを実行し、"tcl/tk and IDLE" オプションを
    echo         有効にして修復インストールしてください。
    pause
    exit /b 1
)

echo Excel to Draw.io Converter を起動しています...
%PYCMD% -m excel_to_drawio.desktop_app
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] GUI の起動に失敗しました。上記のエラーを確認してください。
    pause
)

endlocal
