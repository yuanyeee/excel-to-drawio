@echo off
setlocal

REM Run Plus converter with uv.
REM Usage examples:
REM   run_uv.bat input.xlsx
REM   run_uv.bat input.xlsx -o output.drawio
REM   run_uv.bat -l input.xlsx

uv run excel_to_drawio_plus.py %*

endlocal
