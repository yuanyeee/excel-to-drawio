@echo off
echo ========================================
echo Excel to draw.io Converter
echo ========================================
echo.
echo Starting server...
echo Open http://localhost:8765 in your browser
echo Press Ctrl+C to stop
echo ========================================
echo.

REM Find Python
where python >nul 2>&1
if %errorlevel%==0 (
    python serve.py 8765
) else (
    where python3 >nul 2>&1
    if %errorlevel%==0 (
        python3 serve.py 8765
    ) else (
        echo Python not found. Please install Python 3.8+ from python.org
        pause
    )
)
