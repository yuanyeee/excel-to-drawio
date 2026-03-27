#!/bin/bash
echo "========================================"
echo "Excel to draw.io Converter"
echo "========================================"
echo ""
echo "Starting server..."
echo "Open http://localhost:8765 in your browser"
echo "Press Ctrl+C to stop"
echo "========================================"
echo ""

# Find Python
if command -v python3 &> /dev/null; then
    python3 serve.py 8765
elif command -v python &> /dev/null; then
    python serve.py 8765
else
    echo "Python not found. Please install Python 3.8+"
    exit 1
fi
