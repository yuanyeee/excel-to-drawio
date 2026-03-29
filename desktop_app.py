#!/usr/bin/env python3
"""
Desktop app launcher for Excel to draw.io.

This module intentionally delegates to gui_tkinter so users can run either:
- `python gui_tkinter.py`
- `python desktop_app.py`
"""

from gui_tkinter import main


if __name__ == "__main__":
    main()
