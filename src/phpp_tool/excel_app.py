"""Excel application factory for xlwings.

Finds the newest Excel installation on macOS and launches it hidden.
On Windows, xlwings finds Excel automatically via COM — no path needed.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

import xlwings as xw

_MACOS_EXCEL_PATHS = [
    "/Applications/Microsoft Office 2021/Microsoft Excel.app",
    "/Applications/Microsoft Office 2019/Microsoft Excel.app",
    "/Applications/Microsoft Office 2016/Microsoft Excel.app",
    "/Applications/Microsoft Excel.app",
]


def _find_excel() -> str | None:
    """Return the path to the newest Excel installation, or None."""
    if sys.platform != "darwin":
        return None
    override = os.environ.get("PHPP_EXCEL_PATH")
    if override and Path(override).exists():
        return override
    for p in _MACOS_EXCEL_PATHS:
        if Path(p).exists():
            return p
    return None


_shared_app: xw.App | None = None


def set_shared_app(app: xw.App | None) -> None:
    """Set a shared Excel app instance for reuse (e.g. during test sessions)."""
    global _shared_app
    _shared_app = app


def excel_app() -> xw.App:
    """Return the shared Excel instance if set, otherwise launch a new one."""
    if _shared_app is not None:
        return _shared_app
    spec = _find_excel()
    app = xw.App(spec=spec, visible=False, add_book=False)
    app.display_alerts = False
    return app


def is_shared(app: xw.App) -> bool:
    """Return True if *app* is the shared instance (caller should not quit it)."""
    return _shared_app is not None and app is _shared_app


def open_book(app: xw.App, path: str, **kwargs: object) -> xw.Book:
    """Open a workbook, retrying once on macOS error -50.

    On macOS 26, rapid open/close cycles on a shared Excel instance can
    produce a transient AppleScript parameter error (-50) if the previous
    close hasn't fully settled.
    """
    try:
        return app.books.open(path, **kwargs)
    except Exception:
        import time
        time.sleep(1)
        return app.books.open(path, **kwargs)
