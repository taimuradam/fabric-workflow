"""
main.py — Launches the Textile Lot Costing app in a pywebview window.

The UI is an HTML/CSS/JS frontend (frontend/index.html) rendered in a native
webview; all Python logic is reached through api.Api (see api.py). This file just
creates the window, wires up the js_api, and starts the event loop.
"""

import os
import sys

import webview

from api import Api

# Window defaults — generous size for a non-technical user at arm's length.
_WINDOW_TITLE = "Textile Lot Costing"
_WINDOW_WIDTH = 1200
_WINDOW_HEIGHT = 900
_MIN_SIZE = (960, 720)


def resource_path(relative: str) -> str:
    """Absolute path to a bundled resource, working both as script and as a
    PyInstaller one-file exe.

    PyInstaller unpacks --add-data files into a temp dir exposed as sys._MEIPASS;
    when running from source there's no _MEIPASS, so we fall back to this file's
    own directory.
    """
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


def main():
    api = Api()

    window = webview.create_window(
        _WINDOW_TITLE,
        url=resource_path(os.path.join("frontend", "index.html")),
        js_api=api,
        width=_WINDOW_WIDTH,
        height=_WINDOW_HEIGHT,
        min_size=_MIN_SIZE,
        resizable=True,
    )
    # api needs the window handle to open native file dialogs.
    api.window = window

    # On Windows use the modern Edge WebView2 (Chromium) backend explicitly; on
    # other platforms let pywebview pick its native default (e.g. Cocoa on macOS)
    # so the app is still runnable for development on a Mac.
    gui = "edgechromium" if sys.platform == "win32" else None
    webview.start(gui=gui)


if __name__ == "__main__":
    main()
