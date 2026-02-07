# -*- coding: utf-8 -*-
import sys
import ctypes
from pathlib import Path
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon
from gui.main_window import MainWindow
from gui.style_system import apply_app_style

ICON_PATH = Path(__file__).resolve().parent / "xlM2.0.ico"


def _set_windows_app_id():
    if not sys.platform.startswith("win"):
        return
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("xlMerger.desktop.app")
    except Exception:
        pass

if __name__ == "__main__":
    _set_windows_app_id()
    app = QApplication(sys.argv)
    apply_app_style(app)
    icon = QIcon(str(ICON_PATH))
    if not icon.isNull():
        app.setWindowIcon(icon)

    window = MainWindow()
    if icon.isNull():
        icon = window.windowIcon()
    if not icon.isNull():
        app.setWindowIcon(icon)
        window.setWindowIcon(icon)
    window.show()

    sys.exit(app.exec())
