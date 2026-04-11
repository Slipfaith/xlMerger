# -*- coding: utf-8 -*-
from pathlib import Path
import sys

import pytest

pytest.importorskip("PySide6.QtWidgets")
from PySide6.QtWidgets import QApplication, QWidget

import gui.main_window as main_window_module


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    return app


class _DummyTab(QWidget):
    def __init__(self, *args, **kwargs):
        super().__init__()


def test_main_window_has_no_xlcombine_tab(qapp, monkeypatch):
    monkeypatch.setattr(main_window_module, "FileProcessorApp", _DummyTab)
    monkeypatch.setattr(main_window_module, "LimitsChecker", _DummyTab)
    monkeypatch.setattr(main_window_module, "SplitTab", _DummyTab)
    monkeypatch.setattr(main_window_module, "ExcelBuilderTab", _DummyTab)
    monkeypatch.setattr(main_window_module.MainWindow, "show", lambda self: None)

    window = main_window_module.MainWindow()

    tab_titles = [window.tab_widget.tabText(i) for i in range(window.tab_widget.count())]
    assert "xlCombine" not in tab_titles
    assert window.tab_widget.count() == 3
    window.close()


def test_file_processor_app_has_no_hardcoded_absolute_icon_path():
    source = Path("gui/file_processor_app.py").read_text(encoding="utf-8")
    assert "C:\\Users\\" not in source
