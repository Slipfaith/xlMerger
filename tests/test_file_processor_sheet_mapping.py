# -*- coding: utf-8 -*-
import sys

import pytest

pytest.importorskip("PySide6.QtWidgets")
from PySide6.QtWidgets import QApplication, QMessageBox

from gui.file_processor_app import FileProcessorApp


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    return app


def test_check_sheet_mapping_skips_dialog_when_auto_map_is_complete(qapp, monkeypatch):
    source = "C:/tmp/source_one_sheet.xlsx"
    monkeypatch.setattr(
        "gui.file_processor_app.ExcelProcessor.get_sheet_names",
        lambda path: ["SourceOnly"],
    )

    widget = FileProcessorApp()
    widget.selected_files = [source]
    widget.selected_sheets = ["TargetA", "TargetB"]

    called = {"value": False}

    class _Dialog:
        def __init__(self, *args, **kwargs):
            called["value"] = True

    monkeypatch.setattr("gui.file_processor_app.SheetMappingDialog", _Dialog)

    assert widget.check_sheet_mapping() is True
    assert called["value"] is False
    assert widget.file_to_sheet_map[source] == {
        "TargetA": "SourceOnly",
        "TargetB": "SourceOnly",
    }
    widget.close()


def test_check_sheet_mapping_rejects_incomplete_mapping_from_dialog(qapp, monkeypatch):
    source = "C:/tmp/source_two_sheets.xlsx"
    monkeypatch.setattr(
        "gui.file_processor_app.ExcelProcessor.get_sheet_names",
        lambda path: ["SourceA", "SourceB"],
    )

    widget = FileProcessorApp()
    widget.selected_files = [source]
    widget.selected_sheets = ["TargetA", "TargetB"]

    class _Dialog:
        def __init__(self, *args, **kwargs):
            return None

        def exec(self):
            return True

        def get_mapping(self):
            # Only one target mapped; second target intentionally missing.
            return {source: {"TargetA": "SourceA"}}

    monkeypatch.setattr("gui.file_processor_app.SheetMappingDialog", _Dialog)
    monkeypatch.setattr(
        "gui.file_processor_app.QMessageBox.warning",
        lambda *args, **kwargs: QMessageBox.Ok,
    )

    assert widget.check_sheet_mapping() is False
    widget.close()
