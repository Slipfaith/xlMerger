# -*- coding: utf-8 -*-
import pytest

QtWidgets = pytest.importorskip("PySide6.QtWidgets")
from PySide6.QtWidgets import QApplication

from gui.sheet_mapping_dialog import SheetMappingDialog


def test_sheet_mapping_dialog_returns_target_to_source_mapping():
    app = QApplication.instance() or QApplication([])
    file_path = "C:/tmp/source.xlsx"
    dialog = SheetMappingDialog(
        main_sheets=["TargetA", "TargetB"],
        file_to_sheets={file_path: ["Source1", "Source2"]},
        auto_map={file_path: {"TargetA": "Source1", "TargetB": "Source2"}},
    )

    dialog.comboboxes[(file_path, "Source1")].setCurrentText("TargetA")
    dialog.comboboxes[(file_path, "Source2")].setCurrentText("TargetB")
    mapping = dialog.get_mapping()

    assert mapping[file_path]["TargetA"] == "Source1"
    assert mapping[file_path]["TargetB"] == "Source2"
    dialog.close()


def test_sheet_mapping_dialog_fallback_uses_first_source_when_no_selection():
    app = QApplication.instance() or QApplication([])
    file_path = "C:/tmp/source.xlsx"
    dialog = SheetMappingDialog(
        main_sheets=["TargetA", "TargetB"],
        file_to_sheets={file_path: ["OnlySource"]},
        auto_map={},
    )

    dialog.comboboxes[(file_path, "OnlySource")].setCurrentText("")
    mapping = dialog.get_mapping()

    assert mapping[file_path]["TargetA"] == "OnlySource"
    assert mapping[file_path]["TargetB"] == "OnlySource"
    dialog.close()
