# -*- coding: utf-8 -*-
import sys
import pytest

pytest.importorskip("PySide6.QtWidgets")
from PySide6.QtWidgets import QApplication

from gui.merge_mapping_dialog import MappingCard


@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    return app


def test_header_mapping_mode_returns_column_letters(qapp):
    card = MappingCard(main_structure={"Main": ["RU", "KA"]})
    card.file_path = "source.xlsx"
    card.file_structure = {"Main": ["RU", "KA"]}
    card.source_sheet_combo.addItem("Main")
    card.source_sheet_combo.setCurrentText("Main")
    card.target_sheet_combo.clear()
    card.target_sheet_combo.addItem("Main")
    card.target_sheet_combo.setCurrentText("Main")

    card.header_mode.setChecked(True)
    card.create_mapping_interface()

    source_combo, target_combo, _ = card.header_mappings[0]
    source_combo.setCurrentIndex(1)  # A: RU
    target_combo.setCurrentIndex(2)  # B: KA

    mapping = card.get_mapping()
    assert mapping["source_columns"] == ["A"]
    assert mapping["target_columns"] == ["B"]
