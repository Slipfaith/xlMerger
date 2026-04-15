# -*- coding: utf-8 -*-
import pytest
from openpyxl import Workbook

QtWidgets = pytest.importorskip("PySide6.QtWidgets")
from PySide6.QtWidgets import QApplication
from gui.split_mapping_dialog import SplitMappingDialog


def test_split_mapping_dialog_handles_blank_header(tmp_path):
    app = QApplication.instance() or QApplication([])

    src = tmp_path / "main.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append([None, None])
    ws.append(["hi", "hello"])
    wb.save(src)
    wb.close()

    dialog = SplitMappingDialog(str(src), ["Sheet1"])
    dialog.source_col = 0
    dialog.target_cols = {1}

    selection = dialog.get_selection()
    assert selection == {"Sheet1": ("A", ["B"], [])}

    dialog.close()


def test_split_mapping_dialog_shows_click_hint(tmp_path):
    app = QApplication.instance() or QApplication([])

    src = tmp_path / "main.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["A", "B"])
    ws.append(["v1", "v2"])
    wb.save(src)
    wb.close()

    dialog = SplitMappingDialog(str(src), ["Sheet1"])
    hint_text = dialog.info_label.text()
    assert "\u041b\u0435\u0432\u043e\u0439 \u043a\u043d\u043e\u043f\u043a\u043e\u0439 \u043c\u044b\u0448\u0438" in hint_text
    assert "\u043f\u0440\u0430\u0432\u043e\u0439 \u043a\u043d\u043e\u043f\u043a\u043e\u0439 \u043c\u044b\u0448\u0438" in hint_text

    dialog.close()
