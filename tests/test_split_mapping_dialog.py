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
