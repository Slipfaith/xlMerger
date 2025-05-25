import pytest
from core.main_page_logic import MainPageLogic
from PySide6.QtWidgets import QWidget

class DummyUI(QWidget):
    def __init__(self):
        super().__init__()
        self.folder_entry = type('', (), {'text': lambda s: "", 'setText': lambda s, v: None})()
        self.excel_file_entry = type('', (), {'text': lambda s: "", 'setText': lambda s, v: None})()
        self.copy_column_entry = type('', (), {'text': lambda s: "", 'setText': lambda s, v: None})()
        self.sheet_list = type('', (), {'count': lambda s: 0, 'item': lambda s, i: None})()

def test_main_page_logic_validation(monkeypatch):
    ui = DummyUI()
    logic = MainPageLogic(ui)

    # Нет файлов и папок — не валидно
    assert not logic.validate_inputs()

    # Пробросим валидные значения
    monkeypatch.setattr(ui.folder_entry, "text", lambda: "/tmp")
    monkeypatch.setattr(ui.excel_file_entry, "text", lambda: "/tmp/file.xlsx")
    monkeypatch.setattr(ui.copy_column_entry, "text", lambda: "A")

    class DummyItem:
        def checkState(self):
            return 2  # Qt.Checked
        def text(self):
            return "Sheet1"

    monkeypatch.setattr(ui.sheet_list, "count", lambda: 1)
    monkeypatch.setattr(ui.sheet_list, "item", lambda s, i: DummyItem())

    assert logic.validate_inputs()
