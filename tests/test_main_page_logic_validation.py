# -*- coding: utf-8 -*-
import sys
import pytest

# Skip the test if PySide6 is not available. The project uses this Qt
# binding for the main UI logic, but the testing environment might not
# have it installed. ``pytest.importorskip`` will mark the test as
# skipped instead of failing during import.
pytest.importorskip("PySide6")

from PySide6.QtCore import Signal
from PySide6.QtWidgets import QApplication, QWidget, QMessageBox
from core.main_page_logic import MainPageLogic

@pytest.fixture(scope="session")
def qapp():
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    return app

class DummyEntry:
    def __init__(self, value=""):
        self._value = value
    def text(self):
        return self._value
    def setText(self, v):
        self._value = v

class DummySheetList:
    def __init__(self, count=0, checked=True, names=None):
        self._count = count
        self._checked = checked
        self._names = names or []
    def count(self):
        return self._count
    def item(self, i):
        class DummyItem:
            def checkState(self):  # 2 = Qt.Checked
                return 2 if self._checked else 0
            def text(self):
                return self._names[i] if self._names else "Sheet1"
        item = DummyItem()
        item._checked = self._checked
        item._names = self._names
        return item
    def clear(self):
        self._count = 0
        self._names = []

class DummyUI(QWidget):
    folderSelected = Signal(str)
    filesSelected = Signal(list)
    excelFileSelected = Signal(str)
    processTriggered = Signal()
    previewTriggered = Signal()

    def __init__(self):
        super().__init__()
        self.folder_entry = DummyEntry()
        self.excel_file_entry = DummyEntry()
        self.copy_column_entry = DummyEntry()
        self.sheet_list = DummySheetList()

def test_main_page_logic_validation(qapp, monkeypatch, tmp_path):
    monkeypatch.setattr(QMessageBox, "warning", lambda *a, **k: QMessageBox.Ok)
    monkeypatch.setattr(QMessageBox, "critical", lambda *a, **k: QMessageBox.Ok)

    ui = DummyUI()
    logic = MainPageLogic(ui)

    # Нет файлов/папок — не валидно
    assert not logic.validate_inputs()

    # Валидные значения: папка существует, файл существует и лист с нужным именем
    folder = tmp_path
    file_path = tmp_path / "file.xlsx"
    # Для проверки нам достаточно существования файла. Создаём пустой
    # файл вместо использования ``openpyxl``.
    with open(file_path, "wb"):
        pass

    ui.folder_entry.setText(str(folder))
    ui.excel_file_entry.setText(str(file_path))
    ui.copy_column_entry.setText("A")
    ui.sheet_list = DummySheetList(count=1, checked=True, names=["Sheet1"])

    assert logic.validate_inputs()
