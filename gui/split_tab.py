from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox, QMessageBox
)
from utils.i18n import tr, i18n
from core.drag_drop import DragDropLineEdit
from core.split_excel import split_excel_by_languages
from openpyxl import load_workbook


class SplitTab(QWidget):
    def __init__(self):
        super().__init__()
        self.excel_path = ''
        self.sheet_name = ''
        self.headers = []
        self.init_ui()
        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        self.file_input = DragDropLineEdit(mode='file')
        self.file_input.fileSelected.connect(self.on_file_selected)
        layout.addWidget(QLabel(tr("Файл Excel:")))
        layout.addWidget(self.file_input)

        self.sheet_combo = QComboBox()
        self.sheet_combo.currentTextChanged.connect(self.on_sheet_changed)
        layout.addWidget(self.sheet_combo)

        layout.addWidget(QLabel(tr("Исходный язык:")))
        self.source_combo = QComboBox()
        layout.addWidget(self.source_combo)

        self.split_btn = QPushButton(tr("Разделить"))
        self.split_btn.clicked.connect(self.run_split)
        layout.addWidget(self.split_btn)

    def on_file_selected(self, path):
        try:
            wb = load_workbook(path, read_only=True)
            self.excel_path = path
            self.sheet_combo.clear()
            self.sheet_combo.addItems(wb.sheetnames)
            self.sheet_name = wb.sheetnames[0] if wb.sheetnames else ''
            wb.close()
            self.load_headers()
        except Exception as e:
            QMessageBox.critical(self, tr("Ошибка"), str(e))

    def on_sheet_changed(self, name):
        self.sheet_name = name
        self.load_headers()

    def load_headers(self):
        if not self.excel_path or not self.sheet_name:
            return
        wb = load_workbook(self.excel_path, read_only=True)
        sheet = wb[self.sheet_name]
        self.headers = [
            str(cell.value) if cell.value is not None else ''
            for cell in next(sheet.iter_rows(min_row=1, max_row=1))
        ]
        wb.close()
        self.source_combo.clear()
        self.source_combo.addItems([h for h in self.headers if h])

    def run_split(self):
        if not self.excel_path:
            QMessageBox.critical(self, tr("Ошибка"), tr("Выберите файл Excel."))
            return
        src = self.source_combo.currentText()
        try:
            split_excel_by_languages(self.excel_path, self.sheet_name, src)
            QMessageBox.information(self, tr("Успех"), tr("Файлы успешно сохранены."))
        except Exception as e:
            QMessageBox.critical(self, tr("Ошибка"), str(e))

    def retranslate_ui(self):
        self.setWindowTitle(tr("Разделение"))
        self.split_btn.setText(tr("Разделить"))
        # update labels - they are static but to refresh we need to re-add them? Not necessary
