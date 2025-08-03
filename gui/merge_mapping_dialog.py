# merge_mapping_dialog.py
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QWidget, QFileDialog,
    QLabel, QComboBox, QListWidget, QListWidgetItem, QMessageBox, QSizePolicy, QAbstractItemView
)
from PySide6.QtCore import Qt
import os

try:
    import openpyxl
except ImportError:
    openpyxl = None
try:
    import xlrd
except ImportError:
    xlrd = None

SUPPORTED_EXT = (".xlsx", ".xls")
MAX_PREVIEW_ROWS = 5  # для предпросмотра, если нужно

def get_excel_structure(path):
    sheets = {}
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xlsx" and openpyxl:
        wb = openpyxl.load_workbook(path, read_only=True)
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            headers = []
            for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                headers = [str(val) if val is not None else "" for val in row]
            sheets[sheet] = headers
        wb.close()
    elif ext == ".xls" and xlrd:
        wb = xlrd.open_workbook(path)
        for sheet in wb.sheet_names():
            ws = wb.sheet_by_name(sheet)
            headers = [str(ws.cell_value(0, col)) for col in range(ws.ncols)]
            sheets[sheet] = headers
    else:
        raise Exception("Unsupported file format or required lib missing")
    return sheets

class MappingRow(QWidget):
    def __init__(self, main_structure, parent=None):
        super().__init__(parent)
        self.main_structure = main_structure
        self.file_path = None
        self.file_structure = {}

        layout = QHBoxLayout(self)
        layout.setSpacing(8)

        self.file_btn = QPushButton("Источник...")
        self.file_btn.clicked.connect(self.select_file)
        layout.addWidget(self.file_btn)

        self.sheet_box = QComboBox()
        self.sheet_box.setEnabled(False)
        self.sheet_box.currentIndexChanged.connect(self.on_sheet_change)
        layout.addWidget(self.sheet_box)

        self.cols_list = QListWidget()
        self.cols_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.cols_list.setEnabled(False)
        self.cols_list.setMaximumHeight(60)
        self.cols_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(self.cols_list)

        self.target_sheet_box = QComboBox()
        self.target_sheet_box.addItems(main_structure.keys())
        self.target_sheet_box.setEnabled(True)
        self.target_sheet_box.currentIndexChanged.connect(self.on_target_sheet_change)
        layout.addWidget(self.target_sheet_box)

        self.target_cols_list = QListWidget()
        self.target_cols_list.setSelectionMode(QAbstractItemView.MultiSelection)
        self.target_cols_list.setEnabled(False)
        self.target_cols_list.setMaximumHeight(60)
        self.target_cols_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        layout.addWidget(self.target_cols_list)

        self.remove_btn = QPushButton("✕")
        layout.addWidget(self.remove_btn)

        self.sheet_box.setMinimumWidth(90)
        self.cols_list.setMinimumWidth(110)
        self.target_sheet_box.setMinimumWidth(90)
        self.target_cols_list.setMinimumWidth(110)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Выбери Excel", "", "Excel files (*.xlsx *.xls)")
        if not file_path:
            return
        try:
            structure = get_excel_structure(file_path)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка чтения файла:\n{e}")
            return
        self.file_path = file_path
        self.file_structure = structure
        self.file_btn.setText(os.path.basename(file_path))
        self.sheet_box.clear()
        self.sheet_box.addItems(list(structure.keys()))
        self.sheet_box.setEnabled(True)
        self.on_sheet_change()

    def on_sheet_change(self):
        sheet = self.sheet_box.currentText()
        headers = self.file_structure.get(sheet, [])
        self.cols_list.clear()
        for h in headers:
            item = QListWidgetItem(h)
            item.setCheckState(Qt.Unchecked)
            self.cols_list.addItem(item)
        self.cols_list.setEnabled(bool(headers))

    def on_target_sheet_change(self):
        sheet = self.target_sheet_box.currentText()
        headers = self.main_structure.get(sheet, [])
        self.target_cols_list.clear()
        for h in headers:
            item = QListWidgetItem(h)
            item.setCheckState(Qt.Unchecked)
            self.target_cols_list.addItem(item)
        self.target_cols_list.setEnabled(bool(headers))

    def get_mapping(self):
        # Валидация обязательных полей
        if not self.file_path or not self.sheet_box.currentText():
            return None
        src_cols = [self.cols_list.item(i).text()
                    for i in range(self.cols_list.count())
                    if self.cols_list.item(i).checkState() == Qt.Checked]
        tgt_cols = [self.target_cols_list.item(i).text()
                    for i in range(self.target_cols_list.count())
                    if self.target_cols_list.item(i).checkState() == Qt.Checked]
        return {
            "source_file": self.file_path,
            "source_sheet": self.sheet_box.currentText(),
            "source_columns": src_cols,
            "target_sheet": self.target_sheet_box.currentText(),
            "target_columns": tgt_cols
        }

    def auto_map_columns(self):
        # Сопоставить колонки с совпадающими именами
        src_headers = self.file_structure.get(self.sheet_box.currentText(), [])
        tgt_headers = self.main_structure.get(self.target_sheet_box.currentText(), [])
        min_len = min(len(src_headers), len(tgt_headers))
        # Сопоставлять по одинаковым именам
        for i in range(self.cols_list.count()):
            item = self.cols_list.item(i)
            if item.text() in tgt_headers:
                item.setCheckState(Qt.Checked)
        for i in range(self.target_cols_list.count()):
            item = self.target_cols_list.item(i)
            if item.text() in src_headers:
                item.setCheckState(Qt.Checked)

class MergeMappingDialog(QDialog):
    def __init__(self, main_excel_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Настройки объединения")
        self.setMinimumWidth(1000)
        layout = QVBoxLayout(self)

        # Главный Excel: строим структуру один раз
        try:
            self.main_structure = get_excel_structure(main_excel_path)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка чтения главного файла:\n{e}")
            self.main_structure = {}

        self.rows = []
        self.rows_layout = QVBoxLayout()
        layout.addLayout(self.rows_layout)

        btns_row = QHBoxLayout()
        self.add_btn = QPushButton("Добавить источник")
        self.add_btn.clicked.connect(self.add_row)
        btns_row.addWidget(self.add_btn)

        self.auto_map_btn = QPushButton("Автосопоставить")
        self.auto_map_btn.clicked.connect(self.auto_map_all)
        btns_row.addWidget(self.auto_map_btn)

        btns_row.addStretch()
        layout.addLayout(btns_row)

        ok_row = QHBoxLayout()
        ok_row.addStretch()
        self.ok_btn = QPushButton("Готово")
        self.ok_btn.clicked.connect(self.accept)
        ok_row.addWidget(self.ok_btn)
        self.cancel_btn = QPushButton("Отмена")
        self.cancel_btn.clicked.connect(self.reject)
        ok_row.addWidget(self.cancel_btn)
        layout.addLayout(ok_row)

    def add_row(self):
        row = MappingRow(self.main_structure, self)
        row.remove_btn.clicked.connect(lambda: self.remove_row(row))
        self.rows.append(row)
        self.rows_layout.addWidget(row)

    def remove_row(self, row):
        self.rows.remove(row)
        row.setParent(None)
        row.deleteLater()

    def auto_map_all(self):
        for row in self.rows:
            row.auto_map_columns()

    def get_mappings(self):
        result = []
        for row in self.rows:
            mapping = row.get_mapping()
            if mapping:
                result.append(mapping)
        return result

