import os
import sys
import traceback
import pandas as pd
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QComboBox, QTableView, QMessageBox, QHeaderView
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QStandardItemModel, QStandardItem
from openpyxl import load_workbook
import openpyxl.utils as utils

class ExcelPreviewer(QWidget):
    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path
        self.workbook = load_workbook(excel_path, read_only=True)
        self.sheet_names = self.workbook.sheetnames
        self.current_sheet = self.sheet_names[0]
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(self.tr("Просмотр: ") + self.current_sheet)
        layout = QVBoxLayout()

        self.init_sheet_selector()
        layout.addWidget(self.sheet_selector)

        self.table_view = QTableView(self)
        layout.addWidget(self.table_view)

        self.load_excel_data()

        self.setLayout(layout)
        self.setGeometry(100, 100, 800, 600)
        self.show()

    def init_sheet_selector(self):
        self.sheet_selector = QComboBox(self)
        self.sheet_selector.addItems(self.sheet_names)
        self.sheet_selector.currentTextChanged.connect(self.switch_sheet)

    def switch_sheet(self, sheet_name):
        self.current_sheet = sheet_name
        self.load_excel_data()

    def load_excel_data(self):
        try:
            df = self.read_excel_data()
            self.populate_table_view(df)
        except Exception as e:
            self.handle_load_error(e)

    def read_excel_data(self):
        df = pd.read_excel(self.excel_path, sheet_name=self.current_sheet, header=None, dtype=str)
        df.fillna("", inplace=True)
        return df

    def populate_table_view(self, df):
        model = self.create_table_model(df)
        self.table_view.setModel(model)
        self.adjust_table_headers(df)

    def create_table_model(self, df):
        model = QStandardItemModel(df.shape[0], df.shape[1])
        # Буквенные заголовки столбцов (A, B, C, ...)
        model.setHorizontalHeaderLabels([utils.get_column_letter(col + 1) for col in range(df.shape[1])])
        row_labels = [str(i + 1) for i in range(df.shape[0])]
        model.setVerticalHeaderLabels(row_labels)

        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                item_value = df.iat[row, col]
                # Можно убрать или оставить print для отладки
                # print(f"Строка: {row+1}, Столбец: {col+1}, Длина исходного текста: {len(item_value)}, Текст: {item_value}")
                item = QStandardItem(item_value)
                model.setItem(row, col, item)
        return model

    def adjust_table_headers(self, df):
        vertical_header = self.table_view.verticalHeader()
        vertical_header.setVisible(True)
        vertical_header.setMinimumWidth(40)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_view.verticalHeader().setDefaultSectionSize(20)
        self.table_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        max_width_px = 113

        for col in range(df.shape[1]):
            self.table_view.resizeColumnToContents(col)
            current_width = self.table_view.columnWidth(col)
            new_width = min(current_width, max_width_px)
            self.table_view.setColumnWidth(col, new_width)

    def handle_load_error(self, error):
        QMessageBox.critical(self, self.tr("Ошибка"), self.tr("Не удалось загрузить данные с листа: ") + str(error))
        self.log_error(error)

    def log_error(self, error):
        log_file_path = os.path.join(os.path.dirname(__file__), 'error_log.txt')
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write(self.tr("Ошибка: ") + str(error) + "\n")
            log_file.write(traceback.format_exc())
            log_file.write("\n")
