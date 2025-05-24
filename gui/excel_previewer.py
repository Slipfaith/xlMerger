import os
import sys
import traceback
import pandas as pd
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QComboBox, QTableView, QMessageBox, QHeaderView
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QStandardItemModel, QStandardItem
from openpyxl import load_workbook
import openpyxl.utils as utils
from openpyxl.styles import PatternFill

class ExcelPreviewer(QWidget):
    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path
        self.workbook = load_workbook(excel_path, read_only=True)
        self.sheet_names = self.workbook.sheetnames
        self.current_sheet = self.sheet_names[0]
        self.init_ui()

    def init_ui(self):
        # Заголовок окна: "Просмотр: <текущий лист>"
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
        """Инициализация комбинированного списка для выбора листа."""
        self.sheet_selector = QComboBox(self)
        self.sheet_selector.addItems(self.sheet_names)
        self.sheet_selector.currentTextChanged.connect(self.switch_sheet)

    def switch_sheet(self, sheet_name):
        """Переключение на другой лист."""
        self.current_sheet = sheet_name
        self.load_excel_data()

    def load_excel_data(self):
        """Загрузка и отображение данных с текущего листа Excel."""
        try:
            df = self.read_excel_data()
            self.populate_table_view(df)
        except Exception as e:
            self.handle_load_error(e)

    def read_excel_data(self):
        """Чтение данных из Excel в DataFrame."""
        df = pd.read_excel(self.excel_path, sheet_name=self.current_sheet, header=None, dtype=str)
        df.fillna("", inplace=True)
        return df

    def populate_table_view(self, df):
        """Заполнение QTableView данными."""
        model = self.create_table_model(df)
        self.table_view.setModel(model)
        self.adjust_table_headers(df)

    def create_table_model(self, df):
        """Создание QStandardItemModel из DataFrame с подробным логированием."""
        model = QStandardItemModel(df.shape[0], df.shape[1])
        # Установка буквенных заголовков столбцов (A, B, C, ...)
        model.setHorizontalHeaderLabels([utils.get_column_letter(col + 1) for col in range(df.shape[1])])
        row_labels = [str(i + 1) for i in range(df.shape[0])]
        model.setVerticalHeaderLabels(row_labels)

        max_width = None  # Убрано ограничение на длину текста для отладки

        for row in range(df.shape[0]):
            for col in range(df.shape[1]):
                item_value = df.iat[row, col]
                log_message = self.tr("Строка: {row}, Столбец: {col}, Длина исходного текста: {length}, Текст: {text}").format(
                    row=row+1, col=col+1, length=len(item_value), text=item_value)
                print(log_message)  # Логирование длины текста и его содержимого
                item = QStandardItem(item_value[:max_width] if max_width else item_value)
                model.setItem(row, col, item)
        return model

    def adjust_table_headers(self, df):
        """Настройка внешнего вида заголовков таблицы с ограничением ширины столбцов до 3 см."""
        vertical_header = self.table_view.verticalHeader()
        vertical_header.setVisible(True)
        vertical_header.setMinimumWidth(40)

        # Переводим режим в ручное управление шириной
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table_view.verticalHeader().setDefaultSectionSize(20)
        self.table_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        # 3 см примерно равны 113 пикселям (при среднем разрешении)
        max_width_px = 113

        for col in range(df.shape[1]):
            # Автоматически рассчитываем ширину по содержимому
            self.table_view.resizeColumnToContents(col)
            current_width = self.table_view.columnWidth(col)
            # Если рассчитанная ширина больше максимальной, ограничиваем её
            new_width = min(current_width, max_width_px)
            self.table_view.setColumnWidth(col, new_width)

    def handle_load_error(self, error):
        """Обработка ошибок во время загрузки данных."""
        QMessageBox.critical(self, self.tr("Ошибка"), self.tr("Не удалось загрузить данные с листа: ") + str(error))
        self.log_error(error)

    def log_error(self, error):
        """Запись деталей ошибки в файл."""
        log_file_path = os.path.join(os.path.dirname(__file__), 'error_log.txt')
        with open(log_file_path, 'a') as log_file:
            log_file.write(self.tr("Ошибка: ") + str(error) + "\n")
            log_file.write(traceback.format_exc())
            log_file.write("\n")
