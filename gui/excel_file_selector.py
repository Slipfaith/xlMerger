# -*- coding: utf-8 -*-
import os
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QListWidget, QListWidgetItem,
    QMessageBox, QListView
)
from PySide6.QtCore import Qt

class ExcelFileSelector(QDialog):
    def __init__(self, folder_path, selected_files=None, target_excel=None):
        super().__init__()
        self.folder_path = folder_path
        self.selected_files = selected_files or []
        self.target_excel = target_excel
        self.selected_file = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(self.tr("Превью Excel"))
        layout = QVBoxLayout()

        layout.addWidget(QLabel(self.tr("Загруженные Excel-файлы:")))
        self.file_list = self.create_file_list_widget()
        layout.addWidget(self.file_list)
        layout.addWidget(QLabel(self.tr("Дважды щелкните, чтобы выбрать файл Excel:")))

        self.load_excel_files()

        self.setLayout(layout)
        self.setGeometry(300, 300, 400, 300)
        self.setModal(True)

    def create_file_list_widget(self):
        """Создание QListWidget для отображения файлов."""
        file_list = QListWidget(self)
        file_list.setSelectionMode(QListView.SingleSelection)
        file_list.itemDoubleClicked.connect(self.select_file)
        return file_list

    def load_excel_files(self):
        """Загрузка и отображение файлов Excel из папки."""
        excel_files = self.get_excel_files()
        target_path = self._get_target_path()
        if not excel_files and not target_path:
            self.show_warning_no_files()
            self.close()
            return
        if target_path:
            self.add_file_to_list(
                target_path,
                display_name=self._format_target_name(target_path),
            )
        for file_path in excel_files:
            if target_path and self._normalize_path(file_path) == self._normalize_path(target_path):
                continue
            self.add_file_to_list(file_path)

    def get_excel_files(self):
        """Возвращает список файлов Excel в папке."""
        if self.selected_files:
            files = [
                f for f in self.selected_files
                if os.path.isfile(f) and f.lower().endswith(('.xlsx', '.xls'))
            ]
        elif self.folder_path and os.path.isfile(self.folder_path):
            files = [self.folder_path] if self.folder_path.lower().endswith(('.xlsx', '.xls')) else []
        else:
            files = [
                os.path.join(root, file)
                for root, _, files in os.walk(self.folder_path)
                for file in files
                if file.lower().endswith(('.xlsx', '.xls'))
            ]
        return sorted(files)

    def add_file_to_list(self, file_path, display_name=None):
        """Добавление файла в QListWidget."""
        file_name = display_name or os.path.basename(file_path)
        item = QListWidgetItem(file_name)
        item.setData(Qt.UserRole, file_path)
        item.setToolTip(file_path)
        self.file_list.addItem(item)

    def show_warning_no_files(self):
        """Показ предупреждения, если файлы Excel не найдены."""
        QMessageBox.warning(self, self.tr("Предупреждение"), self.tr("В указанной папке не найдено файлов Excel."))

    def select_file(self, item):
        """Выбор файла из списка."""
        self.selected_file = item.data(Qt.UserRole)
        self.accept()

    def _get_target_path(self):
        if self.target_excel and os.path.isfile(self.target_excel):
            return self.target_excel
        return None

    def _format_target_name(self, path):
        return f"{os.path.basename(path)} ({self.tr('целевой')})"

    def _normalize_path(self, path):
        return os.path.normcase(os.path.abspath(path))
