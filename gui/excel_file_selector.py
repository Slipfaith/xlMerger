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

        layout.addWidget(self.create_target_label())
        layout.addWidget(QLabel(self.tr("Загруженные Excel-файлы:")))
        self.file_list = self.create_file_list_widget()
        layout.addWidget(self.file_list)
        layout.addWidget(QLabel(self.tr("Дважды щелкните, чтобы выбрать файл Excel:")))

        self.load_excel_files()

        self.setLayout(layout)
        self.setGeometry(300, 300, 400, 300)
        self.setModal(True)

    def create_target_label(self):
        target_label = QLabel(self)
        target_name = self.tr("(не выбрано)")
        if self.target_excel:
            target_name = os.path.basename(self.target_excel)
        target_label.setText(f"{self.tr('Целевой Excel:')} {target_name}")
        if self.target_excel:
            target_label.setToolTip(self.target_excel)
        return target_label

    def create_file_list_widget(self):
        """Создание QListWidget для отображения файлов."""
        file_list = QListWidget(self)
        file_list.setSelectionMode(QListView.SingleSelection)
        file_list.itemDoubleClicked.connect(self.select_file)
        return file_list

    def load_excel_files(self):
        """Загрузка и отображение файлов Excel из папки."""
        excel_files = self.get_excel_files()
        if not excel_files:
            self.show_warning_no_files()
            self.close()
        for file_path in excel_files:
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
        if self.target_excel:
            files = [f for f in files if os.path.abspath(f) != os.path.abspath(self.target_excel)]
        return sorted(files)

    def add_file_to_list(self, file_path):
        """Добавление файла в QListWidget."""
        file_name = os.path.basename(file_path)
        item = QListWidgetItem(file_name)
        item.setData(Qt.UserRole, file_path)
        item.setToolTip(file_name)
        self.file_list.addItem(item)

    def show_warning_no_files(self):
        """Показ предупреждения, если файлы Excel не найдены."""
        QMessageBox.warning(self, self.tr("Предупреждение"), self.tr("В указанной папке не найдено файлов Excel."))

    def select_file(self, item):
        """Выбор файла из списка."""
        self.selected_file = item.data(Qt.UserRole)
        self.accept()
