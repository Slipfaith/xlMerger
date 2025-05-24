import os
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QListWidget, QListWidgetItem, QHBoxLayout, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QListView

class ExcelFileSelector(QDialog):
    def __init__(self, folder_path, selected_files=None):
        super().__init__()
        self.folder_path = folder_path
        self.selected_files = selected_files
        self.selected_file = None
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(self.tr("Выберите файл Excel"))
        layout = QVBoxLayout()

        self.file_list = self.create_file_list_widget()
        layout.addWidget(QLabel(self.tr("Дважды щелкните, чтобы выбрать файл Excel:")))
        layout.addWidget(self.file_list)

        self.load_excel_files()

        self.setLayout(layout)
        self.setGeometry(300, 300, 400, 300)
        self.show()

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
        return [
            os.path.join(root, file)
            for root, _, files in os.walk(self.folder_path)
            for file in files if file.endswith(('.xlsx', '.xls')) and
                                 (not self.selected_files or file in self.selected_files)
        ]

    def add_file_to_list(self, file_path):
        """Добавление файла в QListWidget."""
        item = QListWidgetItem(os.path.splitext(os.path.basename(file_path))[0])
        item.setData(Qt.UserRole, file_path)
        self.file_list.addItem(item)

    def show_warning_no_files(self):
        """Показ предупреждения, если файлы Excel не найдены."""
        QMessageBox.warning(self, self.tr("Предупреждение"), self.tr("В указанной папке не найдено файлов Excel."))

    def select_file(self, item):
        """Выбор файла из списка."""
        self.selected_file = item.data(Qt.UserRole)
        self.accept()
