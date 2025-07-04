# gui/main_page.py

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QCheckBox, QRadioButton,
    QListWidget, QLineEdit
)
from PySide6.QtCore import Qt, Signal

from core.drag_drop import DragDropLineEdit

class MainPageWidget(QWidget):
    # Сигналы для FileProcessorApp
    folderSelected = Signal(str)
    filesSelected = Signal(list)
    excelFileSelected = Signal(str)
    processTriggered = Signal()
    previewTriggered = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.addLayout(self.create_folder_selection_layout())
        layout.addLayout(self.create_excel_selection_layout())
        layout.addLayout(self.create_sheet_selection_layout())
        layout.addLayout(self.create_copy_column_layout())
        layout.addWidget(self.create_skip_first_row_checkbox())
        layout.addLayout(self.create_copy_method_selection_layout())
        layout.addWidget(self.create_preview_button(), alignment=Qt.AlignRight)
        layout.addWidget(self.create_process_button(), alignment=Qt.AlignCenter)
        self.setLayout(layout)

        # --- ВАЖНО: Связываем переключатели ---
        self.copy_by_matching_radio.toggled.connect(self.toggle_skip_first_row_checkbox)
        self.copy_by_row_number_radio.toggled.connect(self.toggle_skip_first_row_checkbox)
        self.toggle_skip_first_row_checkbox()  # выставить корректное состояние при старте

    def create_folder_selection_layout(self):
        layout = QHBoxLayout()
        self.folder_entry = DragDropLineEdit(mode='files_or_folder')
        self.folder_entry.setPlaceholderText("Перетащи или кликни дважды")
        self.folder_entry.filesSelected.connect(self.filesSelected)
        self.folder_entry.folderSelected.connect(self.folderSelected)
        layout.addWidget(QLabel("Папка/эксели с переводами:"))
        layout.addWidget(self.folder_entry)
        return layout

    def create_excel_selection_layout(self):
        layout = QHBoxLayout()
        self.excel_file_entry = DragDropLineEdit(mode='file')
        self.excel_file_entry.setPlaceholderText("Перетащи или кликни дважды")
        self.excel_file_entry.fileSelected.connect(self.excelFileSelected)
        layout.addWidget(QLabel("Целевой Excel:"))
        layout.addWidget(self.excel_file_entry)
        return layout

    def create_sheet_selection_layout(self):
        layout = QVBoxLayout()
        self.sheet_list = QListWidget(self)
        self.sheet_list.setSelectionMode(QListWidget.MultiSelection)
        self.sheet_list.setFixedHeight(100)
        button_layout = QHBoxLayout()
        deselect_all_button = QPushButton("Не выбрать все", self)
        select_all_button = QPushButton("Выбрать все", self)
        deselect_all_button.clicked.connect(self.deselect_all_sheets)
        select_all_button.clicked.connect(self.select_all_sheets)
        button_layout.addWidget(deselect_all_button)
        button_layout.addWidget(select_all_button)
        layout.addWidget(QLabel("Выберите листы:"))
        layout.addWidget(self.sheet_list)
        layout.addLayout(button_layout)
        return layout

    def deselect_all_sheets(self):
        for index in range(self.sheet_list.count()):
            self.sheet_list.item(index).setCheckState(Qt.Unchecked)

    def select_all_sheets(self):
        for index in range(self.sheet_list.count()):
            self.sheet_list.item(index).setCheckState(Qt.Checked)

    def create_copy_column_layout(self):
        layout = QHBoxLayout()
        self.copy_column_entry = QLineEdit(self)
        self.copy_column_entry.setMaximumWidth(100)
        copy_column_label = QLabel("Из какой колонки копировать? (буква колонки):")
        copy_column_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        layout.addWidget(copy_column_label)
        layout.addWidget(self.copy_column_entry)
        return layout

    def create_skip_first_row_checkbox(self):
        self.skip_first_row_checkbox = QCheckBox("Первая строка — заголовок в переводах", self)
        return self.skip_first_row_checkbox

    def create_copy_method_selection_layout(self):
        layout = QHBoxLayout()
        self.copy_by_matching_radio = QRadioButton("Нет пустых/скрытых строк в xlsx", self)
        self.copy_by_matching_radio.setChecked(True)
        self.copy_by_row_number_radio = QRadioButton("Есть пустые/скрытые строки в xlsx", self)
        layout.addWidget(self.copy_by_matching_radio)
        layout.addWidget(self.copy_by_row_number_radio)
        return layout

    def toggle_skip_first_row_checkbox(self):
        """
        Деактивирует чекбокс 'Первая строка — заголовок в переводах'
        если выбран режим 'Есть пустые/скрытые строки в xlsx'.
        """
        if self.copy_by_row_number_radio.isChecked():
            self.skip_first_row_checkbox.setChecked(False)
            self.skip_first_row_checkbox.setEnabled(False)
        else:
            self.skip_first_row_checkbox.setEnabled(True)

    def create_preview_button(self):
        preview_button = QPushButton("Превью", self)
        preview_button.clicked.connect(self.previewTriggered)
        return preview_button

    def create_process_button(self):
        process_button = QPushButton("Начать", self)
        process_button.setStyleSheet("""
            QPushButton {
                background-color: #f47929;
                color: white;
                border-radius: 6px;
                padding: 4px 14px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #65d88f;  /* светло-зелёный, неяркий */
                color: #222;
            }
            QPushButton:pressed {
                background-color: #41bb6f;
                color: white;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #888888;
            }
        """)
        process_button.clicked.connect(self.processTriggered)
        return process_button

    def clear(self):
        self.folder_entry.clear()
        self.excel_file_entry.clear()
        self.copy_column_entry.clear()
        self.skip_first_row_checkbox.setChecked(False)
        self.copy_by_matching_radio.setChecked(True)
        self.sheet_list.clear()
