# gui/main_page.py

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QCheckBox, QRadioButton,
    QListWidget, QLineEdit
)
from PySide6.QtCore import Qt, Signal
from utils.i18n import tr, i18n

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
        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()
        self.apply_modern_style()

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)
        layout.addLayout(self.create_folder_selection_layout())
        layout.addLayout(self.create_excel_selection_layout())
        layout.addWidget(self.create_sheet_selection_layout(), alignment=Qt.AlignLeft)
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
        self.folder_entry.setPlaceholderText(tr("Перетащи или кликни дважды"))
        self.folder_entry.filesSelected.connect(self.filesSelected)
        self.folder_entry.folderSelected.connect(self.folderSelected)
        self.folder_label = QLabel()
        layout.addWidget(self.folder_label)
        layout.addWidget(self.folder_entry)
        return layout

    def create_excel_selection_layout(self):
        layout = QHBoxLayout()
        self.excel_file_entry = DragDropLineEdit(mode='file')
        self.excel_file_entry.setPlaceholderText(tr("Перетащи или кликни дважды"))
        self.excel_file_entry.fileSelected.connect(self.excelFileSelected)
        self.excel_label = QLabel()
        layout.addWidget(self.excel_label)
        layout.addWidget(self.excel_file_entry)
        return layout

    def create_sheet_selection_layout(self):
        container = QWidget(self)
        layout = QVBoxLayout(container)
        layout.setSpacing(6)
        self.sheet_label = QLabel()
        self.sheet_list = QListWidget(self)
        self.sheet_list.setSelectionMode(QListWidget.MultiSelection)
        self.sheet_list.setFixedHeight(100)
        button_layout = QHBoxLayout()
        self.deselect_all_button = QPushButton(tr("Не выбрать все"), self)
        self.select_all_button = QPushButton(tr("Выбрать все"), self)
        self.deselect_all_button.clicked.connect(self.deselect_all_sheets)
        self.select_all_button.clicked.connect(self.select_all_sheets)
        button_layout.addWidget(self.deselect_all_button)
        button_layout.addWidget(self.select_all_button)
        layout.addWidget(self.sheet_label)
        layout.addWidget(self.sheet_list)
        layout.addLayout(button_layout)
        container.setFixedWidth(300)
        return container

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
        self.copy_column_label = QLabel()
        self.copy_column_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        layout.addWidget(self.copy_column_label)
        layout.addWidget(self.copy_column_entry)
        return layout

    def create_skip_first_row_checkbox(self):
        self.skip_first_row_checkbox = QCheckBox(tr("Первая строка — заголовок в переводах"), self)
        return self.skip_first_row_checkbox

    def create_copy_method_selection_layout(self):
        layout = QHBoxLayout()
        self.copy_by_matching_radio = QRadioButton(tr("Нет пустых/скрытых строк в xlsx"), self)
        self.copy_by_matching_radio.setChecked(True)
        self.copy_by_row_number_radio = QRadioButton(tr("Есть пустые/скрытые строки в xlsx"), self)
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
        self.preview_button = QPushButton(tr("Настроить"), self)
        self.preview_button.clicked.connect(self.previewTriggered)
        return self.preview_button

    def create_process_button(self):
        self.process_button = QPushButton(tr("Начать"), self)
        self.process_button.clicked.connect(self.processTriggered)
        return self.process_button

    def update_button_heights(self):
        """Adjust button height to fit their text."""
        buttons = [
            self.deselect_all_button,
            self.select_all_button,
            self.preview_button,
            self.process_button,
        ]
        for button in buttons:
            button.setFixedHeight(button.sizeHint().height())

    def clear(self):
        self.folder_entry.clear()
        self.excel_file_entry.clear()
        self.copy_column_entry.clear()
        self.skip_first_row_checkbox.setChecked(False)
        self.copy_by_matching_radio.setChecked(True)
        self.sheet_list.clear()

    def retranslate_ui(self):
        self.folder_entry.setPlaceholderText(tr("Перетащи или кликни дважды"))
        self.excel_file_entry.setPlaceholderText(tr("Перетащи или кликни дважды"))
        self.folder_label.setText(tr("Папка/эксели с переводами:"))
        self.excel_label.setText(tr("Целевой Excel:"))
        self.copy_column_label.setText(tr("Из какой колонки копировать? (буква колонки):"))
        self.skip_first_row_checkbox.setText(tr("Первая строка — заголовок в переводах"))
        self.copy_by_matching_radio.setText(tr("Нет пустых/скрытых строк в xlsx"))
        self.copy_by_row_number_radio.setText(tr("Есть пустые/скрытые строки в xlsx"))
        self.deselect_all_button.setText(tr("Не выбрать все"))
        self.select_all_button.setText(tr("Выбрать все"))
        self.sheet_label.setText(tr("Выберите листы:"))
        self.preview_button.setText(tr("Настроить"))
        self.process_button.setText(tr("Начать"))
        self.update_button_heights()

    def apply_modern_style(self):
        self.setStyleSheet(
            """
            QLabel {
                font-size: 14px;
                color: #2c2c2c;
            }
            QLineEdit, QListWidget {
                border: 1px solid #bdbdbd;
                border-radius: 4px;
                padding: 4px;
            }
            QPushButton {
                background-color: #1976d2;
                color: #ffffff;
                padding: 6px 12px;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #1565c0;
            }
            QCheckBox, QRadioButton {
                color: #2c2c2c;
                padding: 2px;
            }
            """
        )
        self.update_button_heights()
