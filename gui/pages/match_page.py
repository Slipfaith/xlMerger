# gui/pages/match_page.py

import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QScrollArea, QFrame, QGridLayout, QComboBox,
    QPushButton, QHBoxLayout
)
from PySide6.QtCore import Signal, Qt

class MatchPage(QWidget):
    backClicked = Signal()
    nextClicked = Signal(dict, dict)
    saveClicked = Signal()
    loadClicked = Signal()

    def __init__(
        self,
        folder_path,
        selected_files,
        selected_sheets,
        columns,
        file_to_column=None,
        folder_to_column=None
    ):
        super().__init__()
        self.setWindowTitle("Сопоставление")
        self.folder_path = folder_path
        self.selected_files = selected_files
        self.selected_sheets = selected_sheets
        self.columns = columns
        self.file_to_column = file_to_column or {}
        self.folder_to_column = folder_to_column or {}

        self._all_comboboxes = []
        self._combobox_keys = []
        self._is_files = False
        self._is_folder_mapping = False
        self.files_in_root = []
        self.folders_with_excels = []

        self._init_files_and_folders()
        self._build_layout()

    def _init_files_and_folders(self):
        excel_exts = ('.xlsx', '.xls')
        self.files_in_root = []
        self.folders_with_excels = []

        if self.folder_path and os.path.isdir(self.folder_path):
            entries = os.listdir(self.folder_path)
            for entry in entries:
                full_path = os.path.join(self.folder_path, entry)
                if os.path.isfile(full_path) and entry.lower().endswith(excel_exts):
                    self.files_in_root.append(full_path)
                elif os.path.isdir(full_path):
                    files = os.listdir(full_path)
                    if any(f.lower().endswith(excel_exts) for f in files):
                        self.folders_with_excels.append((entry, full_path))

        if self.folders_with_excels:
            self._is_folder_mapping = True
            self._is_files = False
        elif self.files_in_root or self.selected_files:
            self._is_files = True
            self._is_folder_mapping = False
            if not self.selected_files:
                self.selected_files = self.files_in_root

    def _build_layout(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel(
            "Сопоставь имена файлов с колонками:" if self._is_files else "Сопоставь папки с колонками:"
        ))

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QFrame()
        scroll_layout = QGridLayout(scroll_content)
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)

        # Собираем список всех колонок (алфавитно)
        all_columns = [
            col for sheet in self.selected_sheets for col in self.columns[sheet]
            if isinstance(col, str) and col.strip()
        ]
        self._available_columns = [''] + sorted(set(all_columns))

        row, col = 0, 0
        self._all_comboboxes = []
        self._combobox_keys = []

        if self._is_files:
            self.file_to_column_widgets = {}
            for file_path in self.selected_files:
                if row >= 5:
                    row = 0
                    col += 2
                short = self.short_name_no_ext(os.path.basename(file_path), 5)
                file_label = QLabel(short)
                file_label.setToolTip(file_path)
                file_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                column_combobox = QComboBox()
                column_combobox.setMaximumWidth(100)
                for val in self._available_columns:
                    column_combobox.addItem(val)
                # Сетапим текущее значение если было
                if file_path in self.file_to_column:
                    idx = column_combobox.findText(self.file_to_column[file_path])
                    if idx != -1:
                        column_combobox.setCurrentIndex(idx)
                self._all_comboboxes.append(column_combobox)
                self._combobox_keys.append(file_path)
                self.file_to_column_widgets[file_path] = column_combobox
                scroll_layout.addWidget(file_label, row, col)
                scroll_layout.addWidget(column_combobox, row, col + 1)
                row += 1
            self.folder_to_column_widgets = {}
        elif self._is_folder_mapping:
            self.folder_to_column_widgets = {}
            for folder_name, folder_full_path in self.folders_with_excels:
                if row >= 5:
                    row = 0
                    col += 2
                folder_label = QLabel(folder_name)
                folder_label.setToolTip(folder_full_path)
                folder_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                column_combobox = QComboBox()
                column_combobox.setMaximumWidth(100)
                for val in self._available_columns:
                    column_combobox.addItem(val)
                if folder_full_path in self.folder_to_column:
                    idx = column_combobox.findText(self.folder_to_column[folder_full_path])
                    if idx != -1:
                        column_combobox.setCurrentIndex(idx)
                self._all_comboboxes.append(column_combobox)
                self._combobox_keys.append(folder_full_path)
                self.folder_to_column_widgets[folder_full_path] = column_combobox
                scroll_layout.addWidget(folder_label, row, col)
                scroll_layout.addWidget(column_combobox, row, col + 1)
                row += 1
            self.file_to_column_widgets = {}
        else:
            layout.addWidget(QLabel("Не найдены подходящие файлы или папки для сопоставления."))
            self.file_to_column_widgets = {}
            self.folder_to_column_widgets = {}

        layout.addWidget(scroll_area)
        button_layout = QHBoxLayout()
        save_button = QPushButton("Сохранить настройки")
        load_button = QPushButton("Загрузить настройки")
        back_button = QPushButton("Назад")
        next_button = QPushButton("Далее")
        save_button.clicked.connect(self.saveClicked.emit)
        load_button.clicked.connect(self.loadClicked.emit)
        back_button.clicked.connect(self.backClicked.emit)
        next_button.clicked.connect(self._on_next)
        button_layout.addWidget(save_button)
        button_layout.addWidget(load_button)
        button_layout.addWidget(back_button)
        button_layout.addWidget(next_button)
        layout.addLayout(button_layout)
        self.setLayout(layout)

    @staticmethod
    def short_name_no_ext(name, n=5):
        base, ext = os.path.splitext(name)
        if len(base) <= 2 * n:
            return base
        return f"{base[:n]}...{base[-n:]}"


    def apply_mapping(self, file_to_column=None, folder_to_column=None):
        """Применяет подгруженные маппинги к комбобоксам"""
        if file_to_column and hasattr(self, 'file_to_column_widgets'):
            for k, combo in self.file_to_column_widgets.items():
                val = file_to_column.get(k)
                if val is not None:
                    idx = combo.findText(val)
                    if idx != -1:
                        combo.setCurrentIndex(idx)
        if folder_to_column and hasattr(self, 'folder_to_column_widgets'):
            for k, combo in self.folder_to_column_widgets.items():
                val = folder_to_column.get(k)
                if val is not None:
                    idx = combo.findText(val)
                    if idx != -1:
                        combo.setCurrentIndex(idx)

    def _on_next(self):
        file_to_column = {}
        folder_to_column = {}
        if self._is_files:
            for k, v in self.file_to_column_widgets.items():
                file_to_column[k] = v.currentText()
        elif self._is_folder_mapping:
            for k, v in self.folder_to_column_widgets.items():
                folder_to_column[k] = v.currentText()
        self.nextClicked.emit(file_to_column, folder_to_column)
