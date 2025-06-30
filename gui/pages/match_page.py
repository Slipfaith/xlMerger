import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QScrollArea, QFrame, QGridLayout, QComboBox,
    QPushButton, QHBoxLayout, QCheckBox
)
from PySide6.QtCore import Signal, Qt
from utils.i18n import tr, i18n


class MatchPage(QWidget):
    backClicked = Signal()
    nextClicked = Signal(dict, dict, bool)
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
        self.setWindowTitle(tr("Сопоставление"))
        self.folder_path = folder_path
        self.selected_files = [os.path.abspath(f) for f in (selected_files or [])]
        self.selected_sheets = selected_sheets
        self.columns = columns
        self.file_to_column = {os.path.abspath(k): v for k, v in (file_to_column or {}).items()}
        self.folder_to_column = {os.path.abspath(k): v for k, v in (folder_to_column or {}).items()}

        self._all_comboboxes = []
        self._combobox_keys = []
        self._is_files = False
        self._is_folder_mapping = False
        self.files_in_root = []
        self.folders_with_excels = []

        self._init_files_and_folders()
        self._build_layout()
        # Применяем маппинг после построения интерфейса
        self.apply_mapping(self.file_to_column, self.folder_to_column)
        i18n.language_changed.connect(self.retranslate_ui)
        self.retranslate_ui()

    def _init_files_and_folders(self):
        excel_exts = ('.xlsx', '.xls')
        self.files_in_root = []
        self.folders_with_excels = []

        if self.folder_path and os.path.isdir(self.folder_path):
            entries = os.listdir(self.folder_path)
            for entry in entries:
                full_path = os.path.join(self.folder_path, entry)
                if os.path.isfile(full_path) and entry.lower().endswith(excel_exts):
                    self.files_in_root.append(os.path.abspath(full_path))
                elif os.path.isdir(full_path):
                    files = os.listdir(full_path)
                    if any(f.lower().endswith(excel_exts) for f in files):
                        self.folders_with_excels.append((entry, os.path.abspath(full_path)))

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
        label_text = "Сопоставь имена файлов с колонками:" if self._is_files else "Сопоставь папки с колонками:"
        self.top_label = QLabel(label_text)
        layout.addWidget(self.top_label)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QFrame()
        scroll_layout = QGridLayout(scroll_content)
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)

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
                column_combobox.setMaximumWidth(120)
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
                column_combobox.setMaximumWidth(120)
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

        self.format_checkbox = QCheckBox(tr("Копировать с сохранением форматирования"))
        self.format_checkbox.setChecked(False)
        self.format_checkbox.setStyleSheet("QCheckBox { padding: 4px; }")
        layout.addWidget(self.format_checkbox)

        button_layout = QHBoxLayout()
        self.save_button = QPushButton(tr("Сохранить настройки"))
        self.load_button = QPushButton(tr("Загрузить настройки"))
        self.back_button = QPushButton(tr("Назад"))
        self.next_button = QPushButton(tr("Далее"))
        self.save_button.clicked.connect(self.saveClicked.emit)
        self.load_button.clicked.connect(self.loadClicked.emit)
        self.back_button.clicked.connect(self.backClicked.emit)
        self.next_button.clicked.connect(self._on_next)
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.load_button)
        button_layout.addWidget(self.back_button)
        button_layout.addWidget(self.next_button)
        layout.addLayout(button_layout)
        self.setLayout(layout)

        self._populate_comboboxes()
        for combo in self._all_comboboxes:
            combo.currentIndexChanged.connect(self._update_all_comboboxes)

    @staticmethod
    def short_name_no_ext(name, n=5):
        base, ext = os.path.splitext(name)
        if len(base) <= 2 * n:
            return base
        return f"{base[:n]}...{base[-n:]}"

    def get_current_mapping(self):
        """Возвращает {короткое имя: колонка}, только для непустых значений."""
        mapping = {}
        if self._is_files:
            for file, combo in self.file_to_column_widgets.items():
                val = combo.currentText()
                if val:
                    # Берём имя файла без расширения (или с, если нужны дубликаты)
                    mapping[os.path.splitext(os.path.basename(file))[0]] = val
        elif self._is_folder_mapping:
            for folder, combo in self.folder_to_column_widgets.items():
                val = combo.currentText()
                if val:
                    # Берём имя папки
                    mapping[os.path.basename(folder)] = val
        return mapping

    def _populate_comboboxes(self):
        selected = self._get_current_selections()
        for idx, combo in enumerate(self._all_comboboxes):
            current_val = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            not_used = [c for c in self._available_columns if c not in selected or c == current_val]
            not_used = sorted(not_used)
            used_in_others = [c for c in self._available_columns if c in selected and c != current_val and c != '']
            used_in_others = sorted(used_in_others)
            combo.addItem('')
            for val in not_used:
                if val != '':
                    combo.addItem(val)
            if used_in_others:
                combo.addItem('---')
                for val in used_in_others:
                    combo.addItem(val)
            if current_val:
                ix = combo.findText(current_val)
                if ix >= 0:
                    combo.setCurrentIndex(ix)
            combo.blockSignals(False)

        # Восстанавливаем выбор из file_to_column/folder_to_column после построения
        if self._is_files:
            for file_path, combo in self.file_to_column_widgets.items():
                if file_path in self.file_to_column:
                    idx = combo.findText(self.file_to_column[file_path])
                    if idx != -1:
                        combo.setCurrentIndex(idx)
        elif self._is_folder_mapping:
            for folder_full_path, combo in self.folder_to_column_widgets.items():
                if folder_full_path in self.folder_to_column:
                    idx = combo.findText(self.folder_to_column[folder_full_path])
                    if idx != -1:
                        combo.setCurrentIndex(idx)

    def _update_all_comboboxes(self):
        self._populate_comboboxes()

    def _get_current_selections(self):
        vals = set()
        for combo in self._all_comboboxes:
            val = combo.currentText()
            if val and val != '' and val != '---':
                vals.add(val)
        return vals

    def apply_mapping(self, file_to_column=None, folder_to_column=None):
        """
        file_to_column: dict, где ключ — короткое имя файла (например, 'de-DE')
        folder_to_column: dict, где ключ — короткое имя папки (например, 'de-DE')
        """
        # Для файлов (если такая логика вообще используется)
        if file_to_column and hasattr(self, 'file_to_column_widgets'):
            for k, combo in self.file_to_column_widgets.items():
                shortname = os.path.splitext(os.path.basename(k))[0]
                val = file_to_column.get(shortname)
                if combo and val is not None:
                    idx = combo.findText(val)
                    if idx != -1:
                        combo.setCurrentIndex(idx)
        # Для папок
        if folder_to_column and hasattr(self, 'folder_to_column_widgets'):
            for k, combo in self.folder_to_column_widgets.items():
                basename = os.path.basename(k)
                val = folder_to_column.get(basename)
                if combo and val is not None:
                    idx = combo.findText(val)
                    if idx != -1:
                        combo.setCurrentIndex(idx)
        self._populate_comboboxes()  # чтобы порядок был правильный

    def _on_next(self):
        file_to_column = {}
        folder_to_column = {}
        if self._is_files:
            for k, v in self.file_to_column_widgets.items():
                file_to_column[k] = v.currentText()
        elif self._is_folder_mapping:
            for k, v in self.folder_to_column_widgets.items():
                folder_to_column[k] = v.currentText()
        preserve = self.format_checkbox.isChecked()
        self.nextClicked.emit(file_to_column, folder_to_column, preserve)

    def retranslate_ui(self):
        self.setWindowTitle(tr("Сопоставление"))
        label_text = tr("Сопоставь имена файлов с колонками:") if self._is_files else tr("Сопоставь папки с колонками:")
        self.top_label.setText(label_text)
        self.save_button.setText(tr("Сохранить настройки"))
        self.load_button.setText(tr("Загрузить настройки"))
        self.back_button.setText(tr("Назад"))
        self.next_button.setText(tr("Далее"))
        self.format_checkbox.setText(tr("Копировать с сохранением форматирования"))
