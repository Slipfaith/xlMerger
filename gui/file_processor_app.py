import os
import json
import traceback
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QVBoxLayout, QHBoxLayout,
    QComboBox, QMessageBox, QProgressBar, QScrollArea, QFrame, QCheckBox, QListWidget,
    QListWidgetItem, QRadioButton, QGridLayout, QStackedWidget, QLineEdit, QPushButton, QFileDialog
)
from PySide6.QtCore import Qt, Signal

from gui.excel_previewer import ExcelPreviewer
from gui.excel_file_selector import ExcelFileSelector
from core.excel_processor import ExcelProcessor
from core.drag_drop import DragDropLineEdit
from utils.utils import excel_column_to_index


def short_name_no_ext(name, n=5):
    base, ext = os.path.splitext(name)
    if len(base) <= 2 * n:
        return base
    return f"{base[:n]}...{base[-n:]}"


class FileProcessorApp(QWidget):
    copyingStarted = Signal()

    def __init__(self):
        super().__init__()
        self.folder_path = ''
        self.selected_files = []
        self.excel_file_path = ''
        self.copy_column = ''
        self.selected_sheets = []
        self.sheet_names = []
        self.header_row = {}
        self.columns = {}
        self.sheet_to_column = {}
        self.folder_to_column = {}
        self.file_to_column = {}
        self.skip_first_row = False
        self.copy_by_row_number = False
        self.progress_bar = None

        self.stack = QStackedWidget(self)
        self.page_main = self.create_main_page()
        self.stack.addWidget(self.page_main)
        layout = QVBoxLayout()
        layout.addWidget(self.stack)
        self.setLayout(layout)
        self.setWindowTitle("Обработка файлов")
        self.center_window()
        self.setGeometry(300, 300, 600, 400)
        self.show()

    def create_main_page(self):
        widget = QWidget()
        layout = QVBoxLayout()
        layout.addLayout(self.create_folder_selection_layout())
        layout.addLayout(self.create_excel_selection_layout())
        layout.addLayout(self.create_sheet_selection_layout())
        layout.addLayout(self.create_copy_column_layout())
        layout.addWidget(self.create_skip_first_row_checkbox())
        layout.addLayout(self.create_copy_method_selection_layout())
        layout.addWidget(self.create_preview_button(), alignment=Qt.AlignRight)
        layout.addWidget(self.create_process_button(), alignment=Qt.AlignCenter)
        widget.setLayout(layout)
        return widget

    def create_folder_selection_layout(self):
        layout = QHBoxLayout()
        self.folder_entry = DragDropLineEdit(mode='files_or_folder')
        self.folder_entry.setPlaceholderText("Перетащи папку или эксели, или кликни дважды")
        self.folder_entry.filesSelected.connect(self.on_files_selected)
        self.folder_entry.folderSelected.connect(self.on_folder_selected)
        layout.addWidget(QLabel("Папка/файлы переводов:"))
        layout.addWidget(self.folder_entry)
        return layout

    def create_excel_selection_layout(self):
        layout = QHBoxLayout()
        self.excel_file_entry = DragDropLineEdit(mode='file')
        self.excel_file_entry.fileSelected.connect(self.on_excel_file_selected)
        layout.addWidget(QLabel("Файл Excel:"))
        layout.addWidget(self.excel_file_entry)
        return layout

    def on_folder_selected(self, path):
        self.folder_path = path
        self.selected_files = []

    def on_files_selected(self, files):
        self.selected_files = files
        self.folder_path = os.path.dirname(files[0]) if files else ''

    def on_excel_file_selected(self, file):
        self.excel_file_path = file
        self.load_sheet_names()

    def load_sheet_names(self):
        try:
            self.sheet_names = ExcelProcessor.get_sheet_names(self.excel_file_path)
            self.populate_sheet_list()
        except Exception as e:
            self.log_error(e)
            QMessageBox.critical(self, "Ошибка", "Не удалось загрузить листы из файла Excel: " + str(e))

    def populate_sheet_list(self):
        self.sheet_list.clear()
        for sheet_name in self.sheet_names:
            item = QListWidgetItem(sheet_name)
            item.setCheckState(Qt.Checked)
            self.sheet_list.addItem(item)

    def create_sheet_selection_layout(self):
        layout = QVBoxLayout()
        self.sheet_list = QListWidget(self)
        self.sheet_list.setSelectionMode(QListWidget.MultiSelection)
        self.sheet_list.setFixedHeight(100)
        button_layout = QHBoxLayout()
        deselect_all_button = QPushButton("Снять выделение со всех", self)
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
        copy_column_label = QLabel("Столбец копирования:")
        copy_column_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        layout.addWidget(copy_column_label)
        layout.addWidget(self.copy_column_entry)
        return layout

    def create_skip_first_row_checkbox(self):
        self.skip_first_row_checkbox = QCheckBox("Первая строка — заголовок в переводах", self)
        self.skip_first_row_checkbox.stateChanged.connect(self.toggle_skip_first_row)
        return self.skip_first_row_checkbox

    def toggle_skip_first_row(self, state):
        self.skip_first_row = state == Qt.Checked

    def create_copy_method_selection_layout(self):
        layout = QHBoxLayout()
        self.copy_by_matching_radio = QRadioButton("Нет пустых/скрытых строк в xlsx", self)
        self.copy_by_matching_radio.setChecked(True)
        self.copy_by_matching_radio.toggled.connect(self.toggle_copy_method)
        self.copy_by_row_number_radio = QRadioButton("Есть пустые/скрытые строки в xlsx", self)
        layout.addWidget(self.copy_by_matching_radio)
        layout.addWidget(self.copy_by_row_number_radio)
        return layout

    def toggle_copy_method(self):
        self.copy_by_row_number = self.copy_by_row_number_radio.isChecked()
        self.skip_first_row_checkbox.setEnabled(not self.copy_by_row_number)
        if self.copy_by_row_number:
            self.skip_first_row_checkbox.setChecked(False)

    def create_preview_button(self):
        preview_button = QPushButton("Превью", self)
        preview_button.clicked.connect(self.select_excel_file_for_preview)
        return preview_button

    def select_excel_file_for_preview(self):
        if not self.folder_path:
            QMessageBox.warning(self, "Предупреждение", "Сначала выбери папку или эксели с переводами.")
            return
        dialog = ExcelFileSelector(self.folder_path, self.selected_files)
        if dialog.exec():
            selected_file = dialog.selected_file
            if selected_file:
                self.preview_window = ExcelPreviewer(selected_file)
                self.preview_window.show()

    def create_process_button(self):
        process_button = QPushButton("Начать", self)
        process_button.setStyleSheet("background-color: #f47929; color: white;")
        process_button.clicked.connect(self.process_files)
        return process_button

    def process_files(self):
        self.folder_path = self.folder_entry.text()
        self.copy_column = self.copy_column_entry.text()
        if not self.selected_files and self.folder_path and os.path.isdir(self.folder_path):
            self.selected_files = [
                os.path.join(self.folder_path, fname)
                for fname in os.listdir(self.folder_path)
                if fname.lower().endswith(('.xlsx', '.xls'))
            ]
        if not self.excel_file_path:
            self.excel_file_path = self.excel_file_entry.text()
        if not self.validate_inputs():
            return
        self.selected_sheets = [
            self.sheet_list.item(i).text()
            for i in range(self.sheet_list.count())
            if self.sheet_list.item(i).checkState() == Qt.Checked
        ]
        if not self.selected_sheets:
            QMessageBox.critical(self, "Ошибка", "Выбери хотя бы один лист.")
            return
        self.go_to_header_page()

    def validate_inputs(self):
        if self.selected_files:
            for f in self.selected_files:
                if not os.path.isfile(f):
                    QMessageBox.critical(self, "Ошибка", f"Файл Excel не найден: {f}")
                    return False
        elif self.folder_path and not os.path.isdir(self.folder_path):
            QMessageBox.critical(self, "Ошибка", "Указанная папка не существует.")
            return False
        if not self.excel_file_path or not os.path.isfile(self.excel_file_path):
            QMessageBox.critical(self, "Ошибка", "Указанный файл Excel не существует.")
            return False
        if not self.copy_column:
            QMessageBox.critical(self, "Ошибка", "Укажи столбец для копирования.")
            return False
        return True

    def create_header_page(self):
        page = QWidget()
        page.setWindowTitle("Выбери строку заголовка")
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Выбери номер строки заголовка для каждого листа:"))
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QFrame()
        scroll_layout = QGridLayout(scroll_content)
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)
        self.sheet_to_header_row = {}
        for row, sheet_name in enumerate(self.selected_sheets):
            sheet_label = QLabel(sheet_name)
            header_row_combobox = QComboBox()
            header_row_combobox.setMaximumWidth(100)
            header_row_combobox.addItems([str(i) for i in range(1, 11)])
            header_row_combobox.setCurrentIndex(0)
            scroll_layout.addWidget(sheet_label, row, 0)
            scroll_layout.addWidget(header_row_combobox, row, 1)
            self.sheet_to_header_row[sheet_name] = header_row_combobox
        layout.addWidget(scroll_area)
        button_layout = QHBoxLayout()
        back_button = QPushButton("Назад")
        next_button = QPushButton("Далее")
        back_button.clicked.connect(self.go_to_main_page)
        next_button.clicked.connect(self.load_columns_and_go_to_sheet_column)
        button_layout.addWidget(back_button)
        button_layout.addWidget(next_button)
        layout.addLayout(button_layout)
        page.setLayout(layout)
        return page

    def go_to_header_page(self):
        self.page_header = self.create_header_page()
        self.stack.addWidget(self.page_header)
        self.stack.setCurrentWidget(self.page_header)

    def go_to_main_page(self):
        self.stack.setCurrentWidget(self.page_main)

    def load_columns_and_go_to_sheet_column(self):
        try:
            self.header_row = {sheet: int(combo.currentText()) - 1 for sheet, combo in self.sheet_to_header_row.items()}
            self.columns = {}
            for sheet, header_row in self.header_row.items():
                self.columns[sheet] = ExcelProcessor.get_sheet_columns(self.excel_file_path, sheet, header_row)
            self.go_to_sheet_column_page()
        except Exception as e:
            self.log_error(e)
            QMessageBox.critical(self, "Ошибка", "Произошла ошибка при загрузке столбцов: " + str(e))

    def create_sheet_column_page(self):
        page = QWidget()
        page.setWindowTitle("Соответствие лист-столбец")
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Из какого столбца на каждом листе копировать?"))
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QFrame()
        scroll_layout = QGridLayout(scroll_content)
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)
        self.sheet_to_column = {}
        for row, sheet_name in enumerate(self.selected_sheets):
            sheet_label = QLabel(sheet_name)
            column_entry = QLineEdit()
            column_entry.setMaximumWidth(100)
            column_entry.setText(self.copy_column)
            scroll_layout.addWidget(sheet_label, row, 0)
            scroll_layout.addWidget(column_entry, row, 1)
            self.sheet_to_column[sheet_name] = column_entry
        layout.addWidget(scroll_area)
        button_layout = QHBoxLayout()
        back_button = QPushButton("Назад")
        next_button = QPushButton("Далее")
        back_button.clicked.connect(self.go_to_header_page)
        next_button.clicked.connect(self.go_to_match_page)
        button_layout.addWidget(back_button)
        button_layout.addWidget(next_button)
        layout.addLayout(button_layout)
        page.setLayout(layout)
        return page

    def go_to_sheet_column_page(self):
        self.page_sheet_column = self.create_sheet_column_page()
        self.stack.addWidget(self.page_sheet_column)
        self.stack.setCurrentWidget(self.page_sheet_column)

    # ======== Изменённый блок: create_match_page + поддержка алфавитного порядка и "уже выбрано" ========
    def create_match_page(self):
        page = QWidget()
        excel_exts = ('.xlsx', '.xls')
        is_files = False
        is_folder_mapping = False
        files_in_root = []
        folders_with_excels = []

        if self.folder_path and os.path.isdir(self.folder_path):
            entries = os.listdir(self.folder_path)
            for entry in entries:
                full_path = os.path.join(self.folder_path, entry)
                if os.path.isfile(full_path) and entry.lower().endswith(excel_exts):
                    files_in_root.append(full_path)
                elif os.path.isdir(full_path):
                    files = os.listdir(full_path)
                    if any(f.lower().endswith(excel_exts) for f in files):
                        folders_with_excels.append((entry, full_path))

        if folders_with_excels:
            is_folder_mapping = True
            is_files = False
        elif files_in_root or self.selected_files:
            is_files = True
            is_folder_mapping = False
            if not self.selected_files:
                self.selected_files = files_in_root

        page.setWindowTitle("Соответствие файл-столбец" if is_files else "Соответствие папка-столбец")
        layout = QVBoxLayout()
        layout.addWidget(QLabel(
            "Сопоставь имена файлов с колонками:" if is_files else "Сопоставь папки с колонками:"
        ))

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QFrame()
        scroll_layout = QGridLayout(scroll_content)
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)

        # Алфавитный порядок!
        all_columns = [col for sheet in self.selected_sheets for col in self.columns[sheet] if
                       isinstance(col, str) and col.strip()]
        available_columns = [''] + sorted(set(all_columns))

        row, col = 0, 0
        self._all_comboboxes = []
        self._combobox_keys = []  # Чтобы знать, какой combobox за какой файл/папку отвечает

        if is_files:
            self.file_to_column = {}
            for file_path in self.selected_files:
                if row >= 5:
                    row = 0
                    col += 2
                short = short_name_no_ext(os.path.basename(file_path), 5)
                file_label = QLabel(short)
                file_label.setToolTip(file_path)
                file_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                column_combobox = QComboBox()
                column_combobox.setMaximumWidth(100)
                self._all_comboboxes.append(column_combobox)
                self._combobox_keys.append(file_path)
                self.file_to_column[file_path] = column_combobox
                scroll_layout.addWidget(file_label, row, col)
                scroll_layout.addWidget(column_combobox, row, col + 1)
                row += 1
            self.folder_to_column = {}
        elif is_folder_mapping:
            self.folder_to_column = {}
            for folder_name, folder_full_path in folders_with_excels:
                if row >= 5:
                    row = 0
                    col += 2
                folder_label = QLabel(folder_name)
                folder_label.setToolTip(folder_full_path)
                folder_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                column_combobox = QComboBox()
                column_combobox.setMaximumWidth(100)
                self._all_comboboxes.append(column_combobox)
                self._combobox_keys.append(folder_full_path)
                self.folder_to_column[folder_full_path] = column_combobox
                scroll_layout.addWidget(folder_label, row, col)
                scroll_layout.addWidget(column_combobox, row, col + 1)
                row += 1
            self.file_to_column = {}
        else:
            label = QLabel("Не найдены подходящие файлы или папки для сопоставления.")
            layout.addWidget(label)
            self.file_to_column = {}
            self.folder_to_column = {}

        layout.addWidget(scroll_area)

        button_layout = QHBoxLayout()
        save_button = QPushButton("Сохранить настройки")
        load_button = QPushButton("Загрузить настройки")
        back_button = QPushButton("Назад")
        next_button = QPushButton("Далее")
        save_button.clicked.connect(self.save_mapping_settings)
        load_button.clicked.connect(self.load_mapping_settings)
        back_button.clicked.connect(self.go_to_sheet_column_page)
        next_button.clicked.connect(self.go_to_confirmation_page)
        button_layout.addWidget(save_button)
        button_layout.addWidget(load_button)
        button_layout.addWidget(back_button)
        button_layout.addWidget(next_button)
        layout.addLayout(button_layout)
        page.setLayout(layout)

        # Заполняем все комбобоксы
        for combo in self._all_comboboxes:
            combo.blockSignals(True)
            combo.clear()
            for val in available_columns:
                combo.addItem(val)
            combo.blockSignals(False)

        # Связываем обновление
        for combo in self._all_comboboxes:
            combo.currentIndexChanged.connect(self.update_all_comboboxes_for_matching)

        self._matching_available_columns = available_columns  # Сохраним для обновлений

        return page

    def update_all_comboboxes_for_matching(self):
        # Список выбранных значений (кроме пустых)
        used_values = set()
        for combo in self._all_comboboxes:
            val = combo.currentText()
            if val and val != "":
                used_values.add(val)
        # Обновить каждый комбобокс
        for idx, combo in enumerate(self._all_comboboxes):
            current_val = combo.currentText()
            combo.blockSignals(True)
            combo.clear()
            # Сначала — алфавитно все невыбранные нигде значения и (если есть) текущее
            base_columns = self._matching_available_columns[1:]  # Без первого пустого
            not_used = [v for v in base_columns if v not in used_values or v == current_val]
            not_used = sorted(not_used)
            # Уже выбранные в других комбобоксах
            used_in_others = [v for v in base_columns if v in used_values and v != current_val]
            used_in_others = sorted(used_in_others)
            # Пустой элемент
            combo.addItem('')
            # Основные значения
            for val in not_used:
                combo.addItem(val)
            # Разделитель
            if used_in_others:
                combo.addItem('---')
                for val in used_in_others:
                    combo.addItem(f"{val}")
            # Восстанавливаем текущее значение
            if current_val:
                ix = combo.findText(current_val)
                if ix >= 0:
                    combo.setCurrentIndex(ix)
            combo.blockSignals(False)

    def go_to_match_page(self):
        self.page_match = self.create_match_page()
        self.stack.addWidget(self.page_match)
        self.stack.setCurrentWidget(self.page_match)

    def create_confirmation_page(self):
        page = QWidget()
        page.setWindowTitle("Подтверждение сопоставления")
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Проверь сопоставление:"))
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QFrame()
        scroll_layout = QGridLayout(scroll_content)
        scroll_layout.setHorizontalSpacing(10)
        scroll_content.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_content)
        items = self.get_sorted_items()
        row = 0
        col = 0
        for index, (name, combobox) in enumerate(items):
            label = short_name_no_ext(os.path.basename(name), 5) if os.path.isabs(name) else short_name_no_ext(name, 5)
            scroll_layout.addWidget(QLabel(f"{label}: {combobox.currentText()}"), row, col)
            col += 1
            if col > 1:
                col = 0
                row += 1
        layout.addWidget(scroll_area)
        button_layout = QHBoxLayout()
        back_button = QPushButton("Назад")
        start_button = QPushButton("Начать")
        start_button.setStyleSheet("background-color: #f47929; color: white;")
        back_button.clicked.connect(self.go_to_match_page)
        start_button.clicked.connect(self.start_copying)
        button_layout.addWidget(back_button)
        button_layout.addWidget(start_button)
        layout.addLayout(button_layout)
        page.setLayout(layout)
        return page

    def go_to_confirmation_page(self):
        self.page_confirmation = self.create_confirmation_page()
        self.stack.addWidget(self.page_confirmation)
        self.stack.setCurrentWidget(self.page_confirmation)

    def get_sorted_items(self):
        if hasattr(self, 'file_to_column') and self.file_to_column:
            return sorted(self.file_to_column.items(), key=lambda x: (x[1].currentText() == "", x[0]))
        else:
            return sorted(self.folder_to_column.items(), key=lambda x: (x[1].currentText() == "", x[0]))

    def create_progress_page(self):
        page = QWidget()
        page.setWindowTitle("Копирование переводов")
        layout = QVBoxLayout()
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setStyleSheet("""
            QProgressBar { text-align: center; }
            QProgressBar::chunk { background-color: #f47929; }
        """)
        layout.addWidget(self.progress_bar)
        page.setLayout(layout)
        return page

    def go_to_progress_page(self):
        self.page_progress = self.create_progress_page()
        self.stack.addWidget(self.page_progress)
        self.stack.setCurrentWidget(self.page_progress)

    def create_completion_page(self):
        page = QWidget()
        page.setWindowTitle("Копирование завершено")
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Файлы успешно сохранены."))
        button_layout = QHBoxLayout()
        close_button = QPushButton("Закрыть приложение")
        restart_button = QPushButton("Вернуться на главный экран")
        close_button.clicked.connect(QApplication.instance().quit)
        restart_button.clicked.connect(self.return_to_main_screen)
        button_layout.addWidget(close_button)
        button_layout.addWidget(restart_button)
        layout.addLayout(button_layout)
        page.setLayout(layout)
        return page

    def go_to_completion_page(self):
        self.page_completion = self.create_completion_page()
        self.stack.addWidget(self.page_completion)
        self.stack.setCurrentWidget(self.page_completion)

    def return_to_main_screen(self):
        self.reset_to_initial_state()
        self.stack.setCurrentWidget(self.page_main)

    def reset_to_initial_state(self):
        self.folder_entry.clear()
        self.excel_file_entry.clear()
        self.copy_column_entry.clear()
        self.skip_first_row_checkbox.setChecked(False)
        self.copy_by_matching_radio.setChecked(True)
        self.sheet_list.clear()
        self.header_row = {}
        self.columns = {}
        self.sheet_to_column = {}
        self.folder_to_column = {}
        self.file_to_column = {}

    def save_mapping_settings(self):
        try:
            mapping = self.get_current_mapping()
            settings_path, _ = QFileDialog.getSaveFileName(self, "Сохранить настройки", '', "JSON файлы (*.json)")
            if settings_path:
                with open(settings_path, 'w', encoding='utf-8') as f:
                    json.dump(mapping, f)
                QMessageBox.information(self, "Успех", "Настройки успешно сохранены.")
        except Exception as e:
            self.log_error(e)
            QMessageBox.critical(self, "Ошибка", "Произошла ошибка при сохранении настроек: " + str(e))

    def get_current_mapping(self):
        if hasattr(self, 'file_to_column') and self.file_to_column:
            return {file: combobox.currentText() for file, combobox in self.file_to_column.items()}
        else:
            return {folder: combobox.currentText() for folder, combobox in self.folder_to_column.items()}

    def load_mapping_settings(self):
        try:
            settings_path, _ = QFileDialog.getOpenFileName(self, "Загрузить настройки", '', "JSON файлы (*.json)")
            if settings_path:
                with open(settings_path, 'r', encoding='utf-8') as f:
                    mapping = json.load(f)
                self.apply_loaded_mapping(mapping)
        except Exception as e:
            self.log_error(e)
            QMessageBox.critical(self, "Ошибка", "Произошла ошибка при загрузке настроек: " + str(e))

    def apply_loaded_mapping(self, mapping):
        if hasattr(self, 'file_to_column') and self.file_to_column:
            self.apply_mapping_to_comboboxes(mapping, self.file_to_column)
        else:
            self.apply_mapping_to_comboboxes(mapping, self.folder_to_column)

    def apply_mapping_to_comboboxes(self, mapping, combobox_dict):
        for name, column_name in mapping.items():
            if name in combobox_dict:
                index = combobox_dict[name].findText(column_name)
                if index != -1:
                    combobox_dict[name].setCurrentIndex(index)

    def start_copying(self):
        try:
            self.go_to_progress_page()
            file_to_column = {k: v.currentText() for k, v in self.file_to_column.items()} if self.file_to_column else {}
            folder_to_column = {k: v.currentText() for k, v in self.folder_to_column.items()} if self.folder_to_column else {}
            if folder_to_column:
                if not self.folder_path or not os.path.isdir(self.folder_path):
                    raise FileNotFoundError("Указанная папка не существует.")
                folder_path = self.folder_path
            else:
                folder_path = ''

            processor = ExcelProcessor(
                main_excel_path=self.excel_file_path,
                folder_path=folder_path,
                copy_column=self.copy_column,
                selected_sheets=self.selected_sheets,
                sheet_to_header_row=self.header_row,
                sheet_to_column={k: v.text() if hasattr(v, "text") else v for k, v in self.sheet_to_column.items()},
                file_to_column=file_to_column,
                folder_to_column=folder_to_column,
                skip_first_row=self.skip_first_row,
                copy_by_row_number=self.copy_by_row_number
            )

            def progress_callback(progress, total):
                if self.progress_bar:
                    self.progress_bar.setMaximum(total)
                    self.progress_bar.setValue(progress)
                    QApplication.processEvents()

            output_file = processor.copy_data(progress_callback=progress_callback)
            self.finalize_copying_process(output_file)

        except Exception as e:
            self.log_error(e)
            QMessageBox.critical(self, "Ошибка", "Произошла ошибка при запуске процесса копирования: " + str(e))

    def finalize_copying_process(self, output_file):
        self.go_to_completion_page()
        QMessageBox.information(self, "Готово", f"Файлы успешно сохранены как {output_file}.")

    def log_error(self, error):
        log_file_path = os.path.join(os.path.dirname(__file__), 'error_log.txt')
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write("Ошибка: " + str(error) + "\n")
            log_file.write(traceback.format_exc())
            log_file.write("\n")

    def center_window(self, window=None):
        if window is None:
            window = self
        screen = QApplication.primaryScreen()
        if screen:
            rect = screen.availableGeometry()
            x = (rect.width() - self.width()) // 2
            y = (rect.height() - self.height()) // 2
            window.move(x, y)
