import os
import json
import traceback
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QStackedWidget, QMessageBox,
    QProgressBar, QScrollArea, QFrame, QHBoxLayout, QGridLayout, QLabel,
    QPushButton, QComboBox, QLineEdit, QFileDialog
)
from PySide6.QtCore import Qt, Signal

from gui.main_page import MainPageWidget
from gui.pages.sheet_column_page import SheetColumnPage
from gui.pages.match_page import MatchPage
from core.main_page_logic import MainPageLogic
from core.excel_processor import ExcelProcessor
from gui.pages.header_row_page import HeaderRowPage  # подключаем новую страницу

def short_name_no_ext(name, n=5):
    base, ext = os.path.splitext(name)
    if len(base) <= 2 * n:
        return base
    return f"{base[:n]}...{base[-n:]}"


class FileProcessorApp(QWidget):
    copyingStarted = Signal()

    def __init__(self):
        super().__init__()
        self.stack = QStackedWidget(self)
        # --- Главная страница и логика ---
        self.page_main = MainPageWidget()
        self.main_page_logic = MainPageLogic(self.page_main)
        self.stack.addWidget(self.page_main)
        self.page_main.processTriggered.connect(self.process_files)
        self.page_main.previewTriggered.connect(self.preview_excel)

        # Переменные этапов
        self.selected_files = []
        self.folder_path = ''
        self.copy_column = ''
        self.selected_sheets = []
        self.excel_file_path = ''
        self.header_row = {}
        self.columns = {}
        self.sheet_to_column = {}
        self.folder_to_column = {}
        self.file_to_column = {}
        self.skip_first_row = False
        self.copy_by_row_number = False
        self.progress_bar = None

        self.setWindowTitle("Обработка файлов")
        layout = QVBoxLayout()
        layout.addWidget(self.stack)
        self.setLayout(layout)
        self.center_window()
        self.setGeometry(300, 300, 600, 400)
        self.show()

    def process_files(self):
        self.selected_files = self.main_page_logic.selected_files
        self.folder_path = self.main_page_logic.folder_path
        self.copy_column = self.main_page_logic.copy_column
        self.selected_sheets = self.main_page_logic.get_selected_sheets()
        self.excel_file_path = self.main_page_logic.excel_file_path
        self.skip_first_row = self.page_main.skip_first_row_checkbox.isChecked()
        self.copy_by_row_number = self.page_main.copy_by_row_number_radio.isChecked()
        self.go_to_header_page()

    def preview_excel(self):
        self.main_page_logic.on_preview_clicked()

    # === Header Page ===
    def go_to_header_page(self):
        self.page_header = HeaderRowPage(self.selected_sheets)
        self.page_header.backClicked.connect(self.go_to_main_page)
        self.page_header.nextClicked.connect(self.handle_header_row_selected)
        self.stack.addWidget(self.page_header)
        self.stack.setCurrentWidget(self.page_header)

    def handle_header_row_selected(self, sheet_to_header_row):
        self.header_row = sheet_to_header_row
        self.load_columns_and_go_to_sheet_column()

    def go_to_main_page(self):
        self.stack.setCurrentWidget(self.page_main)

    # === Sheet-Column Page ===
    def load_columns_and_go_to_sheet_column(self):
        try:
            self.columns = {}
            for sheet, header_row in self.header_row.items():
                self.columns[sheet] = ExcelProcessor.get_sheet_columns(self.excel_file_path, sheet, header_row)
            self.go_to_sheet_column_page()
        except Exception as e:
            self.log_error(e)
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при загрузке столбцов: {e}")

    def go_to_sheet_column_page(self):
        self.page_sheet_column = SheetColumnPage(self.selected_sheets, self.copy_column)
        self.page_sheet_column.backClicked.connect(self.go_to_header_page)
        self.page_sheet_column.nextClicked.connect(self.handle_sheet_column_selected)
        self.stack.addWidget(self.page_sheet_column)
        self.stack.setCurrentWidget(self.page_sheet_column)

    def handle_sheet_column_selected(self, sheet_to_column):
        self.sheet_to_column = sheet_to_column
        self.go_to_match_page()

    # === Match Page (file/papka -> column mapping) ===
    def go_to_match_page(self):
        self.page_match = MatchPage(
            self.folder_path,
            self.selected_files,
            self.selected_sheets,
            self.columns,
            self.file_to_column,
            self.folder_to_column,
        )
        self.page_match.backClicked.connect(self.go_to_sheet_column_page)
        self.page_match.nextClicked.connect(self.handle_match_selected)
        self.page_match.saveClicked.connect(self.save_mapping_settings)
        self.page_match.loadClicked.connect(self.load_mapping_settings)
        self.stack.addWidget(self.page_match)
        self.stack.setCurrentWidget(self.page_match)

    def handle_match_selected(self, file_to_column, folder_to_column):
        self.file_to_column = file_to_column
        self.folder_to_column = folder_to_column
        self.go_to_confirmation_page()

    # === Confirmation Page ===
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
        # Теперь значения — это строки (названия колонок), а не QComboBox!
        if self.file_to_column:
            return sorted(self.file_to_column.items(), key=lambda x: (x[1] == "", x[0]))
        else:
            return sorted(self.folder_to_column.items(), key=lambda x: (x[1] == "", x[0]))

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
        for index, (name, column_name) in enumerate(items):
            label = short_name_no_ext(os.path.basename(name), 5) if os.path.isabs(name) else short_name_no_ext(name, 5)
            scroll_layout.addWidget(QLabel(f"{label}: {column_name}"), row, col)
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

    # === Progress Page ===
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

    # === Completion Page ===
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
        self.stack.setCurrentWidget(self.page_main)

    # === Save/Load Mapping ===
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
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при сохранении настроек: {e}")

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
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при загрузке настроек: {e}")

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

    # === Копирование ===
    def start_copying(self):
        try:
            self.go_to_progress_page()
            # Здесь просто копируем словари, в которых уже строки, а не QComboBox
            file_to_column = dict(self.file_to_column) if self.file_to_column else {}
            folder_to_column = dict(self.folder_to_column) if self.folder_to_column else {}
            folder_path = self.folder_path if folder_to_column else ''

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
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при запуске процесса копирования: {e}")

            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка при запуске процесса копирования: {e}")

    def finalize_copying_process(self, output_file):
        self.go_to_completion_page()
        QMessageBox.information(self, "Готово", f"Файлы успешно сохранены как {output_file}.")

    # === Error Logging & Center Window ===
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
