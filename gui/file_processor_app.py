import os
import json
import traceback
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QStackedWidget, QMessageBox,
    QHBoxLayout, QLabel,
    QPushButton, QFileDialog
)
from PySide6.QtCore import Signal, QSettings
from PySide6.QtGui import QIcon
from utils.i18n import tr, i18n

from gui.main_page import MainPageWidget
from gui.pages.sheet_column_page import SheetColumnPage
from gui.mapping_settings_dialog import MappingSettingsDialog
from gui.pages.confirm_page import ConfirmPage
from gui.pages.progress_page import ProgressPage
from gui.sheet_mapping_dialog import SheetMappingDialog
from core.main_page_logic import MainPageLogic
from core.excel_processor import ExcelProcessor
from gui.pages.header_row_page import HeaderRowPage

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
        self.setWindowIcon(QIcon(r"C:\Users\yanismik\Desktop\PythonProject1\xlM_2.0\xlM2.0.ico"))
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
        self.file_to_sheet_map = {}
        self.skip_first_row = False
        self.copy_by_row_number = False
        self.preserve_formatting = False
        self.progress_bar = None

        self.settings = QSettings('xlMerger', 'xlMerger')
        self.last_mapping_path = self.settings.value('mapping_path', '')

        i18n.language_changed.connect(self.retranslate_ui)
        
        self.setWindowTitle(tr("Обработка файлов"))
        layout = QVBoxLayout()
        layout.addWidget(self.stack)
        self.setLayout(layout)
        self.setGeometry(300, 300, 600, 400)
        self.center_window()
        self.show()
        self.retranslate_ui()

    def process_files(self):
        self.selected_files = self.main_page_logic.selected_files
        self.folder_path = self.main_page_logic.folder_path
        self.copy_column = self.main_page_logic.copy_column.strip()
        self.selected_sheets = self.main_page_logic.get_selected_sheets()
        self.excel_file_path = self.main_page_logic.excel_file_path
        self.skip_first_row = self.page_main.skip_first_row_checkbox.isChecked()
        self.copy_by_row_number = self.page_main.copy_by_row_number_radio.isChecked()

        if not self.copy_column:
            QMessageBox.warning(self, tr("Error"), tr("Укажи столбец для копирования."))
            return  # НЕ переходим дальше

        if not self.check_sheet_mapping():
            return

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
            QMessageBox.critical(self, tr("Error"), tr("Произошла ошибка при загрузке столбцов: {e}").format(e=e))

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
        self.collect_confirmation_changes()
        dialog = MappingSettingsDialog(
            self.folder_path,
            self.selected_files,
            self.selected_sheets,
            self.columns,
            file_to_column=self.file_to_column,
            folder_to_column=self.folder_to_column,
            preserve_formatting=self.preserve_formatting,
            parent=self,
        )
        dialog.saveClicked.connect(self.save_mapping_settings)
        dialog.loadClicked.connect(lambda: self.load_mapping_settings(dialog))
        if dialog.exec():
            file_to_column, folder_to_column, preserve_formatting = dialog.get_mapping()
            self.handle_match_selected(file_to_column, folder_to_column, preserve_formatting)

    def handle_match_selected(self, file_to_column, folder_to_column, preserve_formatting):
        self.file_to_column = file_to_column
        self.folder_to_column = folder_to_column
        self.preserve_formatting = preserve_formatting
        self.go_to_confirmation_page()

    def collect_confirmation_changes(self):
        if hasattr(self, "page_confirmation") and self.page_confirmation:
            mapping = self.page_confirmation.get_current_mapping()
            if self.file_to_column:
                self.file_to_column = mapping
            else:
                self.folder_to_column = mapping
            self.preserve_formatting = self.page_confirmation.is_format_preserved()

    # === Confirmation Page ===
    def get_sorted_items(self):
        # Теперь значения — это строки (названия колонок)
        if self.file_to_column:
            return sorted(self.file_to_column.items(), key=lambda x: (x[1] == "", x[0]))
        else:
            return sorted(self.folder_to_column.items(), key=lambda x: (x[1] == "", x[0]))

    def go_to_confirmation_page(self):
        items = self.get_sorted_items()
        all_columns = [
            col for sheet in self.selected_sheets for col in self.columns[sheet]
            if isinstance(col, str) and col.strip()
        ]
        available = sorted(set(all_columns))
        self.page_confirmation = ConfirmPage(items, available, self.preserve_formatting)
        self.page_confirmation.backClicked.connect(self.go_to_match_page)
        self.page_confirmation.startClicked.connect(self.start_copying)
        self.stack.addWidget(self.page_confirmation)
        self.stack.setCurrentWidget(self.page_confirmation)

    # === Progress Page ===
    def go_to_progress_page(self):
        self.page_progress = ProgressPage()
        self.progress_bar = self.page_progress.progress_bar  # Доступ к бару
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
    def save_mapping_settings(self, mapping):
        try:
            settings_path, _ = QFileDialog.getSaveFileName(
                self,
                tr("Сохранить настройки"),
                self.last_mapping_path,
                tr("JSON файлы (*.json)")
            )
            if settings_path:
                with open(settings_path, 'w', encoding='utf-8') as f:
                    json.dump(mapping, f, ensure_ascii=False, indent=2)
                self.last_mapping_path = settings_path
                self.settings.setValue('mapping_path', settings_path)
                QMessageBox.information(self, tr("Success"), tr("Настройки успешно сохранены."))
        except Exception as e:
            self.log_error(e)
            QMessageBox.critical(self, tr("Error"), tr("Произошла ошибка при сохранении настроек: {e}").format(e=e))

    def load_mapping_settings(self, dialog):
        try:
            settings_path, _ = QFileDialog.getOpenFileName(
                self,
                tr("Загрузить настройки"),
                self.last_mapping_path,
                tr("JSON файлы (*.json)")
            )
            if settings_path:
                with open(settings_path, 'r', encoding='utf-8') as f:
                    mapping = json.load(f)
                dialog.apply_mapping(mapping)
                self.last_mapping_path = settings_path
                self.settings.setValue('mapping_path', settings_path)
        except Exception as e:
            self.log_error(e)
            QMessageBox.critical(self, tr("Error"), tr("Произошла ошибка при загрузке настроек: {e}").format(e=e))

    def check_sheet_mapping(self):
        """Verify sheet names in selected files and ask user to map if needed."""
        files = self.selected_files
        if not files:
            return True

        main_sheets = self.selected_sheets
        file_to_sheets = {}
        mismatched = False
        for f in files:
            try:
                sheets = ExcelProcessor.get_sheet_names(f)
            except Exception as e:
                QMessageBox.critical(self, tr("Error"), str(e))
                return False
            file_to_sheets[f] = sheets
            for ms in main_sheets:
                if ms not in sheets:
                    mismatched = True

        if not mismatched:
            self.file_to_sheet_map = {f: {ms: ms for ms in main_sheets} for f in files}
            return True

        auto_map = {}
        for f, sheets in file_to_sheets.items():
            auto_map[f] = {}
            for ms in main_sheets:
                if ms in sheets:
                    auto_map[f][ms] = ms
                elif len(sheets) == 1:
                    auto_map[f][ms] = sheets[0]

        dialog = SheetMappingDialog(main_sheets, file_to_sheets, auto_map, self)
        if dialog.exec():
            self.file_to_sheet_map = dialog.get_mapping()
            return True
        return False

    # === Копирование ===
    def start_copying(self):
        try:
            self.collect_confirmation_changes()
            self.go_to_progress_page()
            file_to_column = {os.path.basename(k): v for k, v in self.file_to_column.items()} if self.file_to_column else {}
            folder_to_column = {os.path.basename(k): v for k, v in self.folder_to_column.items()} if self.folder_to_column else {}
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
                file_to_sheet_map=self.file_to_sheet_map,
                skip_first_row=self.skip_first_row,
                copy_by_row_number=self.copy_by_row_number,
                preserve_formatting=self.preserve_formatting
            )

            def progress_callback(progress, total):
                if hasattr(self, 'progress_bar') and self.progress_bar:
                    self.progress_bar.setMaximum(total)
                    self.progress_bar.setValue(progress)
                    QApplication.processEvents()

            output_file = processor.copy_data(progress_callback=progress_callback)
            self.finalize_copying_process(output_file)

        except Exception as e:
            self.log_error(e)
            QMessageBox.critical(self, tr("Error"), tr("Произошла ошибка при запуске процесса копирования: {e}").format(e=e))

    def finalize_copying_process(self, output_file):
        self.go_to_completion_page()
        QMessageBox.information(self, tr("Success"), tr("Файлы успешно сохранены как {output_file}.").format(output_file=output_file))

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

    def retranslate_ui(self):
        self.setWindowTitle(tr("Обработка файлов"))
