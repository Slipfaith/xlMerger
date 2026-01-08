import os
import traceback
from PySide6.QtWidgets import QMessageBox, QListWidgetItem
from PySide6.QtCore import Qt, QObject, Signal

from core.excel_processor import ExcelProcessor
from gui.excel_previewer import ExcelPreviewer
from gui.excel_file_selector import ExcelFileSelector

class MainPageLogic(QObject):
    proceed_to_next = Signal()  # Сигнал для перехода на следующий шаг

    def __init__(self, main_page_widget):
        super().__init__()
        self.ui = main_page_widget

        # Состояние главной страницы
        self.folder_path = ''
        self.selected_files = []
        self.excel_file_path = ''
        self.sheet_names = []
        self.selected_sheets = []
        self.copy_column = ''
        self.preview_window = None

        # Подключаем сигналы UI к логике
        self.ui.folderSelected.connect(self.on_folder_selected)
        self.ui.filesSelected.connect(self.on_files_selected)
        self.ui.excelFileSelected.connect(self.on_excel_file_selected)
        self.ui.previewTriggered.connect(self.on_preview_clicked)
        self.ui.copy_column_entry.textChanged.connect(self.on_copy_column_changed)

        self.ui.sheet_list.clear()
        self.update_process_button_state()

    def on_folder_selected(self, path):
        self.folder_path = path
        self.selected_files = []  # Очищаем только список выбранных файлов
        self.ui.folder_entry.setText(path)
        self.update_process_button_state()
        # Больше ничего не сбрасываем!

    def on_files_selected(self, files):
        self.selected_files = files
        self.folder_path = os.path.dirname(files[0]) if files else ''
        self.ui.folder_entry.setText(self.folder_path)
        self.update_process_button_state()
        # Больше ничего не сбрасываем!

    def on_excel_file_selected(self, file):
        self.excel_file_path = file
        self.ui.excel_file_entry.setText(file)
        self.load_sheet_names()
        self.update_process_button_state()

    def on_copy_column_changed(self, text):
        self.copy_column = text
        self.update_process_button_state()

    def load_sheet_names(self):
        try:
            self.sheet_names = ExcelProcessor.get_sheet_names(self.excel_file_path)
            self.populate_sheet_list()
        except Exception as e:
            self.log_error(e)
            QMessageBox.critical(self.ui, "Ошибка", f"Не удалось загрузить листы из файла Excel: {e}")

    def populate_sheet_list(self):
        self.ui.sheet_list.clear()
        for sheet_name in self.sheet_names:
            item = QListWidgetItem(sheet_name)
            item.setCheckState(Qt.Checked)
            self.ui.sheet_list.addItem(item)

    def deselect_all_sheets(self):
        for index in range(self.ui.sheet_list.count()):
            self.ui.sheet_list.item(index).setCheckState(Qt.Unchecked)

    def select_all_sheets(self):
        for index in range(self.ui.sheet_list.count()):
            self.ui.sheet_list.item(index).setCheckState(Qt.Checked)

    def get_selected_sheets(self):
        return [
            self.ui.sheet_list.item(i).text()
            for i in range(self.ui.sheet_list.count())
            if self.ui.sheet_list.item(i).checkState() == Qt.Checked
        ]

    def on_preview_clicked(self):
        available_files = self._collect_source_files()
        if not available_files:
            QMessageBox.warning(self.ui, "Предупреждение", "Сначала выбери папку или эксели с переводами.")
            return
        dialog = ExcelFileSelector(
            self.folder_path,
            selected_files=self.selected_files,
            target_excel=self.excel_file_path,
        )
        if dialog.exec():
            selected_file = dialog.selected_file
            if selected_file:
                if self.preview_window and self.preview_window.isVisible():
                    self.preview_window.load_file(selected_file)
                    self.preview_window.activateWindow()
                    self.preview_window.raise_()
                else:
                    preview_window = ExcelPreviewer(selected_file)
                    if preview_window is None:
                        QMessageBox.critical(self.ui, "Ошибка", "Не удалось открыть окно превью.")
                        return
                    self.preview_window = preview_window
                    preview_window.destroyed.connect(self._clear_preview_window)
                    self.preview_window.show()

    def _clear_preview_window(self, destroyed_obj=None):
        if destroyed_obj is None or destroyed_obj == self.preview_window:
            self.preview_window = None

    def update_process_button_state(self):
        source_files = self._collect_source_files()
        copy_column = self.ui.copy_column_entry.text().strip()
        target_excel = self.ui.excel_file_entry.text().strip()
        target_ok = bool(target_excel) and os.path.isfile(target_excel)
        can_start = bool(source_files) and bool(copy_column) and target_ok
        self.ui.process_button.setEnabled(can_start)

    def _collect_source_files(self):
        if self.selected_files:
            return [
                f for f in self.selected_files
                if os.path.isfile(f) and f.lower().endswith(('.xlsx', '.xls'))
            ]
        if self.folder_path and os.path.isdir(self.folder_path):
            return [
                os.path.join(root, fname)
                for root, _, files in os.walk(self.folder_path)
                for fname in files
                if fname.lower().endswith(('.xlsx', '.xls'))
            ]
        return []

    def validate_inputs(self):
        folder_path = self.ui.folder_entry.text()
        excel_file_path = self.ui.excel_file_entry.text()
        copy_column = self.ui.copy_column_entry.text()
        selected_sheets = self.get_selected_sheets()

        if self.selected_files:
            for f in self.selected_files:
                if not os.path.isfile(f):
                    QMessageBox.critical(self.ui, "Ошибка", f"Файл Excel не найден: {f}")
                    return False
        elif folder_path and not os.path.isdir(folder_path):
            QMessageBox.critical(self.ui, "Ошибка", "Указанная папка не существует.")
            return False
        elif not self.selected_files and folder_path and os.path.isdir(folder_path):
            self.selected_files = [
                os.path.join(root, fname)
                for root, _, files in os.walk(folder_path)
                for fname in files
                if fname.lower().endswith(('.xlsx', '.xls'))
            ]
        if not excel_file_path or not os.path.isfile(excel_file_path):
            QMessageBox.critical(self.ui, "Ошибка", "Указанный файл Excel не существует.")
            return False
        if not copy_column:
            QMessageBox.critical(self.ui, "Ошибка", "Укажи столбец для копирования.")
            return False
        if not selected_sheets:
            QMessageBox.critical(self.ui, "Ошибка", "Выбери хотя бы один лист.")
            return False
        return True

    def log_error(self, error):
        log_file_path = os.path.join(os.path.dirname(__file__), 'error_log.txt')
        with open(log_file_path, 'a', encoding='utf-8') as log_file:
            log_file.write("Ошибка: " + str(error) + "\n")
            log_file.write(traceback.format_exc())
            log_file.write("\n")
