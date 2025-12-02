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
        self.excel_file_paths = []
        self.sheet_names = []
        self.selected_sheets = []
        self.copy_column = ''

        # Подключаем сигналы UI к логике
        self.ui.folderSelected.connect(self.on_folder_selected)
        self.ui.filesSelected.connect(self.on_files_selected)
        self.ui.excelFileSelected.connect(self.on_excel_file_selected)
        self.ui.excelFilesSelected.connect(self.on_excel_files_selected)
        self.ui.processTriggered.connect(self.on_process_clicked)
        self.ui.previewTriggered.connect(self.on_preview_clicked)

        self.ui.sheet_list.clear()

    def on_folder_selected(self, path):
        self.folder_path = path
        self.selected_files = []  # Очищаем только список выбранных файлов
        self.ui.folder_entry.setText(path)
        # Больше ничего не сбрасываем!

    def on_files_selected(self, files):
        self.selected_files = files
        self.folder_path = os.path.dirname(files[0]) if files else ''
        self.ui.folder_entry.setText(self.folder_path)
        # Больше ничего не сбрасываем!

    def on_excel_file_selected(self, file):
        self.excel_file_path = file
        self.excel_file_paths = [file] if file else []
        self.ui.excel_file_entry.setText(file)
        self.load_sheet_names()

    def on_excel_files_selected(self, files):
        self.excel_file_paths = list(files)
        self.excel_file_path = self.excel_file_paths[0] if self.excel_file_paths else ''
        if len(self.excel_file_paths) == 1:
            self.ui.excel_file_entry.setText(files[0])
        else:
            self.ui.excel_file_entry.setText(tr("Выбрано файлов: {count}").format(count=len(self.excel_file_paths)))
        self.load_sheet_names()

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
        if not self.folder_path and not self.selected_files:
            QMessageBox.warning(self.ui, "Предупреждение", "Сначала выбери папку или эксели с переводами.")
            return
        dialog = ExcelFileSelector(self.folder_path, self.selected_files)
        if dialog.exec():
            selected_file = dialog.selected_file
            if selected_file:
                self.preview_window = ExcelPreviewer(selected_file)
                self.preview_window.show()

    def on_process_clicked(self):
        self.folder_path = self.ui.folder_entry.text()
        self.copy_column = self.ui.copy_column_entry.text()
        self.selected_sheets = self.get_selected_sheets()
        if self.excel_file_paths:
            self.excel_file_path = self.excel_file_paths[0]
        else:
            self.excel_file_path = self.ui.excel_file_entry.text()

        if not self.selected_files and self.folder_path and os.path.isdir(self.folder_path):
            self.selected_files = [
                os.path.join(self.folder_path, fname)
                for fname in os.listdir(self.folder_path)
                if fname.lower().endswith(('.xlsx', '.xls'))
            ]
        if not self.validate_inputs():
            return
        self.proceed_to_next.emit()  # Сигнализируем, что всё готово — пора переходить!

    def validate_inputs(self):
        folder_path = self.ui.folder_entry.text()
        excel_file_path = self.excel_file_paths[0] if self.excel_file_paths else self.ui.excel_file_entry.text()
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
        target_files = self.excel_file_paths or [excel_file_path]
        missing = [f for f in target_files if not os.path.isfile(f)]
        if missing:
            QMessageBox.critical(self.ui, "Ошибка", tr("Указанный файл Excel не существует.") + f"\n{missing[0]}")
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