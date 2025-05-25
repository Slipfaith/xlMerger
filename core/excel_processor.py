import os
import hashlib
from openpyxl import load_workbook
import openpyxl.utils as utils
from openpyxl.styles import PatternFill

from utils.logger import Logger

class ExcelProcessor:
    def __init__(
        self, main_excel_path, folder_path, copy_column, selected_sheets,
        sheet_to_header_row, sheet_to_column, file_to_column=None, folder_to_column=None,
        skip_first_row=False, copy_by_row_number=False, logger=None
    ):
        self.main_excel_path = main_excel_path
        self.folder_path = folder_path
        self.copy_column = copy_column
        self.selected_sheets = selected_sheets
        self.sheet_to_header_row = sheet_to_header_row
        self.sheet_to_column = sheet_to_column
        self.file_to_column = file_to_column or {}
        self.folder_to_column = folder_to_column or {}
        self.skip_first_row = skip_first_row
        self.copy_by_row_number = copy_by_row_number

        self.workbook = None
        self.columns = {}
        self.header_row = {}

        self.logger = logger or Logger()

    @staticmethod
    def get_sheet_names(excel_path):
        wb = load_workbook(excel_path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names

    @staticmethod
    def get_sheet_columns(excel_path, sheet_name, header_row_index):
        wb = load_workbook(excel_path, read_only=True)
        sheet = wb[sheet_name]
        cols = [cell.value for cell in sheet[header_row_index + 1]]
        wb.close()
        return cols

    def validate_paths_and_column(self):
        if self.folder_path and not os.path.isdir(self.folder_path):
            self.logger.log_error("Папка перевода не найдена", "", "", self.folder_path)
            raise FileNotFoundError("Указанная папка не существует.")
        if not os.path.exists(self.main_excel_path):
            self.logger.log_error("Файл Excel не найден", "", "", self.main_excel_path)
            raise FileNotFoundError("Указанный файл Excel не существует.")
        if not self.copy_column:
            self.logger.log_error("Не указан столбец для копирования", "", "", "")
            raise ValueError("Укажите столбец для копирования.")

    def copy_data(self, progress_callback=None):
        self.validate_paths_and_column()
        if not self.selected_sheets:
            self.logger.log_error("Не выбраны листы", "", "", "")
            raise ValueError("Выберите хотя бы один лист.")

        self.workbook = load_workbook(self.main_excel_path)
        self.logger.log_info(f"Загружен основной Excel: {self.main_excel_path}")

        for sheet_name in self.selected_sheets:
            header_row_index = self.sheet_to_header_row[sheet_name]
            self.header_row[sheet_name] = header_row_index
            self.columns[sheet_name] = [
                cell.value for cell in self.workbook[sheet_name][header_row_index + 1]
            ]
            self.logger.log_info(f"Обрабатывается лист: {sheet_name}")

        is_file_mapping = bool(self.file_to_column)
        items = self.file_to_column.items() if is_file_mapping else self.folder_to_column.items()
        total_steps = len(self.selected_sheets) * len([c for _, c in items if c]) if items else 1
        progress = 0

        for sheet_name in self.selected_sheets:
            sheet = self.workbook[sheet_name]
            copy_col_index = utils.column_index_from_string(self.sheet_to_column[sheet_name])
            header_row = self.header_row[sheet_name]

            for name, column_name in items:
                if not column_name:
                    continue
                if column_name not in self.columns[sheet_name]:
                    self.logger.log_error(f"Столбец '{column_name}' не найден на листе '{sheet_name}'", "", "", name)
                    raise Exception(
                        f"Столбец '{column_name}' не найден на листе '{sheet_name}' основного файла Excel."
                    )

                col_index = self.columns[sheet_name].index(column_name) + 1
                if is_file_mapping:
                    file_path = os.path.join(self.folder_path, name)
                    self.logger.log_info(f"Копирование из файла: {file_path}, лист: {sheet_name}, столбец: {column_name}")
                    self._copy_from_file(
                        file_path, sheet_name, copy_col_index, header_row, col_index
                    )
                else:
                    lang_folder_path = os.path.join(self.folder_path, name)
                    self.logger.log_info(f"Копирование из папки: {lang_folder_path}, лист: {sheet_name}, столбец: {column_name}")
                    self._copy_from_folder(
                        lang_folder_path, sheet_name, copy_col_index, header_row, col_index
                    )
                progress += 1
                if progress_callback:
                    progress_callback(progress, total_steps)

        base, ext = os.path.splitext(self.main_excel_path)
        output_file = f"{base}_out{ext}"
        self.workbook.save(output_file)
        self.logger.log_info(f"Файл успешно сохранён: {output_file}")
        self.logger.save()
        self.workbook.close()
        return output_file

    def _find_matching_sheet(self, lang_wb, main_sheet_name):
        if main_sheet_name in lang_wb.sheetnames:
            return main_sheet_name
        elif len(lang_wb.sheetnames) == 1:
            return lang_wb.sheetnames[0]
        else:
            sheets = ', '.join(lang_wb.sheetnames)
            self.logger.log_error(f"Не найден лист '{main_sheet_name}'", "", "", sheets)
            raise Exception(
                f"Не найден лист '{main_sheet_name}' в файле перевода. "
                f"В файле листы: {sheets}. "
                f"Переименуйте листы для автоматического сопоставления, либо удалите лишние листы."
            )

    def _copy_from_file(self, file_path, main_sheet_name, copy_col_index, header_row, col_index):
        if os.path.isfile(file_path) and file_path.endswith(('.xlsx', '.xls')):
            lang_wb = load_workbook(file_path)
            target_sheet_name = self._find_matching_sheet(lang_wb, main_sheet_name)
            lang_sheet = lang_wb[target_sheet_name]
            self._copy_from_sheet(lang_sheet, main_sheet_name, copy_col_index, header_row, col_index)
            lang_wb.close()

    def _copy_from_folder(self, lang_folder_path, main_sheet_name, copy_col_index, header_row, col_index):
        for filename in os.listdir(lang_folder_path):
            file_path = os.path.join(lang_folder_path, filename)
            if os.path.isfile(file_path) and filename.endswith(('.xlsx', '.xls')):
                lang_wb = load_workbook(file_path)
                target_sheet_name = self._find_matching_sheet(lang_wb, main_sheet_name)
                lang_sheet = lang_wb[target_sheet_name]
                self._copy_from_sheet(lang_sheet, main_sheet_name, copy_col_index, header_row, col_index)
                lang_wb.close()

    def _copy_from_sheet(self, lang_sheet, sheet_name, copy_col_index, header_row, col_index):
        for row in range(1, lang_sheet.max_row + 1):
            if row == header_row + 1:
                continue  # Всегда пропускаем строку-заголовок (например, 'RU' или что там)
            source_value = lang_sheet.cell(row=row, column=copy_col_index).value
            if source_value is None or (isinstance(source_value, str) and source_value.strip() == ""):
                continue
            if self.copy_by_row_number:
                target_row = row
            else:
                offset = 2 if self.skip_first_row else 1
                target_row = header_row + offset + (row - offset)
            self._set_cell(sheet_name, target_row, col_index, source_value)

    def _set_cell(self, sheet_name, target_row, col_index, value):
        target_cell = self.workbook[sheet_name].cell(row=target_row, column=col_index)
        def compute_hash(text):
            if text is None:
                text = ""
            return hashlib.sha256(str(text).encode('utf-8')).hexdigest()
        source_hash = compute_hash(value)
        max_attempts = 5
        for attempt in range(max_attempts):
            target_cell.value = value
            if value == target_cell.value and compute_hash(target_cell.value) == source_hash:
                # Успешно скопировано — логируем
                self.logger.log_copy(sheet_name, target_row, col_index, value)
                break
        else:
            fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
            target_cell.fill = fill
            self.logger.log_error(f"Не удалось записать значение после {max_attempts} попыток", "", "", f"{sheet_name}: R{target_row} C{col_index} [{value}]")