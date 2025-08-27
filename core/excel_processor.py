import os
import shutil
from openpyxl import load_workbook
import openpyxl.utils as utils

from utils.logger import Logger

class ExcelProcessor:
    def __init__(
        self, main_excel_path, folder_path, copy_column, selected_sheets,
        sheet_to_header_row, sheet_to_column, file_to_column=None, folder_to_column=None,
        file_to_sheet_map=None, skip_first_row=False, copy_by_row_number=False,
        preserve_formatting=False, logger=None
    ):
        self.main_excel_path = main_excel_path
        self.folder_path = folder_path
        self.copy_column = copy_column
        self.selected_sheets = selected_sheets
        self.sheet_to_header_row = sheet_to_header_row
        self.sheet_to_column = sheet_to_column
        self.file_to_column = file_to_column or {}
        self.folder_to_column = folder_to_column or {}
        self.file_to_sheet_map = file_to_sheet_map or {}
        self.skip_first_row = skip_first_row
        self.copy_by_row_number = copy_by_row_number
        self.preserve_formatting = preserve_formatting

        self.workbook = None
        self.columns = {}
        self.header_row = {}
        self.output_file = None

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

        base, ext = os.path.splitext(self.main_excel_path)
        self.output_file = f"{base}_out{ext}"
        shutil.copyfile(self.main_excel_path, self.output_file)

        workbook = load_workbook(self.main_excel_path, read_only=True)
        self.logger.log_info(f"Загружен основной Excel: {self.main_excel_path}")
        for sheet_name in self.selected_sheets:
            header_row_index = self.sheet_to_header_row[sheet_name]
            self.header_row[sheet_name] = header_row_index
            self.columns[sheet_name] = [
                cell.value for cell in workbook[sheet_name][header_row_index + 1]
            ]
            self.logger.log_info(f"Обрабатывается лист: {sheet_name}")
        workbook.close()

        is_file_mapping = bool(self.file_to_column)
        items = self.file_to_column.items() if is_file_mapping else self.folder_to_column.items()
        total_steps = len(self.selected_sheets) * len([c for _, c in items if c]) if items else 1
        progress = 0

        for sheet_name in self.selected_sheets:
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
                    self._copy_from_file(file_path, sheet_name, copy_col_index, header_row, col_index)
                else:
                    lang_folder_path = os.path.join(self.folder_path, name)
                    self.logger.log_info(f"Копирование из папки: {lang_folder_path}, лист: {sheet_name}, столбец: {column_name}")
                    self._copy_from_folder(lang_folder_path, sheet_name, copy_col_index, header_row, col_index)
                progress += 1
                if progress_callback:
                    progress_callback(progress, total_steps)

        self.logger.log_info(f"Файл успешно сохранён: {self.output_file}")
        self.logger.save()
        return self.output_file

    def _find_matching_sheet(self, lang_wb, main_sheet_name, file_path=None):
        if file_path:
            mapping = self.file_to_sheet_map.get(file_path, {}).get(main_sheet_name)
            if mapping and mapping in lang_wb.sheetnames:
                return mapping
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
            lang_wb = load_workbook(file_path, read_only=True)
            target_sheet_name = self._find_matching_sheet(lang_wb, main_sheet_name, file_path)
            lang_wb.close()
            self._copy_from_sheet(file_path, target_sheet_name, main_sheet_name, copy_col_index, header_row, col_index)

    def _copy_from_folder(self, lang_folder_path, main_sheet_name, copy_col_index, header_row, col_index):
        for filename in os.listdir(lang_folder_path):
            file_path = os.path.join(lang_folder_path, filename)
            if os.path.isfile(file_path) and filename.endswith(('.xlsx', '.xls')):
                lang_wb = load_workbook(file_path, read_only=True)
                target_sheet_name = self._find_matching_sheet(lang_wb, main_sheet_name, file_path)
                lang_wb.close()
                self._copy_from_sheet(file_path, target_sheet_name, main_sheet_name, copy_col_index, header_row, col_index)

    def _copy_from_sheet(self, lang_file_path, lang_sheet_name, sheet_name, copy_col_index, header_row, col_index):
        try:
            import xlwings as xw
        except Exception as e:
            self.logger.log_error("xlwings недоступен", "", "", str(e))
            return

        app = main_wb = lang_wb = None
        try:
            app = xw.App(visible=False, add_book=False)
            dest_path = os.path.abspath(self.output_file or self.main_excel_path)
            main_wb = app.books.open(dest_path)
            lang_wb = app.books.open(os.path.abspath(lang_file_path))
            source_sheet = lang_wb.sheets[lang_sheet_name]
            target_sheet = main_wb.sheets[sheet_name]

            src_col_letter = utils.get_column_letter(copy_col_index)
            dst_col_letter = utils.get_column_letter(col_index)
            last_row = source_sheet.range((source_sheet.cells.last_cell.row, copy_col_index)).end('up').row

            def copy_range(r1, r2, dest_start):
                if r2 < r1:
                    return
                src_range = source_sheet.range(f"{src_col_letter}{r1}:{src_col_letter}{r2}")
                dst_range = target_sheet.range(f"{dst_col_letter}{dest_start}:{dst_col_letter}{dest_start + (r2 - r1)}")
                src_range.api.Copy(dst_range.api)
                values = src_range.value
                if not isinstance(values, list):
                    values = [values]
                for idx, val in enumerate(values, start=0):
                    if isinstance(val, list):
                        val = val[0]
                    if val is None or (isinstance(val, str) and val.strip() == ""):
                        continue
                    self.logger.log_copy(sheet_name, dest_start + idx, col_index, val)

            if header_row > 0:
                copy_range(1, header_row, 1 if self.copy_by_row_number else header_row + 1)

            start_row = header_row + 2
            copy_range(start_row, last_row, start_row if self.copy_by_row_number else start_row + header_row)

            main_wb.save()
        except Exception as e:
            self.logger.log_error("Ошибка при копировании через Excel", "", "", f"{lang_file_path}: {e}")
        finally:
            if lang_wb:
                lang_wb.close()
            if main_wb:
                main_wb.close()
            if app:
                app.quit()
