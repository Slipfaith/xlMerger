import os
import hashlib
from openpyxl import load_workbook
import openpyxl.utils as utils
from openpyxl.styles import PatternFill

class ExcelProcessor:
    def __init__(
        self,
        main_excel_path,
        folder_path,
        copy_column,
        selected_sheets,
        sheet_to_header_row,
        sheet_to_column,
        file_to_column=None,
        folder_to_column=None,
        skip_first_row=False,
        copy_by_row_number=False
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

    # --- Статические методы для получения информации без инициализации объекта ---

    @staticmethod
    def get_sheet_names(main_excel_path):
        wb = load_workbook(main_excel_path, read_only=True)
        sheet_names = wb.sheetnames
        wb.close()
        return sheet_names

    @staticmethod
    def get_sheet_columns(main_excel_path, sheet_name, header_row_index):
        wb = load_workbook(main_excel_path, read_only=True)
        sheet = wb[sheet_name]
        columns = [cell.value for cell in sheet[header_row_index + 1]]
        wb.close()
        return columns

    # --- Основной процесс копирования ---
    def validate_paths_and_column(self):
        if not os.path.exists(self.folder_path):
            raise FileNotFoundError("Указанная папка не существует.")
        if not os.path.exists(self.main_excel_path):
            raise FileNotFoundError("Указанный файл Excel не существует.")
        if not self.copy_column:
            raise ValueError("Укажите столбец для копирования.")

    def copy_data(self, progress_callback=None):
        self.validate_paths_and_column()
        if not self.selected_sheets:
            raise ValueError("Выберите хотя бы один лист.")

        self.workbook = load_workbook(self.main_excel_path)
        for sheet_name in self.selected_sheets:
            header_row_index = self.sheet_to_header_row[sheet_name]
            self.header_row[sheet_name] = header_row_index
            self.columns[sheet_name] = [
                cell.value for cell in self.workbook[sheet_name][header_row_index + 1]
            ]

        # Выбор между file_to_column и folder_to_column (уже только строки!)
        is_file_mapping = bool(self.file_to_column)
        items = self.file_to_column.items() if is_file_mapping else self.folder_to_column.items()
        # Оставляем только непустые строки:
        items = [(k, v) for k, v in items if v]

        total_steps = len(self.selected_sheets) * len(items) if items else 1
        progress = 0

        for sheet_name in self.selected_sheets:
            sheet = self.workbook[sheet_name]
            copy_col_index = utils.column_index_from_string(self.sheet_to_column[sheet_name])
            header_row = self.header_row[sheet_name]

            for name, column_name in items:
                if not column_name:
                    continue
                if column_name not in self.columns[sheet_name]:
                    raise Exception(
                        f"Столбец '{column_name}' не найден на листе '{sheet_name}' основного файла Excel."
                    )

                col_index = self.columns[sheet_name].index(column_name) + 1
                if is_file_mapping:
                    file_path = os.path.join(self.folder_path, name)
                    self._copy_from_file(file_path, sheet_name, copy_col_index, header_row, col_index)
                else:
                    lang_folder_path = os.path.join(self.folder_path, name)
                    self._copy_from_folder(lang_folder_path, sheet_name, copy_col_index, header_row, col_index)
                progress += 1
                if progress_callback:
                    progress_callback(progress, total_steps)

        base, ext = os.path.splitext(self.main_excel_path)
        output_file = f"{base}_out{ext}"
        self.workbook.save(output_file)
        self.workbook.close()
        return output_file

    def _copy_from_file(self, file_path, sheet_name, copy_col_index, header_row, col_index):
        if os.path.isfile(file_path) and file_path.endswith(('.xlsx', '.xls')):
            lang_wb = load_workbook(file_path)
            lang_sheet = lang_wb[sheet_name]
            start_row = 2 if self.skip_first_row else 1
            for row in range(start_row, lang_sheet.max_row + 1):
                if self.copy_by_row_number and row == header_row + 1:
                    continue
                self._copy_cell_value(lang_sheet, sheet_name, row, copy_col_index, header_row, col_index)
            lang_wb.close()

    def _copy_from_folder(self, lang_folder_path, sheet_name, copy_col_index, header_row, col_index):
        for filename in os.listdir(lang_folder_path):
            file_path = os.path.join(lang_folder_path, filename)
            if os.path.isfile(file_path) and filename.endswith(('.xlsx', '.xls')):
                lang_wb = load_workbook(file_path)
                lang_sheet = lang_wb[sheet_name]
                start_row = 2 if self.skip_first_row else 1
                for row in range(start_row, lang_sheet.max_row + 1):
                    if self.copy_by_row_number and row == header_row + 1:
                        continue
                    self._copy_cell_value(lang_sheet, sheet_name, row, copy_col_index, header_row, col_index)
                lang_wb.close()

    def _copy_cell_value(self, lang_sheet, sheet_name, row, copy_col_index, header_row, col_index):
        source_value = lang_sheet.cell(row=row, column=copy_col_index).value
        if source_value is None:
            return

        if self.copy_by_row_number:
            target_row = row
        else:
            start_row_offset = 2 if self.skip_first_row else 1
            target_row = header_row + start_row_offset + (row - start_row_offset)

        target_cell = self.workbook[sheet_name].cell(row=target_row, column=col_index)

        def compute_hash(text):
            if text is None:
                text = ""
            return hashlib.sha256(text.encode('utf-8')).hexdigest()

        source_hash = compute_hash(source_value)
        max_attempts = 5
        for attempt in range(max_attempts):
            target_cell.value = source_value
            if source_value == target_cell.value and compute_hash(target_cell.value) == source_hash:
                return

        fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        target_cell.fill = fill
        # Можно добавить логирование ошибки или выбросить исключение — как нужно для UI
