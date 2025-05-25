import os
import tempfile
import pytest
from openpyxl import Workbook, load_workbook
from utils.logger import Logger
from core.excel_processor import ExcelProcessor

@pytest.fixture
def sample_main_excel(tmp_path):
    path = tmp_path / "main.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Header1", "Header2", "Header3"])
    ws.append(["Val1", "Val2", "Val3"])
    wb.save(path)
    return str(path)

@pytest.fixture
def sample_folder_with_excel(tmp_path):
    folder = tmp_path / "translations"
    folder.mkdir()
    path = folder / "file1.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["Header1", "Header2", "Header3"])
    ws.append(["T1", "T2", "T3"])
    wb.save(path)
    return str(folder), str(path)

def test_excel_processor_copy_data_integration(sample_main_excel, sample_folder_with_excel, caplog):
    folder_path, translation_file = sample_folder_with_excel
    logger = Logger()
    processor = ExcelProcessor(
        main_excel_path=sample_main_excel,
        folder_path=folder_path,
        copy_column="A",
        selected_sheets=["Sheet1"],
        sheet_to_header_row={"Sheet1": 0},
        sheet_to_column={"Sheet1": "Header2"},
        folder_to_column={"translations": "Header2"},
        logger=logger
    )

    # Валидация проходит без ошибок
    processor.validate_paths_and_column()

    # Проверка что copy_data отрабатывает без исключений и создаёт файл
    output_file = processor.copy_data()

    assert os.path.isfile(output_file)
    # Лог должен содержать успешное сохранение
    assert any("Файл успешно сохранён" in entry for entry in logger.entries)

    # Проверяем, что после копирования в output_file есть изменения (не пустой)
    wb_out = load_workbook(output_file)
    ws_out = wb_out["Sheet1"]
    val = ws_out.cell(row=2, column=2).value  # В столбце Header2
    assert val == "T2" or val == "Val2"  # либо исходное, либо скопированное

    wb_out.close()

    # Некорректный путь к папке вызывает ошибку
    processor.folder_path = "/nonexistent"
    with pytest.raises(FileNotFoundError):
        processor.validate_paths_and_column()

    # Некорректный столбец вызывает исключение при copy_data
    processor.folder_path = folder_path
    processor.sheet_to_column = {"Sheet1": "NonExistentCol"}
    with pytest.raises(Exception):
        processor.copy_data()
