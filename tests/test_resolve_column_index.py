from core.excel_processor import ExcelProcessor
from utils.logger import Logger
import os


def make_processor():
    logger = Logger(log_file=os.devnull)
    processor = ExcelProcessor(
        main_excel_path='dummy.xlsx',
        folder_path='',
        copy_column='',
        selected_sheets=['ALL'],
        sheet_to_header_row={'ALL': 0},
        sheet_to_column={'ALL': 'PTBR'},
        logger=logger,
    )
    processor.columns['ALL'] = ['Key', 'PTBR', 'EN']
    return processor


def test_resolve_column_index_by_header():
    proc = make_processor()
    assert proc._resolve_column_index('ALL', 'PTBR') == 2


def test_resolve_column_index_by_letter():
    proc = make_processor()
    assert proc._resolve_column_index('ALL', 'B') == 2
