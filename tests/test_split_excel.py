import os
from openpyxl import Workbook, load_workbook
from core.split_excel import split_excel_by_languages


def test_split_excel(tmp_path):
    src = tmp_path / "main.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["ru", "de", "en", "comment"])
    ws.append(["привет", "hallo", "hello", "c1"])
    wb.save(src)
    wb.close()

    split_excel_by_languages(str(src), "Sheet1", "ru")

    out_de = tmp_path / "ru-de_main.xlsx"
    out_en = tmp_path / "ru-en_main.xlsx"

    assert out_de.is_file()
    assert out_en.is_file()

    wb_de = load_workbook(out_de)
    ws_de = wb_de.active
    assert ws_de.cell(row=1, column=1).value == "ru"
    assert ws_de.cell(row=1, column=2).value == "de"
    assert ws_de.cell(row=2, column=1).value == "привет"
    assert ws_de.cell(row=2, column=2).value == "hallo"
    wb_de.close()
    wb_en = load_workbook(out_en)
    ws_en = wb_en.active
    assert ws_en.cell(row=1, column=1).value == "ru"
    assert ws_en.cell(row=1, column=2).value == "en"
    assert ws_en.cell(row=2, column=1).value == "привет"
    assert ws_en.cell(row=2, column=2).value == "hello"
    wb_en.close()
