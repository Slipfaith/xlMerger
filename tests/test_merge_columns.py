from openpyxl import Workbook, load_workbook
from core.merge_columns import merge_excel_columns


def create_wb(path, data, sheet_name="Sheet1"):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    for col, values in data.items():
        for idx, val in enumerate(values, start=1):
            ws[f"{col}{idx}"] = val
    wb.save(path)
    wb.close()


def test_merge_excel_columns(tmp_path):
    main_path = tmp_path / "main.xlsx"
    src_path = tmp_path / "src.xlsx"

    create_wb(main_path, {"A": ["h1", "h2", "h3"]}, sheet_name="Main")
    create_wb(src_path, {"A": ["x", "y", "z"]})

    mappings = [{
        "source": str(src_path),
        "source_columns": ["A"],
        "target_sheet": "Main",
        "target_columns": ["B"],
    }]

    output = merge_excel_columns(str(main_path), mappings)
    wb = load_workbook(output)
    ws = wb["Main"]
    result = [ws[f"B{i}"].value for i in range(1, 4)]
    assert result == ["x", "y", "z"]
    wb.close()


def test_merge_excel_columns_skips_empty_rows(tmp_path):
    main_path = tmp_path / "main.xlsx"
    src_path = tmp_path / "src.xlsx"

    create_wb(main_path, {}, sheet_name="Main")
    values = [None, "x", None, None, "y", None, "z", "a", "b", "c"]
    create_wb(src_path, {"A": values})

    mappings = [{
        "source": str(src_path),
        "source_columns": ["A"],
        "target_sheet": "Main",
        "target_columns": ["B"],
    }]

    output = merge_excel_columns(str(main_path), mappings)
    wb = load_workbook(output)
    ws = wb["Main"]

    assert ws["B2"].value == "x"
    assert ws["B5"].value == "y"
    assert ws["B7"].value == "z"
    assert ws["B8"].value == "a"
    assert ws["B9"].value == "b"
    assert ws["B10"].value == "c"

    assert ws["B1"].value is None
    assert ws["B3"].value is None
    assert ws["B4"].value is None
    assert ws["B6"].value is None
    wb.close()
