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
