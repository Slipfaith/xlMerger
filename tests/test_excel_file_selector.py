from gui.excel_file_selector import ExcelFileSelector

def test_excel_file_selector_smoke(app, tmp_path):
    d = tmp_path / "excel_files"
    d.mkdir()
    f1 = d / "file1.xlsx"
    f1.write_text("stub")
    # Проверяем что окно создаётся и не падает
    widget = ExcelFileSelector(str(d))
    assert widget is not None
    widget.close()
