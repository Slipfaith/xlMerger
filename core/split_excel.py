from openpyxl import load_workbook, Workbook
import os


def _is_lang_column(name: str) -> bool:
    if not name:
        return False
    name = str(name).strip()
    if len(name) > 5 or ' ' in name or '_' in name:
        return False
    return name.isalpha()


def split_excel_by_languages(excel_path: str, sheet_name: str, source_lang: str, output_dir: str | None = None):
    wb = load_workbook(excel_path, read_only=True)
    sheet = wb[sheet_name]
    headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
    header_map = {str(h): idx + 1 for idx, h in enumerate(headers) if h is not None}

    if source_lang not in header_map:
        wb.close()
        raise ValueError(f"Source column '{source_lang}' not found")

    if output_dir is None:
        output_dir = os.path.dirname(excel_path)

    source_idx = header_map[source_lang]

    for target_lang, idx in header_map.items():
        if target_lang == source_lang:
            continue
        if not _is_lang_column(target_lang):
            continue
        new_wb = Workbook()
        ws_new = new_wb.active
        ws_new.title = sheet_name
        ws_new.cell(row=1, column=1, value=source_lang)
        ws_new.cell(row=1, column=2, value=target_lang)
        for row in range(2, sheet.max_row + 1):
            ws_new.cell(row=row, column=1, value=sheet.cell(row=row, column=source_idx).value)
            ws_new.cell(row=row, column=2, value=sheet.cell(row=row, column=idx).value)
        base, ext = os.path.splitext(os.path.basename(excel_path))
        out_name = f"{source_lang}-{target_lang}_{base}{ext}"
        out_path = os.path.join(output_dir, out_name)
        new_wb.save(out_path)
        new_wb.close()

    wb.close()
