from openpyxl import load_workbook, Workbook
import os
from typing import Callable, Iterable, List


def _is_lang_column(name: str) -> bool:
    if not name:
        return False
    name = str(name).strip()
    if len(name) > 5 or ' ' in name or '_' in name:
        return False
    return name.isalpha()


def split_excel_by_languages(
    excel_path: str,
    sheet_name: str,
    source_lang: str,
    output_dir: str | None = None,
    target_langs: list[str] | None = None,
    extra_columns: list[str] | None = None,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> List[str]:
    """Split Excel into language pairs.

    Parameters
    ----------
    excel_path : str
        Path to the source Excel file.
    sheet_name : str
        Name of the sheet to process.
    source_lang : str
        Column header that contains source text.
    output_dir : str | None, optional
        Directory where new files will be saved. Defaults to the Excel file
        directory.
    target_langs : list[str] | None, optional
        List of target language columns to include. If ``None`` all language
        columns are used.
    extra_columns : list[str] | None, optional
        Additional columns to copy to each output file.
    progress_callback : Callable[[int, int, str], None] | None, optional
        Called after each file is saved with ``(index, total, name)``.
    """
    wb = load_workbook(excel_path)
    sheet = wb[sheet_name]
    headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
    header_map = {str(h): idx + 1 for idx, h in enumerate(headers) if h is not None}

    if source_lang not in header_map:
        wb.close()
        raise ValueError(f"Source column '{source_lang}' not found")

    if output_dir is None:
        output_dir = os.path.dirname(excel_path)

    if target_langs:
        missing = [t for t in target_langs if t not in header_map]
        if missing:
            wb.close()
            raise ValueError(f"Target column(s) {', '.join(missing)} not found")

    targets: List[tuple[str, int]] = []
    for target_lang, idx in header_map.items():
        if target_lang == source_lang:
            continue
        if target_langs is not None and target_lang not in target_langs:
            continue
        if not _is_lang_column(target_lang):
            continue
        targets.append((target_lang, idx))

    extra_idx: List[int] = []
    extra_headers: List[str] = []
    if extra_columns:
        for col in extra_columns:
            if col in header_map and col not in [source_lang]:
                extra_idx.append(header_map[col])
                extra_headers.append(col)

    created: List[str] = []

    source_idx = header_map[source_lang]

    for i, (target_lang, idx) in enumerate(targets, start=1):
        new_wb = Workbook()
        ws_new = new_wb.active
        ws_new.title = sheet_name
        col_pos = 1
        ws_new.cell(row=1, column=col_pos, value=source_lang)
        col_pos += 1
        ws_new.cell(row=1, column=col_pos, value=target_lang)
        col_pos += 1
        for header in extra_headers:
            ws_new.cell(row=1, column=col_pos, value=header)
            col_pos += 1
        for row in range(2, sheet.max_row + 1):
            col_pos = 1
            ws_new.cell(row=row, column=col_pos, value=sheet.cell(row=row, column=source_idx).value)
            col_pos += 1
            ws_new.cell(row=row, column=col_pos, value=sheet.cell(row=row, column=idx).value)
            col_pos += 1
            for ex_idx in extra_idx:
                ws_new.cell(row=row, column=col_pos, value=sheet.cell(row=row, column=ex_idx).value)
                col_pos += 1
        base, ext = os.path.splitext(os.path.basename(excel_path))
        out_name = f"{source_lang}-{target_lang}_{base}{ext}"
        out_path = os.path.join(output_dir, out_name)
        new_wb.save(out_path)
        new_wb.close()
        created.append(out_path)
        if progress_callback:
            progress_callback(i, len(targets), out_name)

    wb.close()
    return created
