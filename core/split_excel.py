from openpyxl import load_workbook, Workbook
from copy import copy
import os
from typing import Callable, List, Dict, Tuple
from openpyxl.utils import get_column_letter


def _is_lang_column(name: str) -> bool:
    if not name:
        return False
    name = str(name).strip()
    if len(name) > 5 or ' ' in name or '_' in name:
        return False
    return name.isalpha()


def _copy_cell(src, dst):
    """Copy value and style from ``src`` cell to ``dst`` cell."""
    dst.value = src.value
    if src.has_style:
        dst.font = copy(src.font)
        dst.border = copy(src.border)
        dst.fill = copy(src.fill)
        dst.number_format = copy(src.number_format)
        dst.protection = copy(src.protection)
        dst.alignment = copy(src.alignment)

def _copy_column_width(src_sheet, dst_sheet, src_idx: int, dst_idx: int):
    src_letter = get_column_letter(src_idx)
    dst_letter = get_column_letter(dst_idx)
    width = src_sheet.column_dimensions[src_letter].width
    if width is not None:
        dst_sheet.column_dimensions[dst_letter].width = width


def _find_last_data_row(sheet, columns: List[int]) -> int:
    """Return the last row index that has a value in the given columns."""
    for row_idx in range(sheet.max_row, 1, -1):
        for col in columns:
            val = sheet.cell(row=row_idx, column=col).value
            if val not in (None, ""):
                return row_idx
    return 1


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
        if target_langs is not None:
            if target_lang not in target_langs:
                continue
        else:
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
        _copy_cell(sheet.cell(row=1, column=source_idx), ws_new.cell(row=1, column=col_pos))
        ws_new.cell(row=1, column=col_pos).value = source_lang
        _copy_column_width(sheet, ws_new, source_idx, col_pos)
        col_pos += 1
        _copy_cell(sheet.cell(row=1, column=idx), ws_new.cell(row=1, column=col_pos))
        ws_new.cell(row=1, column=col_pos).value = target_lang
        _copy_column_width(sheet, ws_new, idx, col_pos)
        col_pos += 1
        for header in extra_headers:
            ex_idx = header_map[header]
            _copy_cell(sheet.cell(row=1, column=ex_idx), ws_new.cell(row=1, column=col_pos))
            ws_new.cell(row=1, column=col_pos).value = header
            _copy_column_width(sheet, ws_new, ex_idx, col_pos)
            col_pos += 1
        last_row = _find_last_data_row(sheet, [source_idx, idx, *extra_idx])
        for row in range(2, last_row + 1):
            col_pos = 1
            _copy_cell(sheet.cell(row=row, column=source_idx), ws_new.cell(row=row, column=col_pos))
            col_pos += 1
            _copy_cell(sheet.cell(row=row, column=idx), ws_new.cell(row=row, column=col_pos))
            col_pos += 1
            for ex_idx in extra_idx:
                _copy_cell(sheet.cell(row=row, column=ex_idx), ws_new.cell(row=row, column=col_pos))
                col_pos += 1
        base, ext = os.path.splitext(os.path.basename(excel_path))
        out_name = f"{base}_{source_lang}-{target_lang}{ext}"
        out_path = os.path.join(output_dir, out_name)
        new_wb.save(out_path)
        new_wb.close()
        created.append(out_path)
        if progress_callback:
            progress_callback(i, len(targets), out_name)

    wb.close()
    return created


def split_excel_multiple_sheets(
    excel_path: str,
    sheet_configs: Dict[str, Tuple[str, List[str] | None, List[str] | None]],
    output_dir: str | None = None,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> List[str]:
    """Split multiple sheets preserving sheet names.

    Parameters
    ----------
    excel_path : str
        Path to the source Excel file.
    sheet_configs : Dict[str, Tuple[str, List[str] | None, List[str] | None]]
        Mapping of sheet name to ``(source_lang, target_langs, extra_columns)``.
    output_dir : str | None, optional
        Directory where new files will be saved. Defaults to the Excel file
        directory.
    progress_callback : Callable[[int, int, str], None] | None, optional
        Called after each file is saved.
    """
    wb = load_workbook(excel_path)

    if output_dir is None:
        output_dir = os.path.dirname(excel_path)

    workbooks: Dict[str, Workbook] = {}
    created: List[str] = []

    def get_wb(target: str) -> Workbook:
        if target not in workbooks:
            workbooks[target] = Workbook()
            workbooks[target].remove(workbooks[target].active)
        return workbooks[target]

    for sheet_name, (src, targets, extras) in sheet_configs.items():
        sheet = wb[sheet_name]
        headers = [cell.value for cell in next(sheet.iter_rows(min_row=1, max_row=1))]
        header_map = {str(h): idx + 1 for idx, h in enumerate(headers) if h is not None}

        if src not in header_map:
            wb.close()
            raise ValueError(f"Source column '{src}' not found in sheet '{sheet_name}'")

        if targets:
            missing = [t for t in targets if t not in header_map]
            if missing:
                wb.close()
                raise ValueError(
                    f"Target column(s) {', '.join(missing)} not found in sheet '{sheet_name}'"
                )

        col_targets: List[Tuple[str, int]] = []
        for tgt, idx in header_map.items():
            if tgt == src:
                continue
            if targets is not None:
                if tgt not in targets:
                    continue
            else:
                if not _is_lang_column(tgt):
                    continue
            col_targets.append((tgt, idx))

        extra_idx: List[int] = []
        extra_headers: List[str] = []
        if extras:
            for col in extras:
                if col in header_map and col not in [src]:
                    extra_idx.append(header_map[col])
                    extra_headers.append(col)

        src_idx = header_map[src]

        for tgt, idx in col_targets:
            wb_out = get_wb(tgt)
            if sheet_name in wb_out.sheetnames:
                ws_new = wb_out[sheet_name]
            else:
                ws_new = wb_out.create_sheet(title=sheet_name)

            if ws_new.max_row == 1 and ws_new.max_column == 1 and ws_new.cell(row=1, column=1).value is None:
                col_pos = 1
                _copy_cell(sheet.cell(row=1, column=src_idx), ws_new.cell(row=1, column=col_pos))
                ws_new.cell(row=1, column=col_pos).value = src
                _copy_column_width(sheet, ws_new, src_idx, col_pos)
                col_pos += 1
                _copy_cell(sheet.cell(row=1, column=idx), ws_new.cell(row=1, column=col_pos))
                ws_new.cell(row=1, column=col_pos).value = tgt
                _copy_column_width(sheet, ws_new, idx, col_pos)
                col_pos += 1
                for header in extra_headers:
                    ex_idx = header_map[header]
                    _copy_cell(sheet.cell(row=1, column=ex_idx), ws_new.cell(row=1, column=col_pos))
                    ws_new.cell(row=1, column=col_pos).value = header
                    _copy_column_width(sheet, ws_new, ex_idx, col_pos)
                    col_pos += 1

            last_row = _find_last_data_row(sheet, [src_idx, idx, *extra_idx])
            for row in range(2, last_row + 1):
                col_pos = 1
                _copy_cell(sheet.cell(row=row, column=src_idx), ws_new.cell(row=row, column=col_pos))
                col_pos += 1
                _copy_cell(sheet.cell(row=row, column=idx), ws_new.cell(row=row, column=col_pos))
                col_pos += 1
                for ex_idx in extra_idx:
                    _copy_cell(sheet.cell(row=row, column=ex_idx), ws_new.cell(row=row, column=col_pos))
                    col_pos += 1

    base, ext = os.path.splitext(os.path.basename(excel_path))
    sources = {cfg[0] for cfg in sheet_configs.values()}

    for i, (tgt, new_wb) in enumerate(workbooks.items(), start=1):
        src_part = next(iter(sources)) if len(sources) == 1 else "src"
        out_name = f"{base}_{src_part}-{tgt}{ext}"
        out_path = os.path.join(output_dir, out_name)
        new_wb.save(out_path)
        new_wb.close()
        created.append(out_path)
        if progress_callback:
            progress_callback(i, len(workbooks), out_name)

    wb.close()
    return created
