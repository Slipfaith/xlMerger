from openpyxl import load_workbook, Workbook
from copy import copy
import os
import subprocess
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


def _run_excelltru(file_path: str) -> None:
    """Запуск Excelltru.vbs для file_path (только на Windows)."""

    if os.name != "nt":
        return

    vbs_path = os.path.join(os.path.dirname(__file__), "Excelltru.vbs")
    if not os.path.isfile(vbs_path):
        return

    try:
        norm_path = os.path.normpath(file_path)
        # Оборачиваем путь в кавычки — для пробелов и кириллицы
        subprocess.run(
            ["cscript.exe", "//nologo", vbs_path, f'"{norm_path}"'],
            check=True,
            shell=True
        )
    except Exception as e:
        print(f"⚠ Ошибка при запуске Excelltru.vbs: {e}")


def split_excel_by_languages(
    excel_path: str,
    sheet_name: str,
    source_lang: str,
    output_dir: str | None = None,
    target_langs: list[str] | None = None,
    extra_columns: list[str] | None = None,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> List[str]:
    """Split Excel into language pairs."""
    wb = load_workbook(excel_path)
    sheet = wb[sheet_name]

    header_map: Dict[str, int] = {}
    col_names: Dict[int, str] = {}
    first_row = next(sheet.iter_rows(min_row=1, max_row=1))
    for idx, cell in enumerate(first_row, start=1):
        letter = get_column_letter(idx)
        header_map[letter] = idx
        val = cell.value
        if val not in (None, ""):
            name = str(val)
            header_map[name] = idx
        else:
            name = letter
        col_names[idx] = name

    if source_lang not in header_map:
        wb.close()
        raise ValueError(f"Source column '{source_lang}' not found")

    if output_dir is None:
        output_dir = os.path.dirname(excel_path)

    target_indices = None
    if target_langs:
        missing = [t for t in target_langs if t not in header_map]
        if missing:
            wb.close()
            raise ValueError(f"Target column(s) {', '.join(missing)} not found")
        target_indices = {header_map[t] for t in target_langs}

    targets: List[tuple[str, int]] = []
    source_idx = header_map[source_lang]
    source_header = col_names[source_idx]
    for idx, name in col_names.items():
        if idx == source_idx:
            continue
        if target_indices is not None:
            if idx not in target_indices:
                continue
        else:
            if not _is_lang_column(name):
                continue
        targets.append((name, idx))

    extra_idx: List[int] = []
    extra_headers: List[str] = []
    if extra_columns:
        for col in extra_columns:
            if col in header_map and header_map[col] != source_idx:
                idx = header_map[col]
                extra_idx.append(idx)
                extra_headers.append(col_names[idx])

    created: List[str] = []
    for i, (target_lang, idx) in enumerate(targets, start=1):
        new_wb = Workbook()
        ws_new = new_wb.active
        ws_new.title = sheet_name
        col_pos = 1
        _copy_cell(sheet.cell(row=1, column=source_idx), ws_new.cell(row=1, column=col_pos))
        ws_new.cell(row=1, column=col_pos).value = source_header
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
        out_name = f"{base}_{source_header}-{target_lang}{ext}"
        out_path = os.path.join(output_dir, out_name)
        new_wb.save(out_path)
        new_wb.close()
        _run_excelltru(out_path)
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
    """Split multiple sheets preserving sheet names."""
    wb = load_workbook(excel_path)

    if output_dir is None:
        output_dir = os.path.dirname(excel_path)

    workbooks: Dict[str, Workbook] = {}
    created: List[str] = []
    source_names: set[str] = set()

    def get_wb(target: str) -> Workbook:
        if target not in workbooks:
            workbooks[target] = Workbook()
            workbooks[target].remove(workbooks[target].active)
        return workbooks[target]

    for sheet_name, (src, targets, extras) in sheet_configs.items():
        sheet = wb[sheet_name]
        header_map: Dict[str, int] = {}
        col_names: Dict[int, str] = {}
        first_row = next(sheet.iter_rows(min_row=1, max_row=1))
        for idx, cell in enumerate(first_row, start=1):
            letter = get_column_letter(idx)
            header_map[letter] = idx
            val = cell.value
            if val not in (None, ""):
                name = str(val)
                header_map[name] = idx
            else:
                name = letter
            col_names[idx] = name

        if src not in header_map:
            wb.close()
            raise ValueError(f"Source column '{src}' not found in sheet '{sheet_name}'")

        src_idx = header_map[src]
        src_name = col_names[src_idx]
        source_names.add(src_name)

        target_indices = None
        if targets:
            missing = [t for t in targets if t not in header_map]
            if missing:
                wb.close()
                raise ValueError(
                    f"Target column(s) {', '.join(missing)} not found in sheet '{sheet_name}'"
                )
            target_indices = {header_map[t] for t in targets}

        col_targets: List[Tuple[str, int]] = []
        for idx, name in col_names.items():
            if idx == src_idx:
                continue
            if target_indices is not None:
                if idx not in target_indices:
                    continue
            else:
                if not _is_lang_column(name):
                    continue
            col_targets.append((name, idx))

        extra_idx: List[int] = []
        extra_headers: List[str] = []
        if extras:
            for col in extras:
                if col in header_map and header_map[col] != src_idx:
                    i = header_map[col]
                    extra_idx.append(i)
                    extra_headers.append(col_names[i])

        for tgt_name, idx in col_targets:
            wb_out = get_wb(tgt_name)
            if sheet_name in wb_out.sheetnames:
                ws_new = wb_out[sheet_name]
            else:
                ws_new = wb_out.create_sheet(title=sheet_name)

            if ws_new.max_row == 1 and ws_new.max_column == 1 and ws_new.cell(row=1, column=1).value is None:
                col_pos = 1
                _copy_cell(sheet.cell(row=1, column=src_idx), ws_new.cell(row=1, column=col_pos))
                ws_new.cell(row=1, column=col_pos).value = src_name
                _copy_column_width(sheet, ws_new, src_idx, col_pos)
                col_pos += 1
                _copy_cell(sheet.cell(row=1, column=idx), ws_new.cell(row=1, column=col_pos))
                ws_new.cell(row=1, column=col_pos).value = tgt_name
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
    sources = source_names

    for i, (tgt, new_wb) in enumerate(workbooks.items(), start=1):
        src_part = next(iter(sources)) if len(sources) == 1 else "src"
        out_name = f"{base}_{src_part}-{tgt}{ext}"
        out_path = os.path.join(output_dir, out_name)
        new_wb.save(out_path)
        new_wb.close()
        _run_excelltru(out_path)
        created.append(out_path)
        if progress_callback:
            progress_callback(i, len(workbooks), out_name)

    wb.close()
    return created
