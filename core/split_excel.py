# -*- coding: utf-8 -*-
from openpyxl import load_workbook
import os
from typing import Callable, List, Dict, Tuple
import xlsxwriter
from openpyxl.utils import get_column_letter


_SOURCE_ONLY_TARGET = "__source_only_output__"


def _is_lang_column(name: str) -> bool:
    if not name:
        return False
    name = str(name).strip()
    if len(name) > 5 or " " in name or "_" in name:
        return False
    return name.isalpha()


def _normalize_color(color) -> str | None:
    if not color:
        return None
    rgb = getattr(color, "rgb", None)
    if not rgb:
        return None
    if isinstance(rgb, bytes):
        rgb = rgb.decode()
    # Some color objects expose an RGB helper instead of a plain string
    if not isinstance(rgb, str):
        nested_rgb = getattr(rgb, "rgb", None) or getattr(rgb, "value", None)
        if isinstance(nested_rgb, bytes):
            nested_rgb = nested_rgb.decode()
        rgb = nested_rgb if isinstance(nested_rgb, str) else None
    if not rgb:
        return None
    if len(rgb) == 8:
        return rgb[-6:]
    if len(rgb) == 6:
        return rgb
    return None


def _get_xlsxwriter_format(workbook, cell, cache):
    if cell is None or not cell.has_style:
        return None

    font = cell.font
    fill = cell.fill
    alignment = cell.alignment

    font_color = _normalize_color(font.color)
    fill_color = (
        _normalize_color(fill.fgColor)
        if getattr(fill, "fill_type", None) == "solid"
        else None
    )

    key = (
        bool(font.bold),
        bool(font.italic),
        font.underline,
        font.name,
        font.sz,
        font_color,
        fill_color,
        cell.number_format,
        alignment.horizontal,
        alignment.vertical,
        alignment.wrap_text,
    )

    if key in cache:
        return cache[key]

    fmt_args: Dict[str, object] = {}
    if font.bold:
        fmt_args["bold"] = True
    if font.italic:
        fmt_args["italic"] = True
    if font.underline:
        fmt_args["underline"] = (
            font.underline if isinstance(font.underline, str) else True
        )
    if font.name:
        fmt_args["font_name"] = font.name
    if font.sz:
        fmt_args["font_size"] = font.sz
    if font_color:
        fmt_args["font_color"] = font_color
    if fill_color:
        fmt_args["bg_color"] = fill_color
        fmt_args["pattern"] = 1
    if cell.number_format:
        fmt_args["num_format"] = cell.number_format
    if alignment.horizontal:
        fmt_args["align"] = alignment.horizontal
    if alignment.vertical:
        fmt_args["valign"] = alignment.vertical
    if alignment.wrap_text:
        fmt_args["text_wrap"] = True

    fmt = workbook.add_format(fmt_args)
    cache[key] = fmt
    return fmt


def _write_rows(
    ws, rows: List[List[Tuple[object, object]]], widths: List[float | None], workbook
) -> None:
    fmt_cache: Dict[Tuple, object] = {}
    for idx, width in enumerate(widths):
        if width is not None:
            ws.set_column(idx, idx, width)

    for r_idx, row in enumerate(rows):
        for c_idx, (value, cell) in enumerate(row):
            fmt = _get_xlsxwriter_format(workbook, cell, fmt_cache)
            ws.write(r_idx, c_idx, value, fmt)


def _find_last_data_row(sheet, columns: List[int]) -> int:
    """Return the last row index that has a value in the given columns."""
    for row_idx in range(sheet.max_row, 1, -1):
        for col in columns:
            val = sheet.cell(row=row_idx, column=col).value
            if val not in (None, ""):
                return row_idx
    return 1


def _build_header_maps(sheet) -> tuple[dict[str, int], dict[int, str]]:
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
    return header_map, col_names


def _resolve_extra_columns(
    source_idx: int,
    extra_columns: list[str] | None,
    header_map: dict[str, int],
    col_names: dict[int, str],
) -> tuple[list[int], list[str]]:
    extra_idx: list[int] = []
    extra_headers: list[str] = []
    seen: set[int] = set()
    if not extra_columns:
        return extra_idx, extra_headers

    for col in extra_columns:
        idx = header_map.get(col)
        if idx is None or idx == source_idx or idx in seen:
            continue
        seen.add(idx)
        extra_idx.append(idx)
        extra_headers.append(col_names[idx])
    return extra_idx, extra_headers


def _build_rows_for_output(
    sheet,
    source_idx: int,
    source_header: str,
    extra_idx: list[int],
    extra_headers: list[str],
    target_idx: int | None = None,
    target_header: str | None = None,
) -> tuple[list[list[tuple[object, object]]], list[float | None]]:
    rows: list[list[tuple[object, object]]] = []
    widths: list[float | None] = []

    headers: list[tuple[object, object]] = []
    for ex_idx, header in zip(extra_idx, extra_headers):
        cell = sheet.cell(row=1, column=ex_idx)
        headers.append((header, cell))
        widths.append(sheet.column_dimensions[get_column_letter(ex_idx)].width)

    source_cell = sheet.cell(row=1, column=source_idx)
    headers.append((source_header, source_cell))
    widths.append(sheet.column_dimensions[get_column_letter(source_idx)].width)

    data_columns = [*extra_idx, source_idx]
    if target_idx is not None and target_header is not None:
        target_cell = sheet.cell(row=1, column=target_idx)
        headers.append((target_header, target_cell))
        widths.append(sheet.column_dimensions[get_column_letter(target_idx)].width)
        data_columns.append(target_idx)

    rows.append(headers)

    last_row = _find_last_data_row(sheet, data_columns)
    for row in range(2, last_row + 1):
        row_data: list[tuple[object, object]] = []
        for ex_idx in extra_idx:
            extra_cell = sheet.cell(row=row, column=ex_idx)
            row_data.append((extra_cell.value, extra_cell))

        src_cell = sheet.cell(row=row, column=source_idx)
        row_data.append((src_cell.value, src_cell))

        if target_idx is not None:
            tgt_cell = sheet.cell(row=row, column=target_idx)
            row_data.append((tgt_cell.value, tgt_cell))
        rows.append(row_data)

    return rows, widths


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
    header_map, col_names = _build_header_maps(sheet)

    if source_lang not in header_map:
        wb.close()
        raise ValueError(f"Source column '{source_lang}' not found")

    if output_dir is None:
        output_dir = os.path.dirname(excel_path)

    target_indices: set[int] | None = None
    if target_langs is not None:
        missing = [t for t in target_langs if t not in header_map]
        if missing:
            wb.close()
            raise ValueError(f"Target column(s) {', '.join(missing)} not found")
        target_indices = {header_map[t] for t in target_langs}

    source_idx = header_map[source_lang]
    source_header = col_names[source_idx]
    targets: list[tuple[str, int]] = []
    for idx, name in col_names.items():
        if idx == source_idx:
            continue
        if target_indices is not None:
            if idx not in target_indices:
                continue
        elif not _is_lang_column(name):
            continue
        targets.append((name, idx))

    extra_idx, extra_headers = _resolve_extra_columns(
        source_idx=source_idx,
        extra_columns=extra_columns,
        header_map=header_map,
        col_names=col_names,
    )

    created: list[str] = []
    base, ext = os.path.splitext(os.path.basename(excel_path))

    if not targets:
        rows, widths = _build_rows_for_output(
            sheet=sheet,
            source_idx=source_idx,
            source_header=source_header,
            extra_idx=extra_idx,
            extra_headers=extra_headers,
        )
        out_name = f"{base}_{source_header}{ext}"
        out_path = os.path.join(output_dir, out_name)

        wb_out = xlsxwriter.Workbook(out_path)
        ws_new = wb_out.add_worksheet(sheet_name)
        _write_rows(ws_new, rows, widths, wb_out)
        wb_out.close()
        created.append(out_path)
        if progress_callback:
            progress_callback(1, 1, out_name)
        wb.close()
        return created

    for i, (target_name, target_idx) in enumerate(targets, start=1):
        rows, widths = _build_rows_for_output(
            sheet=sheet,
            source_idx=source_idx,
            source_header=source_header,
            extra_idx=extra_idx,
            extra_headers=extra_headers,
            target_idx=target_idx,
            target_header=target_name,
        )
        out_name = f"{base}_{source_header}-{target_name}{ext}"
        out_path = os.path.join(output_dir, out_name)

        wb_out = xlsxwriter.Workbook(out_path)
        ws_new = wb_out.add_worksheet(sheet_name)
        _write_rows(ws_new, rows, widths, wb_out)
        wb_out.close()
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

    workbooks: Dict[str, Dict[str, Dict[str, object]]] = {}
    created: List[str] = []
    source_names: set[str] = set()

    for sheet_name, (src, targets, extras) in sheet_configs.items():
        sheet = wb[sheet_name]
        header_map, col_names = _build_header_maps(sheet)

        if src not in header_map:
            wb.close()
            raise ValueError(f"Source column '{src}' not found in sheet '{sheet_name}'")

        src_idx = header_map[src]
        src_name = col_names[src_idx]
        source_names.add(src_name)

        target_indices: set[int] | None = None
        if targets is not None:
            missing = [t for t in targets if t not in header_map]
            if missing:
                wb.close()
                raise ValueError(
                    f"Target column(s) {', '.join(missing)} not found in sheet '{sheet_name}'"
                )
            target_indices = {header_map[t] for t in targets}

        col_targets: list[tuple[str, int]] = []
        for idx, name in col_names.items():
            if idx == src_idx:
                continue
            if target_indices is not None:
                if idx not in target_indices:
                    continue
            elif not _is_lang_column(name):
                continue
            col_targets.append((name, idx))

        extra_idx, extra_headers = _resolve_extra_columns(
            source_idx=src_idx,
            extra_columns=extras,
            header_map=header_map,
            col_names=col_names,
        )

        if not col_targets:
            rows, widths = _build_rows_for_output(
                sheet=sheet,
                source_idx=src_idx,
                source_header=src_name,
                extra_idx=extra_idx,
                extra_headers=extra_headers,
            )
            source_only_workbook = workbooks.setdefault(_SOURCE_ONLY_TARGET, {})
            source_only_workbook[sheet_name] = {"rows": rows, "widths": widths}
            continue

        for tgt_name, tgt_idx in col_targets:
            rows, widths = _build_rows_for_output(
                sheet=sheet,
                source_idx=src_idx,
                source_header=src_name,
                extra_idx=extra_idx,
                extra_headers=extra_headers,
                target_idx=tgt_idx,
                target_header=tgt_name,
            )
            tgt_workbook = workbooks.setdefault(tgt_name, {})
            tgt_workbook[sheet_name] = {"rows": rows, "widths": widths}

    base, ext = os.path.splitext(os.path.basename(excel_path))
    src_part = next(iter(source_names)) if len(source_names) == 1 else "src"

    for i, (tgt, sheets) in enumerate(workbooks.items(), start=1):
        if tgt == _SOURCE_ONLY_TARGET:
            out_name = f"{base}_{src_part}{ext}"
        else:
            out_name = f"{base}_{src_part}-{tgt}{ext}"

        out_path = os.path.join(output_dir, out_name)
        wb_out = xlsxwriter.Workbook(out_path)
        for sheet_name, info in sheets.items():
            ws_new = wb_out.add_worksheet(sheet_name)
            _write_rows(ws_new, info["rows"], info["widths"], wb_out)
        wb_out.close()
        created.append(out_path)
        if progress_callback:
            progress_callback(i, len(workbooks), out_name)

    wb.close()
    return created
