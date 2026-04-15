# -*- coding: utf-8 -*-
from __future__ import annotations

from typing import Dict, Tuple

import xlsxwriter
from openpyxl.utils.cell import column_index_from_string, coordinate_to_tuple, range_boundaries


_BORDER_STYLE_MAP = {
    None: 0,
    "thin": 1,
    "medium": 2,
    "dashed": 3,
    "dotted": 4,
    "thick": 5,
    "double": 6,
    "hair": 7,
    "mediumDashed": 8,
    "dashDot": 9,
    "mediumDashDot": 10,
    "dashDotDot": 11,
    "mediumDashDotDot": 12,
    "slantDashDot": 13,
}


def _normalize_color(color) -> str | None:
    if not color:
        return None

    rgb = getattr(color, "rgb", None)
    if isinstance(rgb, bytes):
        rgb = rgb.decode()
    if not isinstance(rgb, str):
        nested_rgb = getattr(rgb, "rgb", None) or getattr(rgb, "value", None)
        if isinstance(nested_rgb, bytes):
            nested_rgb = nested_rgb.decode()
        rgb = nested_rgb if isinstance(nested_rgb, str) else None

    if not rgb and isinstance(color, str):
        rgb = color

    if not rgb:
        return None

    rgb = rgb.strip().lstrip("#")
    if len(rgb) == 8:
        rgb = rgb[-6:]
    if len(rgb) != 6:
        return None
    return rgb.upper()


def _map_horizontal(value: str | None) -> str | None:
    if value == "centerContinuous":
        return "center_across"
    return value


def _map_vertical(value: str | None) -> str | None:
    if value == "center":
        return "vcenter"
    if value == "justify":
        return "vjustify"
    if value == "distributed":
        return "vdistributed"
    return value


def _build_format_key(cell) -> Tuple:
    font = cell.font
    fill = cell.fill
    border = cell.border
    alignment = cell.alignment
    protection = cell.protection

    return (
        bool(font.bold),
        bool(font.italic),
        bool(font.strike),
        font.underline,
        font.name,
        font.sz,
        _normalize_color(font.color),
        fill.fill_type,
        _normalize_color(fill.fgColor),
        _BORDER_STYLE_MAP.get(border.left.style, 1 if border.left.style else 0),
        _BORDER_STYLE_MAP.get(border.right.style, 1 if border.right.style else 0),
        _BORDER_STYLE_MAP.get(border.top.style, 1 if border.top.style else 0),
        _BORDER_STYLE_MAP.get(border.bottom.style, 1 if border.bottom.style else 0),
        _normalize_color(border.left.color),
        _normalize_color(border.right.color),
        _normalize_color(border.top.color),
        _normalize_color(border.bottom.color),
        cell.number_format,
        _map_horizontal(alignment.horizontal),
        _map_vertical(alignment.vertical),
        bool(alignment.wrap_text),
        alignment.text_rotation,
        alignment.indent,
        bool(alignment.shrink_to_fit),
        protection.locked,
        protection.hidden,
    )


def _get_xlsxwriter_format(workbook, cell, cache: Dict[Tuple, object]):
    if cell is None or not cell.has_style:
        return None

    key = _build_format_key(cell)
    if key in cache:
        return cache[key]

    font = cell.font
    fill = cell.fill
    border = cell.border
    alignment = cell.alignment
    protection = cell.protection

    fmt_args: Dict[str, object] = {}

    if font.bold:
        fmt_args["bold"] = True
    if font.italic:
        fmt_args["italic"] = True
    if font.strike:
        fmt_args["font_strikeout"] = True
    if font.underline:
        fmt_args["underline"] = font.underline if isinstance(font.underline, str) else True
    if font.name:
        fmt_args["font_name"] = font.name
    if font.sz:
        fmt_args["font_size"] = font.sz
    font_color = _normalize_color(font.color)
    if font_color:
        fmt_args["font_color"] = f"#{font_color}"

    fill_color = _normalize_color(fill.fgColor) if fill.fill_type == "solid" else None
    if fill_color:
        fmt_args["bg_color"] = f"#{fill_color}"
        fmt_args["pattern"] = 1

    left_style = _BORDER_STYLE_MAP.get(border.left.style, 1 if border.left.style else 0)
    right_style = _BORDER_STYLE_MAP.get(border.right.style, 1 if border.right.style else 0)
    top_style = _BORDER_STYLE_MAP.get(border.top.style, 1 if border.top.style else 0)
    bottom_style = _BORDER_STYLE_MAP.get(border.bottom.style, 1 if border.bottom.style else 0)
    if left_style:
        fmt_args["left"] = left_style
    if right_style:
        fmt_args["right"] = right_style
    if top_style:
        fmt_args["top"] = top_style
    if bottom_style:
        fmt_args["bottom"] = bottom_style

    left_color = _normalize_color(border.left.color)
    right_color = _normalize_color(border.right.color)
    top_color = _normalize_color(border.top.color)
    bottom_color = _normalize_color(border.bottom.color)
    if left_color:
        fmt_args["left_color"] = f"#{left_color}"
    if right_color:
        fmt_args["right_color"] = f"#{right_color}"
    if top_color:
        fmt_args["top_color"] = f"#{top_color}"
    if bottom_color:
        fmt_args["bottom_color"] = f"#{bottom_color}"

    if cell.number_format and cell.number_format != "General":
        fmt_args["num_format"] = cell.number_format

    horizontal = _map_horizontal(alignment.horizontal)
    vertical = _map_vertical(alignment.vertical)
    if horizontal:
        fmt_args["align"] = horizontal
    if vertical:
        fmt_args["valign"] = vertical
    if alignment.wrap_text:
        fmt_args["text_wrap"] = True
    if alignment.indent:
        fmt_args["indent"] = int(alignment.indent)
    if alignment.shrink_to_fit:
        fmt_args["shrink"] = True

    rotation = alignment.text_rotation
    if isinstance(rotation, int):
        if rotation == 255:
            fmt_args["rotation"] = 270
        elif 0 <= rotation <= 90:
            fmt_args["rotation"] = rotation

    if protection.locked is not None:
        fmt_args["locked"] = bool(protection.locked)
    if protection.hidden is not None:
        fmt_args["hidden"] = bool(protection.hidden)

    fmt = workbook.add_format(fmt_args)
    cache[key] = fmt
    return fmt


def _copy_sheet_dimensions(sheet, ws_out) -> None:
    for key, dim in sheet.column_dimensions.items():
        if dim.min is not None and dim.max is not None:
            start_col = dim.min
            end_col = dim.max
        else:
            try:
                start_col = end_col = column_index_from_string(key)
            except Exception:  # noqa: BLE001
                continue

        options: Dict[str, object] = {}
        if dim.hidden:
            options["hidden"] = True
        if dim.outlineLevel is not None and dim.outlineLevel > 0:
            options["level"] = int(dim.outlineLevel)
        if dim.collapsed:
            options["collapsed"] = True

        if dim.width is None and not options:
            continue
        ws_out.set_column(start_col - 1, end_col - 1, dim.width, None, options or None)

    for row_idx, dim in sheet.row_dimensions.items():
        options: Dict[str, object] = {}
        if dim.hidden:
            options["hidden"] = True
        if dim.outlineLevel is not None and dim.outlineLevel > 0:
            options["level"] = int(dim.outlineLevel)
        if dim.collapsed:
            options["collapsed"] = True

        if dim.height is None and not options:
            continue
        ws_out.set_row(row_idx - 1, dim.height, None, options or None)


def _copy_sheet_properties(sheet, ws_out) -> None:
    if sheet.freeze_panes:
        coord = (
            sheet.freeze_panes.coordinate
            if hasattr(sheet.freeze_panes, "coordinate")
            else str(sheet.freeze_panes)
        )
        if coord and coord != "A1":
            row, col = coordinate_to_tuple(coord)
            ws_out.freeze_panes(row - 1, col - 1)

    if sheet.auto_filter and sheet.auto_filter.ref:
        min_col, min_row, max_col, max_row = range_boundaries(sheet.auto_filter.ref)
        ws_out.autofilter(min_row - 1, min_col - 1, max_row - 1, max_col - 1)

    if sheet.sheet_state == "hidden":
        ws_out.hide()
    elif sheet.sheet_state == "veryHidden":
        ws_out.very_hidden()

    if getattr(sheet.sheet_view, "rightToLeft", False):
        ws_out.right_to_left()

    if getattr(sheet.sheet_view, "showGridLines", True) is False:
        ws_out.hide_gridlines(2)

    zoom = getattr(sheet.sheet_view, "zoomScale", None)
    if isinstance(zoom, int) and zoom > 0:
        ws_out.set_zoom(zoom)

    tab_color = _normalize_color(getattr(sheet.sheet_properties, "tabColor", None))
    if tab_color:
        ws_out.set_tab_color(f"#{tab_color}")


def _copy_sheet_cells(sheet, ws_out, workbook_out, fmt_cache: Dict[Tuple, object]) -> None:
    merged_skip: set[tuple[int, int]] = set()
    merged_written: set[tuple[int, int]] = set()

    for merged_range in sheet.merged_cells.ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        top_left = (min_row, min_col)
        top_left_cell = sheet.cell(row=min_row, column=min_col)
        fmt = _get_xlsxwriter_format(workbook_out, top_left_cell, fmt_cache)
        value = top_left_cell.value if top_left_cell.value is not None else ""

        if min_row != max_row or min_col != max_col:
            ws_out.merge_range(
                min_row - 1,
                min_col - 1,
                max_row - 1,
                max_col - 1,
                value,
                fmt,
            )
            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    if (row, col) != top_left:
                        merged_skip.add((row, col))
        else:
            ws_out.write(min_row - 1, min_col - 1, value, fmt)

        merged_written.add(top_left)

    for (row_idx, col_idx), cell in sorted(sheet._cells.items()):
        if (row_idx, col_idx) in merged_skip or (row_idx, col_idx) in merged_written:
            continue

        fmt = _get_xlsxwriter_format(workbook_out, cell, fmt_cache)
        value = cell.value

        hyperlink = getattr(cell, "hyperlink", None)
        if hyperlink and getattr(hyperlink, "target", None):
            display = str(value) if value is not None else hyperlink.target
            ws_out.write_url(row_idx - 1, col_idx - 1, hyperlink.target, fmt, display)
            continue

        if value is None:
            if fmt is not None:
                ws_out.write_blank(row_idx - 1, col_idx - 1, None, fmt)
        else:
            ws_out.write(row_idx - 1, col_idx - 1, value, fmt)


def save_openpyxl_workbook_with_xlsxwriter(workbook, output_path: str) -> None:
    """Persist an openpyxl workbook using xlsxwriter with style preservation."""
    wb_out = xlsxwriter.Workbook(output_path)
    fmt_cache: Dict[Tuple, object] = {}

    try:
        for sheet in workbook.worksheets:
            ws_out = wb_out.add_worksheet(sheet.title)
            _copy_sheet_dimensions(sheet, ws_out)
            _copy_sheet_properties(sheet, ws_out)
            _copy_sheet_cells(sheet, ws_out, wb_out, fmt_cache)
    finally:
        wb_out.close()
