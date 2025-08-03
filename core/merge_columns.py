import os
from typing import List, Dict
from openpyxl import load_workbook


def merge_excel_columns(main_file: str, mappings: List[Dict[str, object]], output_file: str | None = None,
                        progress_callback=None) -> str:
    """Merge columns from multiple Excel files into a main workbook.

    Args:
        main_file: Path to the main Excel workbook.
        mappings: A list of mapping dictionaries. Each mapping must contain:
            - ``source``: path to the source Excel file.
            - ``source_columns``: list of column letters in the source file.
            - ``target_sheet``: sheet name in the main workbook.
            - ``target_columns``: list of column letters in the target sheet.
              ``source_columns`` and ``target_columns`` must have the same length.
        output_file: Optional path where the merged workbook will be saved. If
            not provided, ``main_file`` suffixed with ``_merged`` is used.
        progress_callback: Optional callback function(idx, total, mapping) for progress updates.

    Returns:
        Path to the saved workbook.

    Raises:
        FileNotFoundError: if ``main_file`` or any ``source`` file is missing.
        KeyError: if ``target_sheet`` does not exist in ``main_file``.
        ValueError: if the number of source and target columns differ.
    """
    if not os.path.isfile(main_file):
        raise FileNotFoundError(main_file)

    wb_main = load_workbook(main_file)
    total_mappings = len(mappings)

    try:
        for idx, mp in enumerate(mappings):
            src = mp.get("source")
            src_cols = mp.get("source_columns", [])
            tgt_sheet = mp.get("target_sheet")
            tgt_cols = mp.get("target_columns", [])

            if not os.path.isfile(src):
                raise FileNotFoundError(src)
            if len(src_cols) != len(tgt_cols):
                raise ValueError("Source and target columns must match in length")
            if tgt_sheet not in wb_main.sheetnames:
                raise KeyError(tgt_sheet)

            ws_main = wb_main[tgt_sheet]
            wb_src = load_workbook(src, read_only=True)
            ws_src = wb_src.active

            for s_col, t_col in zip(src_cols, tgt_cols):
                max_row = max(ws_src.max_row, ws_main.max_row)
                for row in range(1, max_row + 1):
                    ws_main[f"{t_col}{row}"].value = ws_src[f"{s_col}{row}"].value
            wb_src.close()

            if progress_callback:
                progress_callback(idx + 1, total_mappings, mp)

        if output_file is None:
            base, ext = os.path.splitext(main_file)
            output_file = base + "_merged" + ext
        wb_main.save(output_file)
        return output_file
    finally:
        wb_main.close()