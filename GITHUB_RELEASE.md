# xlMerger `15.04.2026`

Release date: `2026-04-15`

## Highlights

- Improved `xlSplit` UX in split configuration dialog:
  - added explicit helper text for selection:
    - left mouse button: select,
    - right mouse button: unselect.
- Updated `xlSplit` export behavior for source-only setup:
  - when only source column is selected and only extra columns are checked, app now exports a single Excel file,
  - exported column order is now:
    - selected extra columns first,
    - source column last.
- Switched key save paths from `openpyxl.save(...)` to `xlsxwriter` export with formatting preservation:
  - `core/excel_processor.py`,
  - `gui/limits_checker.py`,
  - `excel_builder/executor.py`.
- Added centralized exporter `utils/xlsxwriter_export.py` to preserve:
  - cell values and styles,
  - column widths / row heights,
  - merged ranges,
  - freeze panes / autofilter / sheet view flags.
- Added startup warning filter for:
  - `Workbook contains no default style, apply openpyxl's default`
  in `main.py`.

## Verification

- Full test suite executed and passed:
  - `34 passed in 5.34s`.
- Added regression/coverage tests:
  - `tests/test_xlsxwriter_export.py`,
  - new formatting-preservation test in `tests/test_excel_processor_copy_mode.py`,
  - updated split behavior tests in:
    - `tests/test_split_excel.py`,
    - `tests/test_split_mapping_dialog.py`.

---

# xlMerger `11.04.2026`

Release date: `2026-04-11`

## Highlights

- Removed `xlCombine` from the application and cleaned related modules/tests.
- Improved `xlMerger` UX:
  - required copy-column placeholder,
  - disabled + gray `Start` button when translation column is not selected,
  - clear validation message above the button.
- Fixed the warning label clipping in English UI.
- Replaced hardcoded icon path with a portable project-relative path.
- Expanded RU/EN localization entries and added translation coverage test.

## Breaking Changes

- `xlCombine` feature has been removed from this release.

## Verification

- Targeted UI/logics tests were executed locally and passed.
- Translation coverage test is included to prevent missing `tr(...)` keys in RU/EN.
