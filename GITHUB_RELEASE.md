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
