# Changelog

## 2026-07-07 — repository audit

- Confirmed current visible code uses example default path rather than a real internal network path.
- Classified project as an internal Excel automation utility; public release requires sanitized fixtures and full-history data/secret scan.
- Documented current version as the canonical size sorter/watcher.

## v2.11

- Watches invoice folders for `.xls`, `.xlsx` and `.xlsm` files.
- Converts `.xls` to `.xlsx` via Excel COM on Windows.
- Detects size columns and normalizes them to `ORDERED_SIZES`.
- Saves normalized files with `__normalized` suffix and tracks processed files in JSON.
