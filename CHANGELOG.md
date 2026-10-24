# Changelog

All notable changes to Access Exporter. Format based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versions follow [Semantic Versioning](https://semver.org/).

## [Unreleased]

## [1.2.0] - 2026-10-24

### Added

- **✏️ Edit** tab: edit Access tables or exported CSV/Excel files in the program, with search (Ctrl+F), add and
  delete rows, type checking per column and **Save as** (Access changes always go to a new, timestamped copy).
- **🔁 Merge back into Access** tab: merge edited CSV/Excel exports into a new, timestamped copy of the original
  database, with three modes (update and add by primary key, only add, replace), optional deletion, a preview with
  counts and errors, and an all-or-nothing write. The original database is never changed.
- Export manifest (`<export>.export.json`) so merging back fills in the target tables automatically.
- **Theme** selector: System (default, follows Windows), Light or Dark.
- New app icon in the Devoteam tool style, with a database symbol.
- Password-protected databases: the program asks for the password (kept in memory only).
- Drag and drop of `.accdb`/`.mdb` and CSV/Excel files onto the window.
- Option to also export saved queries.
- CSV decimal separator setting (decimal comma or point).
- **Clear finished jobs** button for the queue.
- French README, `CHANGELOG.md`, `build.ps1`, a test workflow for pull requests and a SHA256 checksum next to the exe
  in each release.

### Changed

- READMEs (English, Dutch, French) rewritten for non-technical users.
- Row counts use the thousands separator of the selected language.
- Code split into the `access_exporter` package; tests moved to pytest, including end-to-end tests against a
  temporary Access database.
- Requires `openpyxl` (reading Excel files) and PyQt6 6.5 or newer.

## [1.1.0] - 2026-10-19

### Added

- Interface language selector: Dutch (default), English and French, switching live and remembered between sessions.

## [1.0.0] - 2026-10-19

### Added

- First release: export Access tables to one Excel workbook or to CSV files per table, with a batch queue, table
  selection with preview, progress and log.

[Unreleased]: https://github.com/DevoteamNL/access_exporter/compare/v1.2.0...HEAD
[1.2.0]: https://github.com/DevoteamNL/access_exporter/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/DevoteamNL/access_exporter/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/DevoteamNL/access_exporter/releases/tag/v1.0.0
