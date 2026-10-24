# 04 – Built-in file editor (CSV/XLSX)

Status: **approved** · Size: M–L · Live data: edits export files only (never the `.accdb`)

## Goal

Open an exported CSV or Excel file inside the app, view and edit cells, add/delete rows, search and save – without
Excel. Also the step before remerging (plan 05).

## UI (English labels)

- Entry points: button **Edit file…** in the main window; **Open in editor** in the queue context menu /
  double-click on a finished job.
- Editor window: toolbar **Open…**, **Save**, **Save as…**, **Add row**, **Delete rows**, **Find** (`Ctrl+F`);
  sheet selector (XLSX); status bar "`n` rows · `m` columns · modified".
- Closing with unsaved changes → "Save changes to <file>?" **Save / Discard / Cancel**.

## Design

- Core (no PyQt) `file_io.py`:
  - `read_export(path) -> dict[str, DataFrame]` – CSV: `sep=";"`, `encoding="utf-8-sig"`, `dtype=str`,
    `keep_default_na=False`; XLSX: `pd.read_excel(sheet_name=None, dtype=object)` (**needs `openpyxl`**).
  - `write_export(path, frames)` – same rules as the exporter (xlsxwriter without formulas/urls, `;` CSV with BOM),
    written to a temp file then `os.replace` (atomic).
- UI `editor_window.py`: `DataFrameModel(QAbstractTableModel)` (editable, insert/remove rows) + `QTableView`
  (virtualised, 100k+ rows stay responsive) + `QSortFilterProxyModel` for find.
- Files above 500,000 rows: warning before loading.

## Steps

1. `requirements.txt`: add `openpyxl>=3.1`; PyInstaller `--hidden-import openpyxl` (README + workflow).
2. `file_io.py` read/write + tests.
3. `DataFrameModel` + tests (edit, insert, remove, dirty flag).
4. `EditorWindow` (toolbar, sheet selector, find, save/close flow); translations NL/EN/FR.
5. Entry points in the main window.

## Tests

- `tests/test_file_io.py`: CSV/XLSX round-trip keeps values, leading zeros, `=SUM(..)` as text, empty cells; failed
  write leaves the original intact.
- `tests/test_editor.py` (offscreen): edit → dirty → save → file changed; unsaved-close prompt (patched
  `QMessageBox`).

## Risks

- Saving an XLSX rewrites it: formatting added in Excel is lost → note in README and tooltip.
- File locked by Excel → reuse `controleer_schrijfbaar` with a clear message.
- New dependency `openpyxl` → slightly larger exe.

## PO decisions

- Editor lives in the **Editor** tab and opens **Access tables directly** (table list of the chosen `.accdb/.mdb`,
  read-only connection) as well as exported CSV/XLSX files.
- Saving is always **Save as**: for files a new CSV/XLSX; for Access a new database copy with pre-filled name
  `<name>_edited_YYYYMMDD-HHMMSS.accdb`, then changes are applied to the copy in one transaction (shared write path
  with plan 05: `merge.apply`, keyed by primary key; tables without primary key are read-only in the editor).
- Access tables are loaded with a row limit (default 100,000, warning above) to keep the UI responsive.
