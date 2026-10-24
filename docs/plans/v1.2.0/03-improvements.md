# 03 – Improvements (pick list)

Status: **approved** · Size: S–M (depends on selection) · Live data: no

Found while reviewing v1.1.0. The PO picks which items go into v1.2.0; **A** is recommended as a prerequisite for
plans 04 and 05.

| ID | Improvement | Why | Size |
|----|-------------|-----|------|
| A | **Split the single module** into `access_core.py` (pyodbc/pandas helpers, no PyQt), `worker.py`, `dialogs.py`, `main_window.py`; keep `access_exporter_pyqt6.py` as entry point | 960-line file; editor and remerge would push it past 2,000 lines; core becomes unit-testable | M |
| B | **Unit tests for core helpers** (`maak_bladnaam`, `maak_csv_naam`, `controleer_schrijfbaar`, block reading with a fake cursor) | Only language tests exist today | S |
| C | **CSV decimal separator** `decimal=","` (or a setting) | Files use `;` for Dutch Excel but numbers use `.` → Dutch Excel may read `1.5` as date/text | S |
| D | **Locale-aware thousands separator** in "Succesvol (… rijen)" | Hard-coded `.`; English UI should show `,` | S |
| E | **Password-protected databases**: ask for the password and pass `PWD=` | Now fails with an unclear ODBC error | S |
| F | **Drag & drop** `.accdb/.mdb` onto the window | Faster than the file dialog | S |
| G | **Export saved queries** (views) as an option next to tables | `haal_tabellen_op` only lists `TABLE` | S |
| H | **`build.ps1`** (as in iAsset Syncer) used by README and the release workflow | PyInstaller command duplicated in two places | S |
| I | **Clear finished jobs** button for the queue | Queue only shrinks via manual delete | S |

## Steps

Each selected item is one commit with its own test in `tests/`; new UI strings get NL/EN/FR translations; README
(EN + NL) updated where behaviour changes (C, E, F, G, H).

## Risks

- A: pure refactor; `import access_exporter_pyqt6 as exporter` in existing tests must keep working (re-export or
  update tests in the same commit).
- C: changes CSV output; scripts parsing the CSV may break → make it a setting or note it in the release notes.

## Open question

- Which of A–I go into v1.2.0?
