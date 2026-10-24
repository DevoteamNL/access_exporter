# 05 – Remerge exported CSV/Excel files into the Access database

Status: **done** · Size: L · Live data: **YES – writes to an Access database** (preview + confirmation required)

## Goal

After editing exported CSV/XLSX files (in Excel or the built-in editor), write the changes back into the original
Access database (or a copy), table by table, with a preview of what will change.

## UI flow (English labels)

1. **Remerge into Access…** → choose the target `.accdb/.mdb` and the export file(s) (one XLSX or several CSVs).
2. **Mapping**: File/Sheet → Access table (auto-matched, editable dropdown); column check shows missing/extra
   columns and type problems.
3. **Mode**: *Update and add rows (by primary key)* / *Add rows only* / *Replace table contents*; checkbox
   *Also delete rows that are missing from the file* (update mode only, off by default).
4. **Preview** per table: "`a` added · `u` changed · `d` deleted · `e` errors" plus a sample of changed rows.
5. **Apply** → confirmation naming the database and backup file → progress → result in the log.

## Design

- Core `merge.py` (no PyQt):
  - `read_schema(conn, table)` – columns/types via `cursor.columns()`; primary key via
    `cursor.statistics(table, unique=True)` (Access ODBC `primaryKeys()` support is unreliable – **verify**).
  - `combine_split_sheets(frames)` – rejoins `Table_1`, `Table_2`, … from the Excel split.
  - `coerce(df, schema)` – text → Access types (dates, numbers with `.`/`,`, Yes/No, empty → NULL); per-row errors
    instead of failing the whole table.
  - `diff(source_df, access_rows, pk) -> MergePlan`; `apply(conn, plan)` with parameterised
    `INSERT`/`UPDATE`/`DELETE` (backtick quoting, `executemany`).
- Writing:
  - New `verbind_access_schrijfbaar(pad)` next to the read-only `verbind_access`; `autocommit=False`, one
    transaction per run; any error → `rollback()`.
  - **Before writing:** refuse if a `.laccdb` lock file exists (database open in Access); copy the database to
    `<name>.backup-YYYYMMDD-HHMMSS.accdb` (or write into a new copy, per PO answer).
- Runs in a `QThread` worker; **Stop** → rollback.
- Matching: manifest `<name>.export.json` (if approved) → name rules (reverse of `maak_bladnaam`/`maak_csv_naam`)
  → manual mapping.

## Not supported (shown as skipped columns)

Attachment, OLE object, multi-valued and calculated fields; AutoNumber columns are not written on insert.

## Steps

1. Export manifest (if approved) in `ConversionWorker._schrijf_excel/_schrijf_csv` + test.
2. `merge.py`: schema, combine, coerce + tests (pure pandas).
3. `merge.py`: diff + SQL generation + tests with a fake connection/cursor recording SQL and params.
4. Backup, lock-file check and transaction wrapper + tests (temp dir, fake connection asserting rollback).
5. `RemergeDialog` (mapping, mode, preview) + worker; translations NL/EN/FR; offscreen tests with mocked core.
6. Manual test against a throw-away `.accdb` (never committed), described in the PR.

## Tests

All offline: no real Access database in CI (`*.accdb` must never be committed; GitHub runners have no Access
driver).

## Risks

- **Data loss** → backup, preview, explicit confirmation, single transaction, rollback.
- Tables without primary key → only *Add rows only* / *Replace*.
- Type coercion edge cases (decimal comma, date formats, long text > 255) → per-row error report, nothing silent.
- Referential integrity: deletes/replace may fail on related tables → error shown, transaction rolled back.
- Large tables: compare in blocks by key hash to limit memory.

## PO decisions

- Lives in the **Merge back** tab.
- Always writes into a **new copy**: Save-as dialog pre-filled with `<name>_merged_YYYYMMDD-HHMMSS.accdb` next to the
  original; the original database is only opened read-only. (Replaces the backup step above.)
- Exporter writes the `<name>.export.json` manifest.
- Modes (developer decision): all three, update+add default, deletions opt-in.
