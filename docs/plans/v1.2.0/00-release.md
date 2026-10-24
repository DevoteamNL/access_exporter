# Release plan – Access Exporter v1.2.0

Status: **done**
GitHub issue: #2

## Request (PO, verbatim)

> create a feature to remerge exported csv/excel files with the original accdb and/or a file editor in the
> application itself. also check for improvements, that the light and dark theme can be toggled and that the readme
> files are complete. also create an icon in the same style as my other programs like iasset syncer, but with the
> place marker being replaced with a database icon.

## Current state (v1.1.0)

- One module `access_exporter_pyqt6.py` (~960 lines: helpers, `ConversionWorker`, `TableSelectionDialog`,
  `AccessExporterGUI`) plus `translations.py` (NL source → EN/FR). Tests: `tests/test_languages.py` only.
- Source database is always opened **read-only** (`verbind_access(..., readonly=True)`).
- Dark theme only (`AccessExporterGUI.apply_stylesheet`), colours also hard-coded in constants (`KLEUR_*`,
  `STATUS_KLEUREN` default `#E0E0E0`).
- No application icon; PyInstaller command in README and `.github/workflows/release-assets.yml` has no `--icon`.

## Features

| # | Plan | Size | Depends on | Changes live data? |
|---|------|------|------------|--------------------|
| 01 | [App icon](01-app-icon.md) | S | – | No |
| 02 | [Light/dark theme toggle](02-theme-toggle.md) | M | – | No |
| 03 | [Improvements (pick list)](03-improvements.md) | S–M | – | No |
| 04 | [Built-in file editor (CSV/XLSX)](04-file-editor.md) | M–L | 02 (colours), 03-A (module split) | Edits export files only |
| 05 | [Remerge exports into the Access database](05-remerge.md) | L | 04 (shared file reader), 03-A | **Yes – writes to .accdb** |
| 06 | [README completeness (EN + NL)](06-readme.md) | S | all above | No |

## Recommended order

1. **01 App icon** – isolated, low risk, immediately visible.
2. **02 Theme toggle** – needed before new dialogs so they get both palettes from day one.
3. **03 Improvements** – only the items the PO selects; includes the module split that 04/05 build on.
4. **04 File editor** – introduces the shared CSV/XLSX reader and table model.
5. **05 Remerge** – last, because it is the only feature that writes to a database.
6. **06 README** – quick fix of current gaps first, final pass after 04/05 are merged.

One PR per plan, each linking its issue.

## Proposed version

**1.2.0** (new backwards-compatible features). If 05 needs more time, ship 01–04 as 1.2.0 and 05 as 1.3.0.

## PO decisions (2026-10-24)

1. **Issues:** one umbrella issue for the whole v1.2.0 release.
2. **Theme (02):** `System / Light / Dark`, default **System**.
3. **Improvements (03):** all items A–I.
4. **Editor (04):** edits **Access tables directly** (besides exported CSV/XLSX files); saving is always **Save as**
   (a new file/copy, the original is never overwritten).
5. **Remerge (05):** always writes into a **new copy** of the database, with a pre-filled timestamped file name
   (`<name>_merged_YYYYMMDD-HHMMSS.accdb`) in a Save-as dialog.
6. **Manifest (05):** yes, the exporter writes `<name>.export.json` next to the exports.
7. **Navigation:** tabs **Export** / **Editor** / **Merge back** in the main window.
8. **README:** add French (`README.fr.md`), switch links to `DevoteamNL/access_exporter`, rewrite all READMEs for
   legibility with an introduction for non-developers (e.g. consultants).

## Technical decisions (developer, not asked)

- Icon database colour: same orange `#ff9854` as the iAsset Syncer pin.
- Merge modes: *Update and add rows (by primary key)* (default), *Add rows only*, *Replace table contents*; option
  *Also delete rows missing from the file* off by default. Safe because the original database is never changed.
- Access-table editing (04) and merge (05) share one write path: copy the database to the Save-as target, then
  apply changes in one transaction on the copy.
- README split: `README.md` (EN, with language links), `README.nl.md`, `README.fr.md`.

## Outcome (implementation notes)

All plans 01–06 and improvements A–I are implemented. Deviations from the plans:

- A: the code became the package `access_exporter/` (`core.py`, `file_io.py`, `merge.py`, `worker.py`, `dialogs.py`,
  `main_window.py`, `editor_tab.py`, `merge_tab.py`, `themes.py`, `translations.py`) instead of flat modules;
  `access_exporter_pyqt6.py` stays the entry point and re-exports the old names.
- 04: per PO decision the editor also edits Access databases directly; saving always writes a new, timestamped copy.
- 05: the manifest is written as `<export>.xlsx.export.json` / `<export>.csv.export.json` next to the export.
- 06: a French README (`README.fr.md`) was added next to the English and Dutch ones.
- A `Tests` workflow now runs on pull requests; the release workflow uses pytest and `build.ps1`, and uploads a
  SHA256 checksum.
