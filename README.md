# Access Exporter

<img src="access_exporter/assets/access-exporter.png" alt="Access Exporter icon" width="96" align="right">

**English** | [Nederlands](README.nl.md) | [Français](README.fr.md)

## What does it do?

Access Exporter is a small Windows program to get data **out of** a Microsoft Access database (`.accdb` / `.mdb`)
and, if you want, **back in** again. You don't need Microsoft Access itself, and you don't need any technical
knowledge.

- **Export** – turn the tables of one or more Access databases into an Excel workbook or into CSV files, for example
  to review, clean up or share the data with a customer.
- **Edit** – open an Access database or an exported Excel/CSV file and change values right in the program, like in a
  spreadsheet.
- **Merge back into Access** – put an edited Excel/CSV export back into the database.

**Your original database is never changed.** Every edit or merge is saved as a *new copy* of the database, with the
date and time in its name (for example `Customers_merged_20260520-143000.accdb`). If something does not look right,
simply delete the copy and start again.

Typical use: export a customer database to Excel, let the customer correct the data, merge the corrected file back
into a copy of the database and hand that copy over.

## Contents

- [Installation](#installation)
- [Export tables](#export-tables)
- [Edit data](#edit-data)
- [Merge back into Access](#merge-back-into-access)
- [Appearance and language](#appearance-and-language)
- [Good to know](#good-to-know)
- [Troubleshooting](#troubleshooting)
- [For developers](#for-developers)
- [License](#license)

## Installation

1. Download `AccessExporter.exe` from the
   [latest release](https://github.com/DevoteamNL/access_exporter/releases/latest). No installation is needed;
   you can start it from any folder. Windows may show a SmartScreen warning the first time: choose
   **More info → Run anyway**.
2. The program needs the free **Microsoft Access Database Engine** (64-bit), the part of Access that reads database
   files. Many Office installations already include it. If it is missing, the program tells you at startup; install
   it from [Microsoft](https://www.microsoft.com/en-us/download/details.aspx?id=54920) (`accessdatabaseengine_X64.exe`).

To check the download, compare its SHA256 checksum with the `AccessExporter.exe.sha256` file in the release:
`Get-FileHash .\AccessExporter.exe` in PowerShell.

## Export tables

Use the **📤 Export** tab.

1. **📂 1. Choose Access file** – or drag a `.accdb`/`.mdb` file onto the window. The tables are loaded and all
   selected. Tick **Include saved queries** to also export the saved queries of the database.
2. **⚙ 2. Choose export format** – **📊 One Excel file (.xlsx)** with one sheet per table, or
   **📄 Separate CSV files per table (.csv)**. For CSV you can choose **Decimal comma (1,5)** (Dutch/European Excel)
   or **Decimal point (1.5)**.
3. **📋 3. Choose tables to export** (optional) – filter, tick and preview the first rows of each table.
4. **📁 4. Choose output folder** – by default the folder of the Access file.
5. **➕ 5. Add to queue**, then **▶️ Convert selection** or **🚀 Convert entire queue**.

You can queue several databases, each with its own settings. Progress and messages appear in **Log - Live status**.
**🧹 Clear finished jobs** cleans up the queue; <kbd>Delete</kbd> removes the selected rows. Tick
**Open folder when finished** to open the output folder automatically.

Next to each export the program writes a small `.export.json` file. It remembers which database and which table
every sheet or CSV file came from, so merging back later can fill in the right tables for you. Keep it next to the
export if you plan to merge back.

Details: tables larger than Excel's limit (1,048,576 rows) are spread over several sheets (`Table_1`, `Table_2`, …);
CSV files use `;` and UTF-8, so they open correctly in Excel; text such as `=SUM(A1)` stays plain text.

## Edit data

Use the **✏️ Edit** tab.

1. Click **🗄 Open Access database...** or **📂 Open CSV/Excel file...** (or drag a file onto the tab).
2. Pick a table or sheet in **Table / sheet:** and change cells by double-clicking them. Use
   **🔍 Search all columns... (Ctrl+F)** to find rows, **➕ Add row** and **➖ Delete rows** (or <kbd>Delete</kbd>).
   Values that do not fit the column (for example text in a number column) are refused with a message.
3. Click **💾 Save as...** and choose a name. A name with the date and time is already filled in.
   - For an Access database, all changes are written to a **new copy** of the database, which is then opened in the
     editor.
   - For an Excel file, all sheets are saved; for a CSV file, the current sheet is saved.

If you close the program or open another file with unsaved changes, you are asked whether to discard them.

## Merge back into Access

Use the **🔁 Merge back into Access** tab.

1. **1. Original Access database** – **🗄 Choose database...** (the database the export came from).
2. **2. Edited files and target tables** – **📂 Add files...** with your edited Excel/CSV files. For each sheet or
   file, check the **Target table in Access** (filled in automatically when possible) or choose *(skip)*.
3. **3. How to merge?**
   - **Update and add (by primary key)** – changed rows are updated, new rows are added. Recommended.
   - **Only add new rows** – existing rows are left alone.
   - **Replace table contents** – the table is emptied and filled with the file.
   - **Also delete rows that are no longer in the file** – optional, only with *Update and add*.
4. **4. Review and merge** – **🔍 Calculate preview** shows per table how many rows will be added, changed, deleted
   or left unchanged, and any errors. Then click **💾 Merge into new copy...**, confirm the name of the new copy and
   the summary.

All changes are written in one go: if anything fails or you click **⏹ Stop**, nothing is kept and the unfinished
copy is removed.

## Appearance and language

At the top of the window:

- **Theme** – **System** (follows Windows light/dark mode, the default), **Light** or **Dark**.
- **Language** – **Nederlands** (default), **English** or **Français**. Texts, messages and the log switch
  immediately; your data is never translated.

Both choices are remembered, together with your last folders and export settings (stored in the Windows registry
under `HKEY_CURRENT_USER\Software\AccessExporter\AccessExporterPyQt6`).

## Good to know

- **Password-protected databases** – you are asked for the password. It is kept only while the program is open and
  never saved.
- **Tables without a primary key** are shown read-only in the editor. When merging back, use *Only add new rows* or
  *Replace table contents* for them.
- **Large tables** – the editor loads at most 100,000 rows per table; rows beyond that are not touched when you save.
- **Attachments, pictures and other binary columns** are not exported and are left unchanged when merging back.
- **AutoNumber columns** are filled in by Access for new rows.
- Close the database in Microsoft Access before merging or saving; the program warns you when it appears to be open.

## Troubleshooting

| Message or problem | What to do |
|--------------------|------------|
| *Access driver missing* at startup | Install the 64-bit Microsoft Access Database Engine (see [Installation](#installation)). When running from source, its bitness (32/64-bit) must match Python. |
| *Database in use* | Close the database in Access (a `.laccdb`/`.ldb` file next to it means it is open) and try again. |
| Wrong numbers in Excel after a CSV export | Choose the other decimal setting in step 2. |
| A row shows errors in the merge preview | The value does not fit the column type (for example text in a date column). Fix it in the file or in the **✏️ Edit** tab and calculate the preview again. |

## For developers

Requirements: Windows, Python 3.10+, the Access Database Engine with the same bitness as Python.

```powershell
git clone https://github.com/DevoteamNL/access_exporter.git
cd access_exporter
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt pytest
.venv\Scripts\python.exe access_exporter_pyqt6.py        # run the app
.venv\Scripts\python.exe -m pytest -q tests              # run the tests
powershell -ExecutionPolicy Bypass -File .\build.ps1   # build dist\AccessExporter.exe
```

- Code lives in the `access_exporter` package: `core.py` (database access, no Qt), `file_io.py` (Excel/CSV),
  `merge.py` (comparing and writing back), `main_window.py`, `editor_tab.py`, `merge_tab.py`, `themes.py`,
  `translations.py`. `access_exporter_pyqt6.py` is the entry point.
- Tests that need the Access driver create a temporary database and are skipped when the driver is missing; the
  other tests use Qt's offscreen platform.
- UI texts are written in Dutch; English and French live in `TRANSLATIONS` in `translations.py`.
- The icon is drawn in `access_exporter/assets/access-exporter.svg`; run `python tools/generate_app_icon.py` to
  regenerate the `.ico` and `.png`.
- Publishing a GitHub release builds the exe and uploads it with its checksum (`.github/workflows/release-assets.yml`).
  Changes are listed in [CHANGELOG.md](CHANGELOG.md).

## License

[GNU General Public License v3.0](LICENSE) (GPL-3.0-only), matching the license of PyQt6, which the program
includes.
