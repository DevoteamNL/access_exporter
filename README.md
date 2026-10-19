# Access Exporter

**[English](#english) | [Nederlands](#nederlands)**

---

## English

A Windows desktop tool (PyQt6) that exports tables from Microsoft Access databases (`.accdb` / `.mdb`) to **one Excel workbook** or **separate CSV files per table**, using a queue so multiple databases can be converted in one go.

### Features

- **Batch queue** – add several Access files, each with its own table selection, export format and output folder, then convert the selection or the entire queue.
- **Table selection with preview** – filter tables by name, check/uncheck all, and preview the first 10 rows of each table before exporting.
- **Excel export (`.xlsx`)** – one sheet per table. Tables larger than Excel's row limit (1,048,576 rows) are split automatically over multiple sheets (`Table_1`, `Table_2`, …). Sheet names are sanitised and kept unique, and text such as `=SUM(A1)` or `http://…` stays plain text.
- **CSV export** – one file per table, semicolon-separated, UTF-8 with BOM (opens correctly in Dutch/European Excel).
- **Memory-friendly** – data is read and written in blocks of 50,000 rows, so very large tables don't need to fit in memory.
- **Safe** – the source database is opened read-only; a running conversion can be stopped, and partially written files are removed.
- **Live progress & log** – per-file progress based on rows processed, colour-coded status and a timestamped log.
- **Remembers settings** – output folder, export format and last used folders are saved between sessions.
- Dark theme.

### Requirements

- Windows
- Python 3.10 or newer
- [Microsoft Access Database Engine](https://www.microsoft.com/en-us/download/details.aspx?id=54920) (ODBC driver *"Microsoft Access Driver (\*.mdb, \*.accdb)"*). **Its bitness (32/64-bit) must match your Python installation.** The app warns at startup if the driver is missing.

### Installation

**Option 1 – Standalone executable (no Python needed):** download `AccessExporter.exe` from the [latest release](https://github.com/cramor-devo/access_exporter/releases/latest) and run it. The 64-bit Microsoft Access Database Engine is still required.

**Option 2 – From source:**

```powershell
git clone https://github.com/cramor-devo/access_exporter.git
cd access_exporter
python -m pip install -r requirements.txt
```

### Usage

```powershell
python access_exporter_pyqt6.py
```

1. **Choose Access file** – the tables are loaded automatically (all selected by default).
2. **Choose export format** – one Excel file or separate CSV files.
3. **Choose tables** (optional) – select tables and preview data.
4. **Choose output folder** – defaults to the folder of the Access file.
5. **Add to queue**, then click **Convert Selection** or **Convert Entire Queue**.

Tip: press <kbd>Delete</kbd> to remove selected rows from the queue.

> The interface is in Dutch.

---

## Nederlands

Een Windows-desktopprogramma (PyQt6) dat tabellen uit Microsoft Access-databases (`.accdb` / `.mdb`) exporteert naar **één Excel-werkmap** of **aparte CSV-bestanden per tabel**. Dankzij een wachtrij kun je meerdere databases in één keer converteren.

### Functies

- **Wachtrij** – voeg meerdere Access-bestanden toe, elk met een eigen tabelselectie, exportformaat en doelmap, en converteer daarna de selectie of de hele wachtrij.
- **Tabellen selecteren met voorbeeld** – filter tabellen op naam, vink alles aan/uit en bekijk de eerste 10 rijen van elke tabel vóór het exporteren.
- **Excel-export (`.xlsx`)** – één werkblad per tabel. Tabellen die groter zijn dan de Excel-limiet (1.048.576 rijen) worden automatisch over meerdere bladen verdeeld (`Tabel_1`, `Tabel_2`, …). Bladnamen worden opgeschoond en uniek gemaakt, en tekst zoals `=SUM(A1)` of `http://…` blijft gewone tekst.
- **CSV-export** – één bestand per tabel, puntkomma als scheidingsteken, UTF-8 met BOM (opent correct in de Nederlandse Excel).
- **Zuinig met geheugen** – gegevens worden in blokken van 50.000 rijen gelezen en geschreven, zodat zeer grote tabellen niet in het geheugen hoeven te passen.
- **Veilig** – de bron-database wordt alleen-lezen geopend; een lopende conversie kan worden gestopt en gedeeltelijk geschreven bestanden worden opgeruimd.
- **Live voortgang & logboek** – voortgang per bestand op basis van verwerkte rijen, statuskleuren en een logboek met tijdstempels.
- **Onthoudt instellingen** – doelmap, exportformaat en laatst gebruikte mappen worden tussen sessies bewaard.
- Donker thema.

### Vereisten

- Windows
- Python 3.10 of nieuwer
- [Microsoft Access Database Engine](https://www.microsoft.com/nl-nl/download/details.aspx?id=54920) (ODBC-driver *"Microsoft Access Driver (\*.mdb, \*.accdb)"*). **De bitness (32/64-bit) moet gelijk zijn aan die van je Python-installatie.** Het programma waarschuwt bij het opstarten als de driver ontbreekt.

### Installatie

**Optie 1 – Los programma (geen Python nodig):** download `AccessExporter.exe` uit de [nieuwste release](https://github.com/cramor-devo/access_exporter/releases/latest) en start het. De 64-bit Microsoft Access Database Engine is nog steeds vereist.

**Optie 2 – Vanaf de broncode:**

```powershell
git clone https://github.com/cramor-devo/access_exporter.git
cd access_exporter
python -m pip install -r requirements.txt
```

### Gebruik

```powershell
python access_exporter_pyqt6.py
```

1. **Kies Access-bestand** – de tabellen worden automatisch geladen (standaard allemaal geselecteerd).
2. **Kies exportformaat** – één Excel-bestand of aparte CSV-bestanden.
3. **Kies tabellen** (optioneel) – selecteer tabellen en bekijk een datavoorbeeld.
4. **Kies doelmap** – standaard de map van het Access-bestand.
5. **Voeg toe aan wachtrij** en klik daarna op **Converteer Selectie** of **Converteer Hele Wachtrij**.

Tip: druk op <kbd>Delete</kbd> om geselecteerde rijen uit de wachtrij te verwijderen.
