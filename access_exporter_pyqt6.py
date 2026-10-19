"""Access (.accdb/.mdb) naar Excel/CSV exporter met een PyQt6-interface.

Vereisten: pip install PyQt6 pandas pyodbc xlsxwriter
(openpyxl wordt als terugval gebruikt als xlsxwriter ontbreekt.)
Daarnaast is de "Microsoft Access Driver (*.mdb, *.accdb)" nodig, met dezelfde
bitness (32/64-bit) als de gebruikte Python-installatie.
"""

__version__ = "1.1.0"

import math
import os
import re
import subprocess
import sys
from collections.abc import Iterator
from contextlib import closing
from datetime import datetime
from importlib.util import find_spec
from typing import Any

import pandas as pd
import pyodbc
from PyQt6.QtCore import QSettings, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QCloseEvent, QColor, QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog, QDialogButtonBox,
    QFileDialog, QGridLayout, QGroupBox, QHBoxLayout, QHeaderView, QLabel,
    QLineEdit, QListWidget, QListWidgetItem, QMainWindow, QMessageBox,
    QPlainTextEdit, QProgressBar, QPushButton, QRadioButton, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget
)

from translations import language_support, tr

# =========================================================
# CONSTANTEN & HULPFUNCTIES
# =========================================================
ACCESS_DRIVER = "Microsoft Access Driver (*.mdb, *.accdb)"
EXCEL_ENGINE = next((m for m in ("xlsxwriter", "openpyxl") if find_spec(m)), None)
EXCEL_MAX_DATA_RIJEN = 1_048_575  # 1.048.576 rijen per blad, min de kopregel
EXCEL_MAX_BLADNAAM = 31
BLOKGROOTTE = 50_000  # Aantal rijen dat per keer in het geheugen wordt geladen
PREVIEW_RIJEN = 10

STATUS_WACHTRIJ = "In de wachtrij"
STATUS_BEZIG = "Bezig"
STATUS_KLAAR = "Klaar"
STATUS_FOUT = "Fout"
STATUS_GESTOPT = "Gestopt"

KLEUR_OK = "#4CAF50"
KLEUR_FOUT = "#F44336"
KLEUR_INFO = "#569cd6"
KLEUR_BEZIG = "#FFB74D"
STATUS_KLEUREN = {STATUS_KLAAR: KLEUR_OK, STATUS_FOUT: KLEUR_FOUT, STATUS_BEZIG: KLEUR_BEZIG, STATUS_GESTOPT: KLEUR_BEZIG}

_ONGELDIGE_BLADTEKENS = re.compile(r"[\[\]:*?/\\]")


def access_driver_beschikbaar() -> bool:
    return ACCESS_DRIVER in pyodbc.drivers()


def verbind_access(pad: str) -> pyodbc.Connection:
    # Alleen-lezen: de bron-database wordt nooit gewijzigd
    return pyodbc.connect(f"DRIVER={{{ACCESS_DRIVER}}};DBQ={pad};", readonly=True)


def quote_tabel(naam: str) -> str:
    # Backticks i.p.v. [] zodat een '?' in de naam niet als ODBC-parameter wordt gezien
    return f"`{naam}`"


def haal_tabellen_op(pad: str) -> list[str]:
    with closing(verbind_access(pad)) as conn:
        return [t.table_name for t in conn.cursor().tables(tableType="TABLE")
                if not t.table_name.startswith("MSys")]


def tel_rijen(conn: pyodbc.Connection, tabel: str) -> int:
    rij = conn.cursor().execute(f"SELECT COUNT(*) FROM {quote_tabel(tabel)}").fetchone()
    return int(rij[0]) if rij else 0


def lees_in_blokken(conn: pyodbc.Connection, tabel: str, blokgrootte: int = BLOKGROOTTE) -> Iterator[pd.DataFrame]:
    """Leest een tabel in blokken in, zodat grote tabellen niet in één keer in het geheugen staan.
    Een lege tabel levert één leeg DataFrame op (met kolomnamen)."""
    with closing(conn.cursor()) as cursor:
        cursor.execute(f"SELECT * FROM {quote_tabel(tabel)}")
        kolommen = [c[0] for c in cursor.description]
        leeg = True
        while rijen := cursor.fetchmany(blokgrootte):
            leeg = False
            df = pd.DataFrame.from_records([tuple(r) for r in rijen], columns=kolommen, coerce_float=True)
            # Nullable types: een getalkolom met lege waarden blijft in elk blok een gehele-getallenkolom
            yield df.convert_dtypes()
        if leeg:
            yield pd.DataFrame(columns=kolommen)


def maak_bladnaam(naam: str, gebruikt: set[str], achtervoegsel: str = "") -> str:
    """Geldige, unieke Excel-bladnaam (max. 31 tekens, geen []:*?/\\, niet hoofdlettergevoelig uniek)."""
    basis = _ONGELDIGE_BLADTEKENS.sub("_", naam).strip("'") or "Blad"
    if basis.lower() == "history":  # Gereserveerd door Excel
        basis += "_"
    kandidaat = basis[:EXCEL_MAX_BLADNAAM - len(achtervoegsel)] + achtervoegsel
    teller = 2
    while kandidaat.lower() in gebruikt:
        extra = f"~{teller}{achtervoegsel}"
        kandidaat = basis[:EXCEL_MAX_BLADNAAM - len(extra)] + extra
        teller += 1
    gebruikt.add(kandidaat.lower())
    return kandidaat


def maak_csv_naam(prefix: str, tabel: str, gebruikt: set[str]) -> str:
    veilig = "".join(x for x in tabel if x.isalnum() or x in " _-").strip() or "tabel"
    basis = f"{prefix}_{veilig}"
    naam, teller = basis, 2
    while naam.lower() in gebruikt:
        naam = f"{basis}_{teller}"
        teller += 1
    gebruikt.add(naam.lower())
    return f"{naam}.csv"


def controleer_schrijfbaar(pad: str) -> None:
    """Faalt direct als een bestaand doelbestand vergrendeld is (bijv. geopend in Excel)."""
    if os.path.exists(pad):
        try:
            open(pad, "r+b").close()
        except PermissionError:
            raise PermissionError(f"Doelbestand is in gebruik of alleen-lezen: {os.path.basename(pad)}") from None


def open_in_verkenner(pad: str) -> None:
    pad = os.path.normpath(pad)
    if os.path.isfile(pad):
        subprocess.Popen(f'explorer /select,"{pad}"')
    else:
        subprocess.Popen(["explorer", pad])


class ConversieGestopt(Exception):
    pass


# =========================================================
# WORKER THREAD: Voert de zware conversies uit op de achtergrond
# =========================================================
class ConversionWorker(QThread):
    # Signalen om veilig met de GUI te communiceren
    update_table_signal = pyqtSignal(int, str, str, str)  # rij, voortgang, status, melding
    log_signal = pyqtSignal(str)
    job_done_signal = pyqtSignal()

    def __init__(self, taken: list[tuple[int, dict[str, Any]]], open_map: bool):
        super().__init__()
        self.taken = taken
        self.open_map = open_map
        self._rij = 0
        self._rijen_totaal = 0
        self._rijen_klaar = 0
        self._laatste_pct = -1
        self._max_pct = 100

    def run(self) -> None:
        for rij, job in self.taken:
            if self.isInterruptionRequested():
                self.log_signal.emit("⏹ Conversie gestopt; resterende taken blijven in de wachtrij.")
                break
            self._verwerk_job(rij, job)
            self.job_done_signal.emit()

    def _controleer_stop(self) -> None:
        if self.isInterruptionRequested():
            raise ConversieGestopt()

    def _voortgang(self, extra_rijen: int, melding: str) -> None:
        self._rijen_klaar += extra_rijen
        pct = self._max_pct if not self._rijen_totaal else int(self._rijen_klaar / self._rijen_totaal * self._max_pct)
        if pct != self._laatste_pct:
            self._laatste_pct = pct
            self.update_table_signal.emit(self._rij, f"{pct}%", STATUS_BEZIG, melding)

    def _verwerk_job(self, rij: int, job: dict[str, Any]) -> None:
        naam = job["access_naam"]
        export_type = job["export_type"]
        tabellen: list[str] = job["tabellen"]
        aangemaakt: list[str] = []

        self._rij, self._rijen_klaar, self._laatste_pct = rij, 0, -1
        self.update_table_signal.emit(rij, "0%", STATUS_BEZIG, "Verbinding maken...")
        self.log_signal.emit(f"Start verwerking: {naam} (Formaat: {export_type.upper()})")

        try:
            with closing(verbind_access(job["access_pad"])) as conn:
                self.update_table_signal.emit(rij, "0%", STATUS_BEZIG, "Rijen tellen...")
                tellingen = {t: tel_rijen(conn, t) for t in tabellen}
                self._rijen_totaal = sum(tellingen.values())

                if export_type == "xlsx":
                    resultaat = self._schrijf_excel(conn, job, tellingen, aangemaakt)
                elif export_type == "csv":
                    resultaat = self._schrijf_csv(conn, job, aangemaakt)
                else:
                    raise ValueError(f"Onbekend exportformaat: {export_type}")

        except ConversieGestopt:
            self._ruim_op(aangemaakt)
            self.update_table_signal.emit(rij, "0%", STATUS_GESTOPT, "Door gebruiker gestopt")
            self.log_signal.emit(f"⏹ Gestopt: {naam} (gedeeltelijke bestanden verwijderd)")
            return
        except Exception as e:
            self._ruim_op(aangemaakt)
            foutreden = str(e).split("\n")[0]
            self.update_table_signal.emit(rij, "0%", STATUS_FOUT, foutreden)
            self.log_signal.emit(f"❌ Fout bij {naam}: {foutreden}")
            return

        self.update_table_signal.emit(rij, "100%", STATUS_KLAAR, f"Succesvol ({self._rijen_totaal:,} rijen)".replace(",", "."))
        self.log_signal.emit(f"✅ Conversie succesvol afgerond: {naam}")
        if self.open_map:
            open_in_verkenner(resultaat)

    def _schrijf_excel(self, conn: pyodbc.Connection, job: dict[str, Any], tellingen: dict[str, int], aangemaakt: list[str]) -> str:
        if EXCEL_ENGINE is None:
            raise RuntimeError("Geen Excel-bibliotheek gevonden. Installeer met: pip install xlsxwriter")

        excel_pad = os.path.join(job["doelmap"], f"{job['naam_zonder_ext']}.xlsx")
        controleer_schrijfbaar(excel_pad)
        tabellen: list[str] = job["tabellen"]
        totaal = len(tabellen)
        gebruikte_bladen: set[str] = set()
        # Tekst als '=...' of 'http://...' moet als tekst blijven, niet als formule/hyperlink
        engine_kwargs = {"options": {"strings_to_formulas": False, "strings_to_urls": False}} if EXCEL_ENGINE == "xlsxwriter" else {}
        self._max_pct = 95  # Laatste stap is het opslaan van het bestand

        aangemaakt.append(excel_pad)
        with pd.ExcelWriter(excel_pad, engine=EXCEL_ENGINE, engine_kwargs=engine_kwargs) as writer:
            for index, tabel in enumerate(tabellen, 1):
                self._controleer_stop()
                self.log_signal.emit(f"-> Inlezen tabel '{tabel}' ({index}/{totaal})...")
                aantal_delen = max(1, math.ceil(tellingen[tabel] / EXCEL_MAX_DATA_RIJEN))
                if aantal_delen > 1:
                    self.log_signal.emit(f"   (Te groot, opsplitsen in {aantal_delen} delen...)")

                bladen: list[str] = []
                blad_idx = 0
                rijen_in_blad = 0
                for df in lees_in_blokken(conn, tabel):
                    self._controleer_stop()
                    if df.empty:  # Lege tabel: alleen de kopregel wegschrijven
                        df.to_excel(writer, sheet_name=maak_bladnaam(tabel, gebruikte_bladen), index=False)
                        continue
                    start = 0
                    while start < len(df):
                        if rijen_in_blad == EXCEL_MAX_DATA_RIJEN:
                            blad_idx += 1
                            rijen_in_blad = 0
                        if blad_idx == len(bladen):
                            splitsen = aantal_delen > 1 or blad_idx > 0
                            bladen.append(maak_bladnaam(tabel, gebruikte_bladen, f"_{blad_idx + 1}" if splitsen else ""))
                        deel = df.iloc[start:start + EXCEL_MAX_DATA_RIJEN - rijen_in_blad]
                        deel.to_excel(
                            writer, sheet_name=bladen[blad_idx], index=False,
                            header=rijen_in_blad == 0, startrow=rijen_in_blad + 1 if rijen_in_blad else 0
                        )
                        rijen_in_blad += len(deel)
                        start += len(deel)
                    self._voortgang(len(df), f"Tabel {index}/{totaal}: {tabel}")
                self._voortgang(0, f"Tabel {index}/{totaal} ingelezen")

            self.update_table_signal.emit(self._rij, "95%", STATUS_BEZIG, "Excel-bestand opslaan...")
            self.log_signal.emit("-> Opslaan van Excel-bestand...")
        return excel_pad

    def _schrijf_csv(self, conn: pyodbc.Connection, job: dict[str, Any], aangemaakt: list[str]) -> str:
        tabellen: list[str] = job["tabellen"]
        totaal = len(tabellen)
        gebruikte_namen: set[str] = set()
        self._max_pct = 100

        for index, tabel in enumerate(tabellen, 1):
            self._controleer_stop()
            csv_pad = os.path.join(job["doelmap"], maak_csv_naam(job["naam_zonder_ext"], tabel, gebruikte_namen))
            self.log_signal.emit(f"-> Schrijven CSV: '{tabel}' ({index}/{totaal})...")
            with open(csv_pad, "w", encoding="utf-8-sig", newline="") as f:
                aangemaakt.append(csv_pad)
                for blok_nr, df in enumerate(lees_in_blokken(conn, tabel)):
                    self._controleer_stop()
                    df.to_csv(f, index=False, sep=";", header=blok_nr == 0)
                    self._voortgang(len(df), f"CSV {index}/{totaal}: {tabel}")
            self._voortgang(0, f"CSV {index}/{totaal} geschreven")
        return job["doelmap"]

    def _ruim_op(self, paden: list[str]) -> None:
        for pad in paden:
            try:
                os.remove(pad)
            except OSError:
                pass


# =========================================================
# POP-UP DIALOOG: Tabellen selecteren & Preview
# =========================================================
class TableSelectionDialog(QDialog):
    def __init__(self, db_path: str, tabellen: list[str], geselecteerde_tabellen: list[str], parent: QWidget | None = None):
        super().__init__(parent)
        self.setWindowTitle("Tabellen Selecteren en Bekijken")
        self.resize(1000, 550)
        self.db_path = db_path
        self.alle_tabellen = tabellen
        self.geselecteerde = list(geselecteerde_tabellen)
        self._conn: pyodbc.Connection | None = None  # Eén verbinding voor alle previews

        hoofd_layout = QVBoxLayout(self)
        layout = QHBoxLayout()
        hoofd_layout.addLayout(layout)

        # Links: Lijst met Checkboxes
        left_layout = QVBoxLayout()
        left_layout.addWidget(self._vet_label("1. Selecteer tabellen (vink aan):"))

        self.txt_filter = QLineEdit()
        self.txt_filter.setPlaceholderText("🔍 Filter op tabelnaam...")
        self.txt_filter.setClearButtonEnabled(True)
        self.txt_filter.textChanged.connect(self.filter_lijst)
        left_layout.addWidget(self.txt_filter)

        self.list_widget = QListWidget()
        left_layout.addWidget(self.list_widget)

        geselecteerd = set(self.geselecteerde)
        for t in self.alle_tabellen:
            item = QListWidgetItem(t)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked if t in geselecteerd else Qt.CheckState.Unchecked)
            self.list_widget.addItem(item)

        # Zodra de gebruiker een ander item aanklikt, update de preview
        self.list_widget.currentItemChanged.connect(self.update_preview)
        self.list_widget.itemChanged.connect(self.update_teller)

        knoppen_layout = QHBoxLayout()
        btn_alles = QPushButton("☑ Alles aan")
        btn_alles.setToolTip("Vink alle (gefilterde) tabellen aan")
        btn_alles.clicked.connect(lambda: self.zet_alle_vinkjes(True))
        btn_geen = QPushButton("☐ Alles uit")
        btn_geen.setToolTip("Vink alle (gefilterde) tabellen uit")
        btn_geen.clicked.connect(lambda: self.zet_alle_vinkjes(False))
        knoppen_layout.addWidget(btn_alles)
        knoppen_layout.addWidget(btn_geen)
        left_layout.addLayout(knoppen_layout)

        self.lbl_teller = QLabel()
        left_layout.addWidget(self.lbl_teller)
        layout.addLayout(left_layout, 1)

        # Rechts: Preview
        right_layout = QVBoxLayout()
        right_layout.addWidget(self._vet_label(f"2. Datavoorbeeld (Eerste {PREVIEW_RIJEN} rijen):"))
        self.preview_table = QTableWidget()
        self.preview_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.preview_table.setAlternatingRowColors(True)
        right_layout.addWidget(self.preview_table)
        layout.addLayout(right_layout, 3)

        # Onder: Bevestigen / Annuleren
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btn_ok = buttons.button(QDialogButtonBox.StandardButton.Ok)
        if btn_ok:
            btn_ok.setText("✅ Bevestig Selectie")
            btn_ok.setObjectName("successButton")
        btn_annuleer = buttons.button(QDialogButtonBox.StandardButton.Cancel)
        if btn_annuleer:
            btn_annuleer.setText("Annuleren")
        buttons.accepted.connect(self.bevestig)
        buttons.rejected.connect(self.reject)
        hoofd_layout.addWidget(buttons)

        self.finished.connect(self._sluit_verbinding)
        language_support.refresh(self)
        self.update_teller()

        # Laad direct het voorbeeld van de eerste tabel in
        if self.alle_tabellen:
            self.list_widget.setCurrentRow(0)

    @staticmethod
    def _vet_label(tekst: str) -> QLabel:
        lbl = QLabel(tekst)
        font = lbl.font()
        font.setBold(True)
        lbl.setFont(font)
        return lbl

    def _items(self) -> Iterator[QListWidgetItem]:
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item is not None:
                yield item

    def filter_lijst(self, tekst: str) -> None:
        zoek = tekst.casefold()
        for item in self._items():
            item.setHidden(zoek not in item.text().casefold())

    def zet_alle_vinkjes(self, aan: bool) -> None:
        status = Qt.CheckState.Checked if aan else Qt.CheckState.Unchecked
        self.list_widget.blockSignals(True)
        for item in self._items():
            if not item.isHidden():
                item.setCheckState(status)
        self.list_widget.blockSignals(False)
        self.update_teller()

    def update_teller(self) -> None:
        aantal = sum(1 for item in self._items() if item.checkState() == Qt.CheckState.Checked)
        language_support.set_text(self.lbl_teller, f"{aantal} van {self.list_widget.count()} tabellen geselecteerd")

    def update_preview(self) -> None:
        huidig_item = self.list_widget.currentItem()
        if huidig_item is None:
            return

        self.preview_table.clear()
        try:
            if self._conn is None:
                self._conn = verbind_access(self.db_path)
            with closing(self._conn.cursor()) as cursor:
                cursor.execute(f"SELECT TOP {PREVIEW_RIJEN} * FROM {quote_tabel(huidig_item.text())}")
                kolommen = [c[0] for c in cursor.description]
                rijen = cursor.fetchall()

            self.preview_table.setColumnCount(len(kolommen))
            self.preview_table.setRowCount(len(rijen))
            self.preview_table.setHorizontalHeaderLabels(kolommen)
            for i, rij in enumerate(rijen):
                for j, val in enumerate(rij):
                    self.preview_table.setItem(i, j, QTableWidgetItem("" if val is None else str(val)))
            self.preview_table.resizeColumnsToContents()

        except Exception as e:
            self._sluit_verbinding()
            self.preview_table.setColumnCount(1)
            self.preview_table.setRowCount(1)
            self.preview_table.setHorizontalHeaderLabels([tr("Foutmelding")])
            item = QTableWidgetItem(str(e).split("\n")[0])
            item.setToolTip(str(e))
            self.preview_table.setItem(0, 0, item)
            self.preview_table.resizeColumnsToContents()

    def _sluit_verbinding(self) -> None:
        if self._conn is not None:
            try:
                self._conn.close()
            except pyodbc.Error:
                pass
            self._conn = None

    def bevestig(self) -> None:
        # Verzamel alleen de tabellen waarvan het vinkje (checkbox) aan staat
        self.geselecteerde = [item.text() for item in self._items() if item.checkState() == Qt.CheckState.Checked]
        self.accept()


# =========================================================
# HOOFDPROGRAMMA (GUI)
# =========================================================
class AccessExporterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Access naar Excel/CSV Exporter v{__version__}")
        self.resize(1100, 800)

        # State variabelen
        self.settings = QSettings("AccessExporter", "AccessExporterPyQt6")
        self.access_pad = ""
        self.doelmap_pad = ""
        self.huidige_db_tabellen: list[str] = []
        self.geselecteerde_tabellen: list[str] = []
        self.jobs: list[dict[str, Any]] = []
        self.worker: ConversionWorker | None = None
        self.log_entries: list[tuple[str, str]] = []

        self.init_ui()
        language_support.refresh(self)
        self.apply_stylesheet()
        self.laad_instellingen()
        self.controleer_omgeving()

    def init_ui(self) -> None:
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        language_layout = QHBoxLayout()
        language_layout.addStretch()
        language_label = QLabel("Taal")
        self.language_combo = QComboBox()
        for label, code in (("Nederlands", "nl"), ("English", "en"), ("Français", "fr")):
            self.language_combo.addItem(label, code)
        language_label.setBuddy(self.language_combo)
        self.language_combo.currentIndexChanged.connect(self.wijzig_taal)
        language_layout.addWidget(language_label)
        language_layout.addWidget(self.language_combo)
        main_layout.addLayout(language_layout)

        # --- STAP 1: Instellingen ---
        gb_settings = QGroupBox("Stap 1: Bestanden, tabellen en formaat instellen")
        grid_settings = QGridLayout(gb_settings)

        # Rij 0
        self.btn_kies_access = QPushButton("📂 1. Kies Access-bestand")
        self.btn_kies_access.clicked.connect(self.kies_access)
        grid_settings.addWidget(self.btn_kies_access, 0, 0)

        self.format_group = QGroupBox("⚙ 2. Kies Exportformaat")
        format_layout = QHBoxLayout(self.format_group)
        self.rb_xlsx = QRadioButton("📊 1 Excel-bestand (.xlsx)")
        self.rb_csv = QRadioButton("📄 Aparte CSV's per tabel (.csv)")
        self.rb_xlsx.setChecked(True)
        format_layout.addWidget(self.rb_xlsx)
        format_layout.addWidget(self.rb_csv)
        grid_settings.addWidget(self.format_group, 0, 1)

        # Rij 1
        self.lbl_access_pad = QLabel("Geen Access-bestand geselecteerd")
        self.lbl_access_pad.setStyleSheet(f"color: {KLEUR_INFO};")
        grid_settings.addWidget(self.lbl_access_pad, 1, 0, 1, 2)

        # Rij 2
        self.btn_kies_tabellen = QPushButton("📋 3. Kies te exporteren tabellen")
        self.btn_kies_tabellen.clicked.connect(self.kies_tabellen)
        grid_settings.addWidget(self.btn_kies_tabellen, 2, 0)

        self.lbl_tabellen_status = QLabel("Geen tabellen geladen")
        grid_settings.addWidget(self.lbl_tabellen_status, 2, 1)

        # Rij 3
        self.btn_kies_doelmap = QPushButton("📁 4. Kies Doelmap")
        self.btn_kies_doelmap.clicked.connect(self.kies_doelmap)
        grid_settings.addWidget(self.btn_kies_doelmap, 3, 0)

        self.btn_voeg_toe = QPushButton("➕ 5. Voeg toe aan wachtrij")
        self.btn_voeg_toe.setObjectName("accentButton")
        self.btn_voeg_toe.clicked.connect(self.voeg_toe)
        grid_settings.addWidget(self.btn_voeg_toe, 3, 1)

        # Rij 4
        self.lbl_doelmap_pad = QLabel("Geen map geselecteerd")
        self.lbl_doelmap_pad.setStyleSheet(f"color: {KLEUR_INFO};")
        grid_settings.addWidget(self.lbl_doelmap_pad, 4, 0, 1, 2)

        main_layout.addWidget(gb_settings)

        # --- STAP 2: Wachtrij ---
        gb_queue = QGroupBox("Stap 2: Wachtrij")
        queue_layout = QVBoxLayout(gb_queue)

        self.table_queue = QTableWidget(0, 5)
        self.table_queue.setHorizontalHeaderLabels(["Access-bestand", "Doelbestand(en)", "Voortgang", "Status", "Melding (Statusreden)"])
        self.table_queue.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_queue.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_queue.setAlternatingRowColors(True)

        header = self.table_queue.horizontalHeader()
        if header:
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table_queue.setColumnWidth(0, 220)
        self.table_queue.setColumnWidth(1, 220)

        # Delete-toets verwijdert de geselecteerde rijen
        QShortcut(QKeySequence.StandardKey.Delete, self.table_queue).activated.connect(self.verwijder_selectie)

        queue_layout.addWidget(self.table_queue)
        main_layout.addWidget(gb_queue)

        # --- STAP 3: Logboek ---
        gb_log = QGroupBox("Logboek - Live status")
        log_layout = QVBoxLayout(gb_log)
        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setFixedHeight(120)
        log_layout.addWidget(self.txt_log)
        main_layout.addWidget(gb_log)

        # --- STAP 4: Acties ---
        action_layout = QHBoxLayout()
        self.btn_verwijder = QPushButton("🗑️ Verwijder Selectie")
        self.btn_verwijder.setToolTip("Verwijder de geselecteerde rijen uit de wachtrij (Delete)")
        self.btn_verwijder.clicked.connect(self.verwijder_selectie)

        self.chk_open_map = QCheckBox("Map openen na afloop")
        self.chk_open_map.setChecked(True)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 1)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%v/%m")
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setFixedWidth(150)

        self.btn_stop = QPushButton("⏹ Stop")
        self.btn_stop.setToolTip("Stop de lopende conversie (gedeeltelijke bestanden worden verwijderd)")
        self.btn_stop.setEnabled(False)
        self.btn_stop.clicked.connect(self.stop_conversie)

        self.btn_conv_sel = QPushButton("▶️ Converteer Selectie")
        self.btn_conv_sel.setObjectName("warningButton")
        self.btn_conv_sel.clicked.connect(lambda: self.start_conversie(False))

        self.btn_conv_all = QPushButton("🚀 Converteer Hele Wachtrij")
        self.btn_conv_all.setObjectName("successButton")
        self.btn_conv_all.setToolTip("Verwerkt alle taken met status 'In de wachtrij', 'Fout' of 'Gestopt'")
        self.btn_conv_all.clicked.connect(lambda: self.start_conversie(True))

        action_layout.addWidget(self.btn_verwijder)
        action_layout.addWidget(self.chk_open_map)
        action_layout.addStretch()
        action_layout.addWidget(self.progress_bar)
        action_layout.addWidget(self.btn_stop)
        action_layout.addWidget(self.btn_conv_sel)
        action_layout.addWidget(self.btn_conv_all)

        main_layout.addLayout(action_layout)
        self.log_msg("Applicatie opgestart en klaar voor gebruik.")

    def wijzig_taal(self) -> None:
        language_support.set_language(self.language_combo.currentData())
        self.settings.setValue("taal", language_support.language)
        language_support.refresh(self)
        self.table_queue.setHorizontalHeaderLabels([
            tr(source) for source in (
                "Access-bestand", "Doelbestand(en)", "Voortgang", "Status", "Melding (Statusreden)"
            )
        ])
        for row in range(self.table_queue.rowCount()):
            for column in (1, 3, 4):
                item = self.table_queue.item(row, column)
                if item is not None:
                    source = item.data(Qt.ItemDataRole.UserRole)
                    if source is not None:
                        item.setText(tr(source))
                        if column == 4:
                            item.setToolTip(tr(source))
        self.txt_log.setPlainText("\n".join(
            f"[{time}] {tr(source)}" for time, source in self.log_entries
        ))
        self.txt_log.verticalScrollBar().setValue(self.txt_log.verticalScrollBar().maximum())

    # --- THEMA (QSS) Dark Mode ---
    def apply_stylesheet(self) -> None:
        style = """
            QMainWindow, QDialog, QWidget { background-color: #121212; color: #E0E0E0; }
            QGroupBox { font-weight: bold; border: 1px solid #333333; border-radius: 5px; margin-top: 10px; padding-top: 15px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 3px; color: #FFFFFF; }
            QPushButton { background-color: #2D2D2D; border: 1px solid #444444; border-radius: 4px; padding: 6px; font-size: 13px; color: #E0E0E0; }
            QPushButton:hover { background-color: #3D3D3D; border-color: #0078D7; }
            QPushButton#accentButton { background-color: #0078D7; color: white; border: none; font-weight: bold; }
            QPushButton#accentButton:hover { background-color: #005a9e; }
            QPushButton#successButton { background-color: #28a745; color: white; border: none; font-weight: bold; }
            QPushButton#successButton:hover { background-color: #218838; }
            QPushButton#warningButton { background-color: #fd7e14; color: white; border: none; font-weight: bold; }
            QPushButton#warningButton:hover { background-color: #e36209; }
            QPushButton:disabled, QPushButton#accentButton:disabled, QPushButton#successButton:disabled, QPushButton#warningButton:disabled { background-color: #1E1E1E; color: #666666; border: 1px solid #333333; }
            QTableWidget, QListWidget { background-color: #1E1E1E; alternate-background-color: #252525; border: 1px solid #333333; color: #E0E0E0; gridline-color: #333333; }
            QHeaderView::section { background-color: #2D2D2D; padding: 4px; border: 1px solid #333333; font-weight: bold; color: #FFFFFF; }
            QPlainTextEdit, QLineEdit, QComboBox { background-color: #1E1E1E; border: 1px solid #333333; color: #E0E0E0; }
            QPlainTextEdit { font-family: Consolas, monospace; }
            QLineEdit { padding: 4px; border-radius: 4px; }
            QLineEdit:focus { border-color: #0078D7; }
            QProgressBar { border: 1px solid #444444; border-radius: 4px; text-align: center; background-color: #1E1E1E; color: #FFFFFF; }
            QProgressBar::chunk { background-color: #0078D7; width: 20px; }
            QLabel, QRadioButton, QCheckBox { color: #E0E0E0; }
            QRadioButton:disabled { color: #666666; }
        """
        self.setStyleSheet(style)

    # --- INSTELLINGEN & OMGEVING ---
    def laad_instellingen(self) -> None:
        language = str(self.settings.value("taal", "nl") or "nl")
        index = self.language_combo.findData(language)
        self.language_combo.blockSignals(True)
        self.language_combo.setCurrentIndex(max(0, index))
        self.language_combo.blockSignals(False)
        self.wijzig_taal()
        doelmap = str(self.settings.value("doelmap", "") or "")
        if doelmap and os.path.isdir(doelmap):
            self.zet_doelmap(doelmap)
        if self.settings.value("formaat", "xlsx") == "csv":
            self.rb_csv.setChecked(True)
        self.chk_open_map.setChecked(self.settings.value("open_map", True, type=bool))

    def bewaar_instellingen(self) -> None:
        self.settings.setValue("taal", language_support.language)
        self.settings.setValue("doelmap", self.doelmap_pad)
        self.settings.setValue("formaat", "xlsx" if self.rb_xlsx.isChecked() else "csv")
        self.settings.setValue("open_map", self.chk_open_map.isChecked())

    def controleer_omgeving(self) -> None:
        if EXCEL_ENGINE is None:
            self.rb_xlsx.setEnabled(False)
            self.rb_xlsx.setProperty("source_toolTip", "Installeer xlsxwriter: pip install xlsxwriter")
            self.rb_xlsx.setToolTip(tr("Installeer xlsxwriter: pip install xlsxwriter"))
            self.rb_csv.setChecked(True)
            self.log_msg("⚠ Geen Excel-bibliotheek gevonden; alleen CSV-export beschikbaar (pip install xlsxwriter).")
        elif EXCEL_ENGINE != "xlsxwriter":
            self.log_msg(f"ℹ Excel-export via {EXCEL_ENGINE}. Installeer xlsxwriter voor snellere export.")

        if not access_driver_beschikbaar():
            bits = 64 if sys.maxsize > 2**32 else 32
            melding = (f"De ODBC-driver '{ACCESS_DRIVER}' is niet gevonden.\n\n"
                       f"Installeer de Microsoft Access Database Engine ({bits}-bit, gelijk aan deze Python-installatie).")
            self.log_msg(f"❌ {melding.splitlines()[0]}")
            QMessageBox.critical(self, tr("Access-driver ontbreekt"), tr(melding))

    # --- ACTIES ---
    def log_msg(self, msg: str) -> None:
        tijd = datetime.now().strftime("%H:%M:%S")
        self.log_entries.append((tijd, msg))
        self.txt_log.appendPlainText(f"[{tijd}] {tr(msg)}")

    def zet_tabellen_status(self, tekst: str, kleur: str = KLEUR_OK) -> None:
        language_support.set_text(self.lbl_tabellen_status, tekst)
        self.lbl_tabellen_status.setStyleSheet(f"color: {kleur};")

    def zet_doelmap(self, pad: str) -> None:
        self.doelmap_pad = os.path.normpath(pad)
        language_support.set_text(self.lbl_doelmap_pad, self.doelmap_pad)

    def kies_access(self) -> None:
        startmap = str(self.settings.value("laatste_access_map", "") or "")
        pad, _ = QFileDialog.getOpenFileName(self, tr("Selecteer Access Database"), startmap, tr("Access Databases (*.accdb *.mdb);;Alle bestanden (*.*)"))
        if not pad:
            return

        self.access_pad = os.path.normpath(pad)
        self.settings.setValue("laatste_access_map", os.path.dirname(self.access_pad))
        language_support.set_text(self.lbl_access_pad, self.access_pad)
        self.log_msg(f"Bestand geselecteerd: {os.path.basename(self.access_pad)}")

        if not self.doelmap_pad:
            self.zet_doelmap(os.path.dirname(self.access_pad))
            self.log_msg(f"Doelmap ingesteld op map van het Access-bestand: {self.doelmap_pad}")

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            self.huidige_db_tabellen = haal_tabellen_op(self.access_pad)
            self.geselecteerde_tabellen = self.huidige_db_tabellen.copy()
            if self.huidige_db_tabellen:
                self.zet_tabellen_status(f"Alle ({len(self.geselecteerde_tabellen)}) tabellen geselecteerd")
            else:
                self.zet_tabellen_status("Geen tabellen gevonden in dit bestand", KLEUR_FOUT)
        except Exception as e:
            self.huidige_db_tabellen = []
            self.geselecteerde_tabellen = []
            self.log_msg(f"❌ Fout bij ophalen tabellen: {e}")
            self.zet_tabellen_status("Fout bij laden van tabellen", KLEUR_FOUT)
        finally:
            QApplication.restoreOverrideCursor()

    def kies_tabellen(self) -> None:
        if not self.access_pad or not self.huidige_db_tabellen:
            QMessageBox.warning(self, tr("Let op"), tr("Kies eerst een geldig Access-bestand."))
            return

        dialog = TableSelectionDialog(self.access_pad, self.huidige_db_tabellen, self.geselecteerde_tabellen, self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        self.geselecteerde_tabellen = dialog.geselecteerde
        aantal = len(self.geselecteerde_tabellen)
        if aantal == len(self.huidige_db_tabellen):
            self.zet_tabellen_status(f"Alle ({aantal}) tabellen geselecteerd")
        elif aantal == 0:
            self.zet_tabellen_status("Let op: 0 tabellen geselecteerd!", KLEUR_FOUT)
        else:
            kort = ", ".join(self.geselecteerde_tabellen[:2]) + ("..." if aantal > 2 else "")
            self.zet_tabellen_status(f"{aantal} geselecteerd ({kort})")
        self.lbl_tabellen_status.setToolTip("\n".join(self.geselecteerde_tabellen))

    def kies_doelmap(self) -> None:
        pad = QFileDialog.getExistingDirectory(self, tr("Kies de doelmap"), self.doelmap_pad)
        if pad:
            self.zet_doelmap(pad)
            self.log_msg(f"Doelmap ingesteld: {self.doelmap_pad}")

    def voeg_toe(self) -> None:
        if not self.access_pad or not self.doelmap_pad:
            QMessageBox.warning(self, tr("Incompleet"), tr("Selecteer eerst een Access-bestand en een doelmap."))
            return
        if not self.geselecteerde_tabellen:
            QMessageBox.warning(self, tr("Incompleet"), tr("Je hebt geen tabellen geselecteerd."))
            return

        export_type = "xlsx" if self.rb_xlsx.isChecked() else "csv"
        dubbel = any(j["access_pad"] == self.access_pad and j["doelmap"] == self.doelmap_pad and j["export_type"] == export_type
                     for j in self.jobs)
        if dubbel and QMessageBox.question(
            self, tr("Al in wachtrij"),
            tr("Dit bestand staat al met hetzelfde formaat en dezelfde doelmap in de wachtrij.\nToch nogmaals toevoegen?")
        ) != QMessageBox.StandardButton.Yes:
            return

        access_naam = os.path.basename(self.access_pad)
        naam_zonder_ext = os.path.splitext(access_naam)[0]
        doel_weergave = f"{naam_zonder_ext}.xlsx" if export_type == "xlsx" else f"Meerdere CSV's ({naam_zonder_ext}_...)"

        row = self.table_queue.rowCount()
        self.table_queue.insertRow(row)
        for kolom, tekst in enumerate([access_naam, doel_weergave, "0%", STATUS_WACHTRIJ,
                                       f"{len(self.geselecteerde_tabellen)} tabellen geselecteerd"]):
            item = QTableWidgetItem(tr(tekst) if kolom in (1, 3, 4) else tekst)
            if kolom in (1, 3, 4):
                item.setData(Qt.ItemDataRole.UserRole, tekst)
            self.table_queue.setItem(row, kolom, item)
        bron_item = self.table_queue.item(row, 0)
        if bron_item:
            bron_item.setToolTip(self.access_pad)
        doel_item = self.table_queue.item(row, 1)
        if doel_item:
            doel_item.setToolTip(self.doelmap_pad)

        self.jobs.append({
            "access_pad": self.access_pad,
            "doelmap": self.doelmap_pad,
            "access_naam": access_naam,
            "naam_zonder_ext": naam_zonder_ext,
            "export_type": export_type,
            "tabellen": self.geselecteerde_tabellen.copy(),
        })
        self.log_msg(f"Toegevoegd aan wachtrij: {access_naam} als {export_type.upper()}")

        self.access_pad = ""
        language_support.set_text(self.lbl_access_pad, "Geen Access-bestand geselecteerd")
        self.huidige_db_tabellen = []
        self.geselecteerde_tabellen = []
        self.zet_tabellen_status("Geen tabellen geladen")
        self.lbl_tabellen_status.setToolTip("")

    def geselecteerde_rijen(self) -> list[int]:
        model = self.table_queue.selectionModel()
        return sorted(index.row() for index in model.selectedRows()) if model else []

    def verwijder_selectie(self) -> None:
        if self.conversie_bezig():
            return
        for row in reversed(self.geselecteerde_rijen()):
            self.log_msg(f"Verwijderd uit wachtrij: {self.jobs[row]['access_naam']}")
            self.table_queue.removeRow(row)
            del self.jobs[row]

    def conversie_bezig(self) -> bool:
        return self.worker is not None and self.worker.isRunning()

    def status_van_rij(self, row: int) -> str:
        item = self.table_queue.item(row, 3)
        return item.data(Qt.ItemDataRole.UserRole) if item else ""

    def start_conversie(self, doe_alles: bool = False) -> None:
        if self.conversie_bezig():
            return

        if doe_alles:
            rows = [r for r in range(len(self.jobs)) if self.status_van_rij(r) in (STATUS_WACHTRIJ, STATUS_FOUT, STATUS_GESTOPT)]
            if not rows:
                QMessageBox.information(self, tr("Info"), tr("Er zijn geen bestanden in de wachtrij om te verwerken."))
                return
        else:
            rows = self.geselecteerde_rijen()
            if not rows:
                QMessageBox.warning(self, tr("Geen selectie"), tr("Selecteer eerst een rij in de wachtrij."))
                return

        self.zet_knoppen_bezig(True)
        self.progress_bar.setRange(0, len(rows))
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)

        # Kopie van de jobs: de wachtrij kan tijdens de conversie worden aangevuld
        taken = [(r, dict(self.jobs[r])) for r in rows]
        self.worker = ConversionWorker(taken, self.chk_open_map.isChecked())
        self.worker.update_table_signal.connect(self.update_table_row)
        self.worker.log_signal.connect(self.log_msg)
        self.worker.job_done_signal.connect(lambda: self.progress_bar.setValue(self.progress_bar.value() + 1))
        self.worker.finished.connect(self.conversie_klaar)
        self.worker.start()

    def stop_conversie(self) -> None:
        if self.worker is not None and self.conversie_bezig():
            self.worker.requestInterruption()
            self.btn_stop.setEnabled(False)
            self.log_msg("⏹ Stoppen aangevraagd...")

    def zet_knoppen_bezig(self, bezig: bool) -> None:
        self.btn_conv_sel.setEnabled(not bezig)
        self.btn_conv_all.setEnabled(not bezig)
        self.btn_verwijder.setEnabled(not bezig)
        self.btn_stop.setEnabled(bezig)

    def update_table_row(self, row: int, voortgang: str, status: str, melding: str) -> None:
        for kolom, tekst in ((2, voortgang), (3, status), (4, melding)):
            item = self.table_queue.item(row, kolom)
            if item is None:
                item = QTableWidgetItem()
                self.table_queue.setItem(row, kolom, item)
            item.setData(Qt.ItemDataRole.UserRole, tekst)
            item.setText(tr(tekst))
            if kolom == 4:
                item.setToolTip(tr(tekst))
        status_item = self.table_queue.item(row, 3)
        if status_item:
            status_item.setForeground(QColor(STATUS_KLEUREN.get(status, "#E0E0E0")))

    def conversie_klaar(self) -> None:
        self.log_msg("Alle geselecteerde taken in wachtrij voltooid.")
        self.zet_knoppen_bezig(False)

    def closeEvent(self, a0: QCloseEvent | None) -> None:
        if self.worker is not None and self.conversie_bezig():
            antwoord = QMessageBox.question(self, tr("Conversie bezig"), tr("Er loopt nog een conversie. Stoppen en afsluiten?"))
            if antwoord != QMessageBox.StandardButton.Yes:
                if a0:
                    a0.ignore()
                return
            self.worker.requestInterruption()
            self.worker.wait()
        self.bewaar_instellingen()
        super().closeEvent(a0)


def toon_onverwachte_fout(exc_type, exc_value, exc_tb) -> None:
    # PyQt6 sluit de applicatie af bij een onafgehandelde fout in een slot; toon in plaats daarvan een melding
    import traceback
    details = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    print(details, file=sys.stderr)
    box = QMessageBox(QMessageBox.Icon.Critical, tr("Onverwachte fout"), str(exc_value) or exc_type.__name__)
    box.setDetailedText(details)
    box.exec()


if __name__ == "__main__":
    sys.excepthook = toon_onverwachte_fout
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    window = AccessExporterGUI()
    window.show()
    sys.exit(app.exec())
