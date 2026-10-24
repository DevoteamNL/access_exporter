"""Hoofdvenster: taal- en themakeuze, tabbladen Exporteren / Bewerken / Terugzetten en een gedeeld logboek."""

import os
import sys
from datetime import datetime
from typing import Any

from PyQt6.QtCore import QSettings, Qt
from PyQt6.QtGui import QCloseEvent, QColor, QDragEnterEvent, QDropEvent, QGuiApplication, QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QDialog, QFileDialog, QGridLayout, QGroupBox,
    QHBoxLayout, QHeaderView, QLabel, QMainWindow, QMessageBox, QPlainTextEdit, QProgressBar, QPushButton,
    QRadioButton, QTableWidget, QTableWidgetItem, QTabWidget, QVBoxLayout, QWidget,
)

from . import __version__, themes
from .core import (
    ACCESS_DRIVER, EXCEL_ENGINE, STATUS_BEZIG, STATUS_FOUT, STATUS_GESTOPT, STATUS_KLAAR, STATUS_WACHTRIJ,
    access_driver_beschikbaar, haal_tabellen_op, is_access_bestand,
)
from .dialogs import TableSelectionDialog, met_wachtwoord
from .editor_tab import EditorTab
from .merge_tab import MergeTab
from .translations import language_support, tr
from .worker import ConversionWorker

STATUS_KLEURSLEUTELS = {STATUS_KLAAR: "ok", STATUS_FOUT: "fout", STATUS_BEZIG: "bezig", STATUS_GESTOPT: "bezig"}
TABBLADEN = ("📤 Exporteren", "✏️ Bewerken", "🔁 Terugzetten naar Access")
THEMA_LABELS = {themes.THEMA_SYSTEEM: "Systeem", themes.THEMA_LICHT: "Licht", themes.THEMA_DONKER: "Donker"}


class AccessExporterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Access naar Excel/CSV Exporter v{__version__}")
        self.resize(1150, 850)
        self.setAcceptDrops(True)

        self.settings = QSettings("AccessExporter", "AccessExporterPyQt6")
        self.access_pad = ""
        self.access_wachtwoord = ""
        self.doelmap_pad = ""
        self.huidige_db_tabellen: list[str] = []
        self.geselecteerde_tabellen: list[str] = []
        self.jobs: list[dict[str, Any]] = []
        self.worker: ConversionWorker | None = None
        self.log_entries: list[tuple[str, str]] = []
        self.tabellen_status: tuple[str, str] = ("Geen tabellen geladen", "ok")

        self.init_ui()
        language_support.refresh(self)
        self.laad_instellingen()
        self.controleer_omgeving()

    # =========================================================
    # OPBOUW
    # =========================================================
    def init_ui(self) -> None:
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        top_layout = QHBoxLayout()
        top_layout.addStretch()
        theme_label = QLabel("Thema")
        self.theme_combo = QComboBox()
        for code in themes.THEMAS:
            self.theme_combo.addItem(tr(THEMA_LABELS[code]), code)
        theme_label.setBuddy(self.theme_combo)
        self.theme_combo.currentIndexChanged.connect(self.wijzig_thema)
        top_layout.addWidget(theme_label)
        top_layout.addWidget(self.theme_combo)
        top_layout.addSpacing(12)
        language_label = QLabel("Taal")
        self.language_combo = QComboBox()
        for label, code in (("Nederlands", "nl"), ("English", "en"), ("Français", "fr")):
            self.language_combo.addItem(label, code)
        language_label.setBuddy(self.language_combo)
        self.language_combo.currentIndexChanged.connect(self.wijzig_taal)
        top_layout.addWidget(language_label)
        top_layout.addWidget(self.language_combo)
        main_layout.addLayout(top_layout)

        self.tabs = QTabWidget()
        export_tab = QWidget()
        self.init_export_tab(export_tab)
        self.editor_tab = EditorTab(self.log_msg, self.settings)
        self.merge_tab = MergeTab(self.log_msg, self.settings)
        for widget, titel in zip((export_tab, self.editor_tab, self.merge_tab), TABBLADEN):
            self.tabs.addTab(widget, tr(titel))
        main_layout.addWidget(self.tabs, 1)

        gb_log = QGroupBox("Logboek - Live status")
        log_layout = QVBoxLayout(gb_log)
        self.txt_log = QPlainTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setFixedHeight(110)
        log_layout.addWidget(self.txt_log)
        main_layout.addWidget(gb_log)
        self.log_msg("Applicatie opgestart en klaar voor gebruik.")

    def init_export_tab(self, tab: QWidget) -> None:
        tab_layout = QVBoxLayout(tab)

        # --- STAP 1: Instellingen ---
        gb_settings = QGroupBox("Stap 1: Bestanden, tabellen en formaat instellen")
        grid_settings = QGridLayout(gb_settings)

        self.btn_kies_access = QPushButton("📂 1. Kies Access-bestand")
        self.btn_kies_access.setToolTip("Of sleep een .accdb/.mdb-bestand naar dit venster")
        self.btn_kies_access.clicked.connect(self.kies_access)
        grid_settings.addWidget(self.btn_kies_access, 0, 0)

        self.format_group = QGroupBox("⚙ 2. Kies Exportformaat")
        format_layout = QHBoxLayout(self.format_group)
        self.rb_xlsx = QRadioButton("📊 1 Excel-bestand (.xlsx)")
        self.rb_csv = QRadioButton("📄 Aparte CSV's per tabel (.csv)")
        self.rb_xlsx.setChecked(True)
        self.cmb_decimaal = QComboBox()
        self.cmb_decimaal.setToolTip("Decimaalteken in CSV-bestanden")
        self.cmb_decimaal.addItem(tr("Decimale komma (1,5)"), ",")
        self.cmb_decimaal.addItem(tr("Decimale punt (1.5)"), ".")
        self.rb_csv.toggled.connect(self.cmb_decimaal.setEnabled)
        self.cmb_decimaal.setEnabled(False)
        format_layout.addWidget(self.rb_xlsx)
        format_layout.addWidget(self.rb_csv)
        format_layout.addWidget(self.cmb_decimaal)
        grid_settings.addWidget(self.format_group, 0, 1)

        self.lbl_access_pad = QLabel("Geen Access-bestand geselecteerd")
        self.lbl_access_pad.setObjectName("infoLabel")
        grid_settings.addWidget(self.lbl_access_pad, 1, 0, 1, 2)

        self.btn_kies_tabellen = QPushButton("📋 3. Kies te exporteren tabellen")
        self.btn_kies_tabellen.clicked.connect(self.kies_tabellen)
        grid_settings.addWidget(self.btn_kies_tabellen, 2, 0)

        tabellen_layout = QHBoxLayout()
        self.lbl_tabellen_status = QLabel("Geen tabellen geladen")
        self.chk_queries = QCheckBox("Ook opgeslagen query's")
        self.chk_queries.setToolTip("Toon ook opgeslagen selectiequery's in de tabellijst, zodat je hun resultaat kunt exporteren")
        self.chk_queries.toggled.connect(self.laad_tabellen)
        tabellen_layout.addWidget(self.lbl_tabellen_status, 1)
        tabellen_layout.addWidget(self.chk_queries)
        grid_settings.addLayout(tabellen_layout, 2, 1)

        self.btn_kies_doelmap = QPushButton("📁 4. Kies Doelmap")
        self.btn_kies_doelmap.clicked.connect(self.kies_doelmap)
        grid_settings.addWidget(self.btn_kies_doelmap, 3, 0)

        self.btn_voeg_toe = QPushButton("➕ 5. Voeg toe aan wachtrij")
        self.btn_voeg_toe.setObjectName("accentButton")
        self.btn_voeg_toe.clicked.connect(self.voeg_toe)
        grid_settings.addWidget(self.btn_voeg_toe, 3, 1)

        self.lbl_doelmap_pad = QLabel("Geen map geselecteerd")
        self.lbl_doelmap_pad.setObjectName("infoLabel")
        grid_settings.addWidget(self.lbl_doelmap_pad, 4, 0, 1, 2)
        tab_layout.addWidget(gb_settings)

        # --- STAP 2: Wachtrij ---
        gb_queue = QGroupBox("Stap 2: Wachtrij")
        queue_layout = QVBoxLayout(gb_queue)
        self.table_queue = QTableWidget(0, 5)
        self.table_queue.setHorizontalHeaderLabels(["Access-bestand", "Doelbestand(en)", "Voortgang", "Status", "Melding (Statusreden)"])
        self.table_queue.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table_queue.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_queue.setAlternatingRowColors(True)
        self.table_queue.setToolTip("Dubbelklik op een afgeronde Excel-taak om het bestand in de editor te openen")
        self.table_queue.cellDoubleClicked.connect(self.open_job_in_editor)
        header = self.table_queue.horizontalHeader()
        if header:
            header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
            header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
            header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
            header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        self.table_queue.setColumnWidth(0, 220)
        self.table_queue.setColumnWidth(1, 220)
        QShortcut(QKeySequence.StandardKey.Delete, self.table_queue).activated.connect(self.verwijder_selectie)
        queue_layout.addWidget(self.table_queue)
        tab_layout.addWidget(gb_queue, 1)

        # --- Acties ---
        action_layout = QHBoxLayout()
        self.btn_verwijder = QPushButton("🗑️ Verwijder Selectie")
        self.btn_verwijder.setToolTip("Verwijder de geselecteerde rijen uit de wachtrij (Delete)")
        self.btn_verwijder.clicked.connect(self.verwijder_selectie)

        self.btn_wis_klaar = QPushButton("🧹 Wis afgeronde taken")
        self.btn_wis_klaar.setToolTip("Verwijder alle taken met status 'Klaar' uit de wachtrij")
        self.btn_wis_klaar.clicked.connect(self.wis_afgeronde_taken)

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

        for widget in (self.btn_verwijder, self.btn_wis_klaar, self.chk_open_map):
            action_layout.addWidget(widget)
        action_layout.addStretch()
        for widget in (self.progress_bar, self.btn_stop, self.btn_conv_sel, self.btn_conv_all):
            action_layout.addWidget(widget)
        tab_layout.addLayout(action_layout)

    # =========================================================
    # TAAL & THEMA
    # =========================================================
    def wijzig_taal(self) -> None:
        language_support.set_language(self.language_combo.currentData())
        self.settings.setValue("taal", language_support.language)
        language_support.refresh(self)
        self.table_queue.setHorizontalHeaderLabels([
            tr(source) for source in ("Access-bestand", "Doelbestand(en)", "Voortgang", "Status", "Melding (Statusreden)")
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
        for index, titel in enumerate(TABBLADEN):
            self.tabs.setTabText(index, tr(titel))
        for index in range(self.theme_combo.count()):
            self.theme_combo.setItemText(index, tr(THEMA_LABELS[self.theme_combo.itemData(index)]))
        for index, bron in enumerate(("Decimale komma (1,5)", "Decimale punt (1.5)")):
            self.cmb_decimaal.setItemText(index, tr(bron))
        self.editor_tab.vertaal()
        self.merge_tab.vertaal()
        self.txt_log.setPlainText("\n".join(f"[{time}] {tr(source)}" for time, source in self.log_entries))
        scrollbar = self.txt_log.verticalScrollBar()
        if scrollbar:
            scrollbar.setValue(scrollbar.maximum())

    def wijzig_thema(self) -> None:
        thema = self.theme_combo.currentData() or themes.THEMA_SYSTEEM
        self.settings.setValue("thema", thema)
        themes.pas_toe(thema)
        self.kleur_opnieuw()

    def systeemthema_gewijzigd(self) -> None:
        if self.theme_combo.currentData() == themes.THEMA_SYSTEEM:
            self.wijzig_thema()

    def kleur_opnieuw(self) -> None:
        """Kleuren die niet via de stylesheet lopen (status-kolom, tabelstatus) opnieuw toepassen."""
        tekst, sleutel = self.tabellen_status
        self.zet_tabellen_status(tekst, sleutel)
        for row in range(self.table_queue.rowCount()):
            self.kleur_status(row)
        self.merge_tab.kleur_opnieuw()

    def kleur_status(self, row: int) -> None:
        item = self.table_queue.item(row, 3)
        if item:
            sleutel = STATUS_KLEURSLEUTELS.get(item.data(Qt.ItemDataRole.UserRole), "tekst")
            item.setForeground(QColor(themes.kleur(sleutel)))

    # =========================================================
    # INSTELLINGEN & OMGEVING
    # =========================================================
    def laad_instellingen(self) -> None:
        thema = str(self.settings.value("thema", themes.THEMA_SYSTEEM) or themes.THEMA_SYSTEEM)
        self.theme_combo.blockSignals(True)
        self.theme_combo.setCurrentIndex(max(0, self.theme_combo.findData(thema)))
        self.theme_combo.blockSignals(False)
        self.wijzig_thema()
        hints = QGuiApplication.styleHints()
        if hints is not None:
            hints.colorSchemeChanged.connect(self.systeemthema_gewijzigd)

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
        self.cmb_decimaal.setCurrentIndex(max(0, self.cmb_decimaal.findData(self.settings.value("csv_decimaal", ","))))
        self.chk_open_map.setChecked(self.settings.value("open_map", True, type=bool))
        self.chk_queries.blockSignals(True)
        self.chk_queries.setChecked(self.settings.value("met_queries", False, type=bool))
        self.chk_queries.blockSignals(False)

    def bewaar_instellingen(self) -> None:
        self.settings.setValue("taal", language_support.language)
        self.settings.setValue("thema", self.theme_combo.currentData())
        self.settings.setValue("doelmap", self.doelmap_pad)
        self.settings.setValue("formaat", "xlsx" if self.rb_xlsx.isChecked() else "csv")
        self.settings.setValue("csv_decimaal", self.cmb_decimaal.currentData())
        self.settings.setValue("open_map", self.chk_open_map.isChecked())
        self.settings.setValue("met_queries", self.chk_queries.isChecked())

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

    # =========================================================
    # SLEPEN & NEERZETTEN
    # =========================================================
    @staticmethod
    def _paden(event: QDragEnterEvent | QDropEvent | None) -> list[str]:
        mime = event.mimeData() if event else None
        if mime is None or not mime.hasUrls():
            return []
        return [os.path.normpath(u.toLocalFile()) for u in mime.urls() if u.isLocalFile()]

    def dragEnterEvent(self, a0: QDragEnterEvent | None) -> None:
        if a0 and any(p.lower().endswith((".accdb", ".mdb", ".csv", ".xlsx")) for p in self._paden(a0)):
            a0.acceptProposedAction()

    def dropEvent(self, a0: QDropEvent | None) -> None:
        paden = self._paden(a0)
        if not paden:
            return
        huidig = self.tabs.currentWidget()
        if huidig is self.editor_tab:
            self.editor_tab.open_pad(paden[0])
        elif huidig is self.merge_tab:
            self.merge_tab.open_paden(paden)
        else:
            access = next((p for p in paden if is_access_bestand(p)), None)
            if access:
                self.open_access(access)

    # =========================================================
    # EXPORT-ACTIES
    # =========================================================
    def log_msg(self, msg: str) -> None:
        tijd = datetime.now().strftime("%H:%M:%S")
        self.log_entries.append((tijd, msg))
        self.txt_log.appendPlainText(f"[{tijd}] {tr(msg)}")

    def zet_tabellen_status(self, tekst: str, kleur: str = "ok") -> None:
        self.tabellen_status = (tekst, kleur)
        language_support.set_text(self.lbl_tabellen_status, tekst)
        self.lbl_tabellen_status.setStyleSheet(f"color: {themes.kleur(kleur)};")

    def zet_doelmap(self, pad: str) -> None:
        self.doelmap_pad = os.path.normpath(pad)
        language_support.set_text(self.lbl_doelmap_pad, self.doelmap_pad)

    def kies_access(self) -> None:
        startmap = str(self.settings.value("laatste_access_map", "") or "")
        pad, _ = QFileDialog.getOpenFileName(self, tr("Selecteer Access Database"), startmap,
                                             tr("Access Databases (*.accdb *.mdb);;Alle bestanden (*.*)"))
        if pad:
            self.open_access(pad)

    def open_access(self, pad: str) -> None:
        self.access_pad = os.path.normpath(pad)
        self.access_wachtwoord = ""
        self.settings.setValue("laatste_access_map", os.path.dirname(self.access_pad))
        language_support.set_text(self.lbl_access_pad, self.access_pad)
        self.log_msg(f"Bestand geselecteerd: {os.path.basename(self.access_pad)}")
        if not self.doelmap_pad:
            self.zet_doelmap(os.path.dirname(self.access_pad))
            self.log_msg(f"Doelmap ingesteld op map van het Access-bestand: {self.doelmap_pad}")
        self.laad_tabellen()

    def laad_tabellen(self) -> None:
        if not self.access_pad:
            return
        try:
            resultaat = met_wachtwoord(
                self, self.access_pad,
                lambda ww: haal_tabellen_op(self.access_pad, ww, self.chk_queries.isChecked()),
                self.access_wachtwoord)
        except Exception as e:
            resultaat = None
            self.log_msg(f"❌ Fout bij ophalen tabellen: {e}")
            self.zet_tabellen_status("Fout bij laden van tabellen", "fout")
        if resultaat is None:
            self.huidige_db_tabellen = []
            self.geselecteerde_tabellen = []
            if self.tabellen_status[1] != "fout":
                self.zet_tabellen_status("Geen tabellen geladen")
            return
        self.huidige_db_tabellen, self.access_wachtwoord = resultaat
        self.geselecteerde_tabellen = self.huidige_db_tabellen.copy()
        if self.huidige_db_tabellen:
            self.zet_tabellen_status(f"Alle ({len(self.geselecteerde_tabellen)}) tabellen geselecteerd")
        else:
            self.zet_tabellen_status("Geen tabellen gevonden in dit bestand", "fout")

    def kies_tabellen(self) -> None:
        if not self.access_pad or not self.huidige_db_tabellen:
            QMessageBox.warning(self, tr("Let op"), tr("Kies eerst een geldig Access-bestand."))
            return

        dialog = TableSelectionDialog(self.access_pad, self.huidige_db_tabellen, self.geselecteerde_tabellen,
                                      self, self.access_wachtwoord)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        self.geselecteerde_tabellen = dialog.geselecteerde
        aantal = len(self.geselecteerde_tabellen)
        if aantal == len(self.huidige_db_tabellen):
            self.zet_tabellen_status(f"Alle ({aantal}) tabellen geselecteerd")
        elif aantal == 0:
            self.zet_tabellen_status("Let op: 0 tabellen geselecteerd!", "fout")
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
            "wachtwoord": self.access_wachtwoord,
            "doelmap": self.doelmap_pad,
            "access_naam": access_naam,
            "naam_zonder_ext": naam_zonder_ext,
            "export_type": export_type,
            "csv_decimaal": self.cmb_decimaal.currentData() or ",",
            "tabellen": self.geselecteerde_tabellen.copy(),
        })
        self.log_msg(f"Toegevoegd aan wachtrij: {access_naam} als {export_type.upper()}")

        self.access_pad = ""
        self.access_wachtwoord = ""
        language_support.set_text(self.lbl_access_pad, "Geen Access-bestand geselecteerd")
        self.huidige_db_tabellen = []
        self.geselecteerde_tabellen = []
        self.zet_tabellen_status("Geen tabellen geladen")
        self.lbl_tabellen_status.setToolTip("")

    def geselecteerde_rijen(self) -> list[int]:
        model = self.table_queue.selectionModel()
        return sorted(index.row() for index in model.selectedRows()) if model else []

    def _verwijder_rijen(self, rijen: list[int]) -> None:
        for row in sorted(rijen, reverse=True):
            self.log_msg(f"Verwijderd uit wachtrij: {self.jobs[row]['access_naam']}")
            self.table_queue.removeRow(row)
            del self.jobs[row]

    def verwijder_selectie(self) -> None:
        if not self.conversie_bezig():
            self._verwijder_rijen(self.geselecteerde_rijen())

    def wis_afgeronde_taken(self) -> None:
        if not self.conversie_bezig():
            self._verwijder_rijen([r for r in range(len(self.jobs)) if self.status_van_rij(r) == STATUS_KLAAR])

    def open_job_in_editor(self, row: int, _column: int = 0) -> None:
        if not 0 <= row < len(self.jobs) or self.status_van_rij(row) != STATUS_KLAAR:
            return
        job = self.jobs[row]
        if job["export_type"] == "xlsx":
            pad = os.path.join(job["doelmap"], f"{job['naam_zonder_ext']}.xlsx")
        else:
            pad, _ = QFileDialog.getOpenFileName(self, tr("Open bestand"), job["doelmap"], tr("CSV-bestanden (*.csv)"))
        if pad and os.path.exists(pad):
            self.tabs.setCurrentWidget(self.editor_tab)
            self.editor_tab.open_pad(pad)

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
        for knop in (self.btn_conv_sel, self.btn_conv_all, self.btn_verwijder, self.btn_wis_klaar):
            knop.setEnabled(not bezig)
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
        self.kleur_status(row)

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
        if not self.editor_tab.mag_sluiten() or not self.merge_tab.mag_sluiten():
            if a0:
                a0.ignore()
            return
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


__all__ = ["AccessExporterGUI", "QApplication", "toon_onverwachte_fout"]
