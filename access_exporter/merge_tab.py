"""Tabblad Terugzetten naar Access: bewerkte CSV/Excel-exports samenvoegen in een NIEUWE kopie van de database.
De originele database wordt nooit gewijzigd."""

import os
from collections.abc import Callable
from contextlib import closing
from typing import Any

from PyQt6.QtCore import QSettings, Qt
from PyQt6.QtGui import QColor
from PyQt6.QtWidgets import (
    QAbstractItemView, QButtonGroup, QCheckBox, QComboBox, QFileDialog, QGroupBox, QHBoxLayout, QHeaderView, QLabel,
    QMessageBox, QProgressBar, QPushButton, QRadioButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from . import themes
from .core import haal_tabellen_op, is_access_bestand, is_vergrendeld, verbind_access, voorgestelde_kopienaam
from .dialogs import met_wachtwoord
from .merge import (
    MODUS_BIJWERKEN, MODUS_TOEVOEGEN, MODUS_VERVANGEN, Bron, Gestopt, TabelPlan, bereken_plannen, bronnen_uit_bestand,
    raad_tabel, zet_terug_in_kopie,
)
from .translations import language_support, tr
from .worker import TaakWorker

OVERSLAAN = "(overslaan)"
MAX_MELDINGEN = 200
MODI = (
    (MODUS_BIJWERKEN, "Bijwerken en toevoegen (op primaire sleutel)",
     "Bestaande rijen worden bijgewerkt op basis van de primaire sleutel; nieuwe rijen worden toegevoegd"),
    (MODUS_TOEVOEGEN, "Alleen nieuwe rijen toevoegen",
     "Rijen waarvan de sleutel al bestaat worden overgeslagen; er wordt niets gewijzigd of verwijderd"),
    (MODUS_VERVANGEN, "Tabelinhoud vervangen",
     "Alle rijen in de gekoppelde tabellen worden verwijderd en vervangen door de rijen uit het bestand"),
)


class MergeTab(QWidget):
    def __init__(self, log: Callable[[str], None], settings: QSettings, parent: QWidget | None = None):
        super().__init__(parent)
        self.log = log
        self.settings = settings
        self.db_pad = ""
        self.wachtwoord = ""
        self.tabellen: list[str] = []
        self.bronnen: list[Bron] = []
        self.plannen: list[TabelPlan] = []
        self.worker: TaakWorker | None = None

        layout = QVBoxLayout(self)
        uitleg = QLabel("Zet bewerkte CSV/Excel-exports terug in Access. Er wordt altijd een nieuwe kopie van de "
                        "database gemaakt; het origineel blijft ongewijzigd.")
        uitleg.setWordWrap(True)
        uitleg.setObjectName("infoLabel")
        layout.addWidget(uitleg)

        # Stap 1: database
        groep_db = QGroupBox("1. Originele Access-database")
        rij_db = QHBoxLayout(groep_db)
        self.btn_db = QPushButton("🗄 Kies database...")
        self.btn_db.clicked.connect(self.kies_database)
        self.lbl_db = QLabel("Geen database geselecteerd")
        rij_db.addWidget(self.btn_db)
        rij_db.addWidget(self.lbl_db, 1)
        layout.addWidget(groep_db)

        # Stap 2: bestanden en koppeling
        groep_bestanden = QGroupBox("2. Bewerkte bestanden en doeltabellen")
        lay_bestanden = QVBoxLayout(groep_bestanden)
        rij_knoppen = QHBoxLayout()
        self.btn_bestanden = QPushButton("📂 Bestanden toevoegen...")
        self.btn_bestanden.clicked.connect(self.kies_bestanden)
        self.btn_bestanden_wissen = QPushButton("🗑 Lijst wissen")
        self.btn_bestanden_wissen.clicked.connect(self.wis_bestanden)
        rij_knoppen.addWidget(self.btn_bestanden)
        rij_knoppen.addWidget(self.btn_bestanden_wissen)
        rij_knoppen.addStretch()
        lay_bestanden.addLayout(rij_knoppen)
        self.tbl_koppeling = QTableWidget(0, 3)
        self.tbl_koppeling.setHorizontalHeaderLabels([tr("Bestand / blad"), tr("Rijen"), tr("Doeltabel in Access")])
        koppen = self.tbl_koppeling.horizontalHeader()
        if koppen is not None:
            koppen.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
            koppen.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
            koppen.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.tbl_koppeling.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)
        self.tbl_koppeling.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        lay_bestanden.addWidget(self.tbl_koppeling)
        layout.addWidget(groep_bestanden, 1)

        # Stap 3: opties
        groep_opties = QGroupBox("3. Hoe samenvoegen?")
        lay_opties = QVBoxLayout(groep_opties)
        self.modus_groep = QButtonGroup(self)
        self.modus_knoppen: dict[str, QRadioButton] = {}
        for modus, label, uitleg_modus in MODI:
            knop = QRadioButton(label)
            knop.setToolTip(uitleg_modus)
            self.modus_groep.addButton(knop)
            self.modus_knoppen[modus] = knop
            lay_opties.addWidget(knop)
            knop.toggled.connect(self.opties_gewijzigd)
        self.chk_verwijderen = QCheckBox("Rijen die niet meer in het bestand staan ook verwijderen")
        self.chk_verwijderen.setToolTip("Alleen bij 'Bijwerken en toevoegen'. Er wordt niets verwijderd als er fouten zijn.")
        self.chk_verwijderen.toggled.connect(self.opties_gewijzigd)
        lay_opties.addWidget(self.chk_verwijderen)
        layout.addWidget(groep_opties)

        # Stap 4: voorbeeld en uitvoeren
        groep_voorbeeld = QGroupBox("4. Controleren en samenvoegen")
        lay_voorbeeld = QVBoxLayout(groep_voorbeeld)
        rij_acties = QHBoxLayout()
        self.btn_voorbeeld = QPushButton("🔍 Voorbeeld berekenen")
        self.btn_voorbeeld.setToolTip("Toont wat er per tabel verandert, zonder iets op te slaan")
        self.btn_voorbeeld.clicked.connect(self.bereken_voorbeeld)
        self.btn_samenvoegen = QPushButton("💾 Samenvoegen in nieuwe kopie...")
        self.btn_samenvoegen.setObjectName("successButton")
        self.btn_samenvoegen.clicked.connect(self.samenvoegen)
        self.btn_stop = QPushButton("⏹ Stop")
        self.btn_stop.setObjectName("dangerButton")
        self.btn_stop.clicked.connect(self.stop)
        self.voortgang = QProgressBar()
        self.voortgang.setVisible(False)
        rij_acties.addWidget(self.btn_voorbeeld)
        rij_acties.addWidget(self.btn_samenvoegen)
        rij_acties.addWidget(self.btn_stop)
        rij_acties.addWidget(self.voortgang, 1)
        lay_voorbeeld.addLayout(rij_acties)
        self.tbl_voorbeeld = QTableWidget(0, 6)
        self.tbl_voorbeeld.setHorizontalHeaderLabels(
            [tr(k) for k in ("Tabel", "Toevoegen", "Wijzigen", "Verwijderen", "Ongewijzigd", "Fouten")])
        koppen = self.tbl_voorbeeld.horizontalHeader()
        if koppen is not None:
            koppen.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.tbl_voorbeeld.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tbl_voorbeeld.setMaximumHeight(160)
        lay_voorbeeld.addWidget(self.tbl_voorbeeld)
        self.lbl_samenvatting = QLabel("")
        self.lbl_samenvatting.setWordWrap(True)
        lay_voorbeeld.addWidget(self.lbl_samenvatting)
        layout.addWidget(groep_voorbeeld)

        modus = str(self.settings.value("merge_modus", MODUS_BIJWERKEN) or MODUS_BIJWERKEN)
        self.modus_knoppen.get(modus, self.modus_knoppen[MODUS_BIJWERKEN]).setChecked(True)
        self.werk_knoppen_bij()

    # --- Algemeen ---
    def bezig(self) -> bool:
        return self.worker is not None and self.worker.isRunning()

    def mag_sluiten(self) -> bool:
        if not self.bezig():
            return True
        QMessageBox.warning(self, tr("Bezig"), tr("Wacht tot het samenvoegen klaar is of druk op Stop."))
        return False

    def huidige_modus(self) -> str:
        return next((m for m, knop in self.modus_knoppen.items() if knop.isChecked()), MODUS_BIJWERKEN)

    def koppelingen(self) -> dict[str, list[Bron]]:
        resultaat: dict[str, list[Bron]] = {}
        for rij, bron in enumerate(self.bronnen):
            combo = self.tbl_koppeling.cellWidget(rij, 2)
            if isinstance(combo, QComboBox) and combo.currentIndex() > 0:
                resultaat.setdefault(combo.currentText(), []).append(bron)
        return resultaat

    def werk_knoppen_bij(self) -> None:
        bezig = self.bezig()
        klaar_voor_plan = bool(self.db_pad and self.koppelingen()) and not bezig
        uitvoerbaar = [p for p in self.plannen if not p.geblokkeerd and p.aantal_bewerkingen]
        for widget in (self.btn_db, self.btn_bestanden, self.btn_bestanden_wissen, self.tbl_koppeling,
                       self.chk_verwijderen, *self.modus_knoppen.values()):
            widget.setEnabled(not bezig)
        self.chk_verwijderen.setEnabled(not bezig and self.huidige_modus() == MODUS_BIJWERKEN)
        self.btn_voorbeeld.setEnabled(klaar_voor_plan)
        self.btn_samenvoegen.setEnabled(klaar_voor_plan and bool(uitvoerbaar))
        self.btn_stop.setEnabled(bezig)

    def vervallen_voorbeeld(self) -> None:
        self.plannen = []
        self.tbl_voorbeeld.setRowCount(0)
        language_support.set_text(self.lbl_samenvatting, "")
        self.werk_knoppen_bij()

    def opties_gewijzigd(self) -> None:
        self.settings.setValue("merge_modus", self.huidige_modus())
        self.vervallen_voorbeeld()

    def vertaal(self) -> None:
        self.tbl_koppeling.setHorizontalHeaderLabels([tr("Bestand / blad"), tr("Rijen"), tr("Doeltabel in Access")])
        self.tbl_voorbeeld.setHorizontalHeaderLabels(
            [tr(k) for k in ("Tabel", "Toevoegen", "Wijzigen", "Verwijderen", "Ongewijzigd", "Fouten")])
        for rij in range(self.tbl_koppeling.rowCount()):
            combo = self.tbl_koppeling.cellWidget(rij, 2)
            if isinstance(combo, QComboBox):
                combo.setItemText(0, tr(OVERSLAAN))
        if self.plannen:
            self.toon_plannen()

    def kleur_opnieuw(self) -> None:
        if self.plannen:
            self.toon_plannen()

    # --- Database en bestanden ---
    def kies_database(self) -> None:
        startmap = str(self.settings.value("laatste_access_map", "") or "")
        pad, _ = QFileDialog.getOpenFileName(self, tr("Selecteer Access Database"), startmap,
                                             tr("Access Databases (*.accdb *.mdb);;Alle bestanden (*.*)"))
        if pad:
            self.open_database(pad)

    def open_database(self, pad: str) -> None:
        pad = os.path.normpath(pad)
        try:
            resultaat = met_wachtwoord(self, pad, lambda ww: haal_tabellen_op(pad, ww))
        except Exception as e:
            QMessageBox.critical(self, tr("Fout"), tr(f"Kan bestand niet openen: {str(e).splitlines()[0]}"))
            return
        if resultaat is None:
            return
        self.tabellen, self.wachtwoord = resultaat
        self.db_pad = pad
        self.settings.setValue("laatste_access_map", os.path.dirname(pad))
        language_support.set_text(self.lbl_db, pad)
        self.vul_koppeling()

    def kies_bestanden(self) -> None:
        startmap = str(self.settings.value("laatste_export_map", "") or "")
        paden, _ = QFileDialog.getOpenFileNames(self, tr("Bestanden toevoegen"), startmap,
                                                tr("Export-bestanden (*.xlsx *.csv);;Alle bestanden (*.*)"))
        if paden:
            self.open_paden(paden)

    def open_paden(self, paden: list[str]) -> None:
        databases = [p for p in paden if is_access_bestand(p)]
        if databases:
            self.open_database(databases[0])
        for pad in (os.path.normpath(p) for p in paden if not is_access_bestand(p)):
            if any(b.pad == pad for b in self.bronnen):
                continue
            try:
                self.bronnen += bronnen_uit_bestand(pad)
            except Exception as e:
                self.log(f"❌ Fout bij {os.path.basename(pad)}: {str(e).splitlines()[0]}")
                continue
            self.settings.setValue("laatste_export_map", os.path.dirname(pad))
        self.vul_koppeling()

    def wis_bestanden(self) -> None:
        self.bronnen = []
        self.vul_koppeling()

    def vul_koppeling(self) -> None:
        vorige = {str(c.property("bron_label")): c.currentText() for c in self._combos() if c.currentIndex() > 0}
        self.tbl_koppeling.setRowCount(len(self.bronnen))
        for rij, bron in enumerate(self.bronnen):
            self.tbl_koppeling.setItem(rij, 0, QTableWidgetItem(bron.label))
            aantal = QTableWidgetItem(str(len(bron.rijen)))
            aantal.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
            self.tbl_koppeling.setItem(rij, 1, aantal)
            combo = QComboBox()
            combo.setProperty("bron_label", bron.label)
            combo.addItem(tr(OVERSLAAN))
            combo.addItems(self.tabellen)
            keuze = vorige.get(bron.label) or raad_tabel(bron, self.tabellen)
            if keuze and keuze in self.tabellen:
                combo.setCurrentIndex(self.tabellen.index(keuze) + 1)
            combo.currentIndexChanged.connect(self.vervallen_voorbeeld)
            self.tbl_koppeling.setCellWidget(rij, 2, combo)
        self.vervallen_voorbeeld()

    def _combos(self) -> list[QComboBox]:
        combos = []
        for rij in range(self.tbl_koppeling.rowCount()):
            combo = self.tbl_koppeling.cellWidget(rij, 2)
            if isinstance(combo, QComboBox):
                combos.append(combo)
        return combos

    # --- Voorbeeld en uitvoeren ---
    def _start(self, functie: Callable[..., Any], klaar: Callable[[Any], None]) -> None:
        self.worker = TaakWorker(functie)
        self.worker.voortgang_signal.connect(self._voortgang)
        self.worker.klaar_signal.connect(klaar)
        self.worker.fout_signal.connect(self._fout)
        self.worker.finished.connect(self._worker_klaar)
        self.voortgang.setRange(0, 0)
        self.voortgang.setVisible(True)
        self.worker.start()
        self.werk_knoppen_bij()

    def _voortgang(self, klaar: int, totaal: int) -> None:
        self.voortgang.setRange(0, max(totaal, 1))
        self.voortgang.setValue(klaar)

    def _worker_klaar(self) -> None:
        self.voortgang.setVisible(False)
        self.worker = None
        self.werk_knoppen_bij()

    def _fout(self, fout: BaseException) -> None:
        if isinstance(fout, Gestopt):
            self.log("⏹ Samenvoegen gestopt; er is niets opgeslagen.")
            return
        melding = str(fout).splitlines()[0] if str(fout) else type(fout).__name__
        self.log(f"❌ Samenvoegen mislukt: {melding}")
        QMessageBox.critical(self, tr("Fout"), tr(f"Samenvoegen mislukt: {melding}"))

    def stop(self) -> None:
        if self.worker is not None:
            self.worker.requestInterruption()

    def bereken_voorbeeld(self) -> None:
        koppelingen = self.koppelingen()
        if not self.db_pad or not koppelingen:
            return
        pad, wachtwoord, modus, verwijderen = self.db_pad, self.wachtwoord, self.huidige_modus(), self.chk_verwijderen.isChecked()

        def taak(stop: Callable[[], bool], voortgang: Callable[[int, int], None]) -> list[TabelPlan]:
            with closing(verbind_access(pad, wachtwoord)) as conn:
                return bereken_plannen(conn, koppelingen, modus, verwijderen, stop)

        self.log(f"Voorbeeld berekenen voor {os.path.basename(pad)}...")
        self._start(taak, self._voorbeeld_klaar)

    def _voorbeeld_klaar(self, plannen: list[TabelPlan]) -> None:
        self.plannen = plannen
        self.toon_plannen()
        for plan in plannen:
            for melding in plan.fouten[:MAX_MELDINGEN]:
                self.log(f"⚠ {plan.tabel}: {melding}")
            if len(plan.fouten) > MAX_MELDINGEN:
                self.log(f"⚠ {plan.tabel}: nog {len(plan.fouten) - MAX_MELDINGEN} fouten")
            for melding in plan.opmerkingen:
                self.log(f"ℹ {plan.tabel}: {melding}")
        self.werk_knoppen_bij()

    def toon_plannen(self) -> None:
        self.tbl_voorbeeld.setRowCount(len(self.plannen))
        for rij, plan in enumerate(self.plannen):
            waarden = [plan.tabel, len(plan.invoegen),
                       len(plan.wijzigen), tr("alles") if plan.alles_verwijderen else len(plan.verwijderen),
                       plan.ongewijzigd, len(plan.fouten)]
            for kolom, waarde in enumerate(waarden):
                item = QTableWidgetItem(str(waarde))
                if kolom:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                if plan.geblokkeerd or (kolom == 5 and plan.fouten):
                    item.setForeground(QColor(themes.kleur("fout")))
                    if plan.fouten:
                        item.setToolTip(tr(plan.fouten[0]))
                self.tbl_voorbeeld.setItem(rij, kolom, item)
        totaal = sum(p.aantal_bewerkingen for p in self.plannen if not p.geblokkeerd)
        fouten = sum(len(p.fouten) for p in self.plannen)
        tekst = f"Totaal {totaal} wijzigingen klaar om samen te voegen."
        if fouten:
            tekst += f" {fouten} rijen of tabellen worden overgeslagen vanwege fouten (zie log)."
        language_support.set_text(self.lbl_samenvatting, tekst)

    def samenvoegen(self) -> None:
        uitvoerbaar = [p for p in self.plannen if not p.geblokkeerd and p.aantal_bewerkingen]
        if not uitvoerbaar or self.bezig():
            return
        startpad = voorgestelde_kopienaam(self.db_pad, "merged")
        doel, _ = QFileDialog.getSaveFileName(self, tr("Samengevoegde database opslaan als"), startpad,
                                              tr("Access Databases (*.accdb *.mdb)"))
        if not doel:
            return
        doel = os.path.normpath(doel)
        if os.path.normcase(os.path.abspath(doel)) == os.path.normcase(os.path.abspath(self.db_pad)):
            QMessageBox.warning(self, tr("Fout"), tr("Kies een andere naam: het origineel wordt nooit overschreven."))
            return
        if is_vergrendeld(self.db_pad) and QMessageBox.question(
                self, tr("Database in gebruik"),
                tr("De database lijkt geopend te zijn in Access. Sluit Access voor een betrouwbare kopie. Toch doorgaan?")
        ) != QMessageBox.StandardButton.Yes:
            return
        totaal = sum(p.aantal_bewerkingen for p in uitvoerbaar)
        verwijderen = sum(len(p.verwijderen) for p in uitvoerbaar) + sum(p.alles_verwijderen for p in uitvoerbaar)
        vraag = f"{totaal} wijzigingen samenvoegen in {os.path.basename(doel)}?"
        if verwijderen:
            vraag += f" Let op: dit omvat {verwijderen} verwijderingen."
        if QMessageBox.question(self, tr("Bevestigen"), tr(vraag)) != QMessageBox.StandardButton.Yes:
            return
        bron, wachtwoord = self.db_pad, self.wachtwoord

        def taak(stop: Callable[[], bool], voortgang: Callable[[int, int], None]) -> str:
            zet_terug_in_kopie(bron, doel, uitvoerbaar, wachtwoord, stop, voortgang)
            return doel

        self.log(f"Samenvoegen in nieuwe kopie: {os.path.basename(doel)}...")
        self._start(taak, self._samenvoegen_klaar)

    def _samenvoegen_klaar(self, doel: str) -> None:
        self.log(f"✅ Samengevoegd in nieuwe kopie: {doel}")
        QMessageBox.information(self, tr("Klaar"), tr(f"Samengevoegd in nieuwe kopie: {doel}"))
