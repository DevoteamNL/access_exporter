"""Tabblad Bewerken: Access-tabellen en geëxporteerde CSV/Excel-bestanden bekijken en bewerken.
Opslaan gebeurt altijd via 'Opslaan als': een nieuw bestand of een nieuwe kopie van de database."""

import os
import re
from collections.abc import Callable
from contextlib import closing
from datetime import datetime
from decimal import Decimal
from typing import Any

from PyQt6.QtCore import QAbstractTableModel, QModelIndex, QSettings, QSortFilterProxyModel, Qt, pyqtSignal
from PyQt6.QtGui import QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QFileDialog, QHBoxLayout, QLabel, QLineEdit, QMessageBox,
    QPushButton, QTableView, QVBoxLayout, QWidget,
)

from .core import (
    haal_tabellen_op, is_access_bestand, is_vergrendeld, lees_rijen, tel_rijen, verbind_access,
    voorgestelde_kopienaam,
)
from .dialogs import met_wachtwoord
from .file_io import lees_export, schrijf_csv, schrijf_excel
from .merge import MODUS_BIJWERKEN, OmzetFout, Schema, _getaltekst, _naar_datum, gelijk, lees_schema, maak_plan, zet_om, zet_terug_in_kopie
from .translations import language_support, tr

MAX_EDITOR_RIJEN = 100_000
_GEHEEL = re.compile(r"^[+-]?\d+$")


def weergave(waarde: Any) -> str:
    if waarde is None:
        return ""
    if isinstance(waarde, bool):
        return tr("Ja") if waarde else tr("Nee")
    if isinstance(waarde, datetime):
        if (waarde.hour, waarde.minute, waarde.second) == (0, 0, 0):
            return waarde.date().isoformat()
        return waarde.isoformat(" ", "seconds")
    if isinstance(waarde, (bytes, bytearray, memoryview)):
        return f"<{len(waarde)} bytes>"
    if isinstance(waarde, float):
        return repr(waarde) if not waarde.is_integer() else str(int(waarde)) if abs(waarde) < 1e15 else repr(waarde)
    return str(waarde)


def behoud_type(oud: Any, tekst: str) -> Any:
    """Bestandseditor: een getal of datum blijft een getal of datum als de nieuwe tekst dat toelaat."""
    if tekst.strip() == "":
        return None
    try:
        if isinstance(oud, int) and not isinstance(oud, bool) and _GEHEEL.match(tekst.strip()):
            return int(tekst)
        if isinstance(oud, (float, Decimal)) and not isinstance(oud, bool):
            return float(_getaltekst(tekst))
        if isinstance(oud, datetime):
            return _naar_datum(tekst)
    except (ValueError, OmzetFout):
        pass
    return tekst


class TabelModel(QAbstractTableModel):
    """Bewerkbaar tabelmodel over een lijst van rijen (alleen zichtbare rijen worden getekend)."""
    gewijzigd_signal = pyqtSignal()
    fout_signal = pyqtSignal(str)

    def __init__(self, kolommen: list[str], rijen: list[list[Any]], bewerkbaar: bool = True,
                 alleen_lezen: set[int] | None = None, omzetter: Callable[[int, str], Any] | None = None):
        super().__init__()
        self.kolommen = kolommen
        self.rijen = rijen
        self.bewerkbaar = bewerkbaar
        self.alleen_lezen = alleen_lezen or set()
        self.omzetter = omzetter
        self.gewijzigd = False

    def rowCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: B008
        return 0 if parent.isValid() else len(self.rijen)

    def columnCount(self, parent: QModelIndex = QModelIndex()) -> int:  # noqa: B008
        return 0 if parent.isValid() else len(self.kolommen)

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if not index.isValid():
            return None
        if role in (Qt.ItemDataRole.DisplayRole, Qt.ItemDataRole.EditRole):
            rij = self.rijen[index.row()]
            return weergave(rij[index.column()] if index.column() < len(rij) else None)
        return None

    def headerData(self, section: int, orientation: Qt.Orientation, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return self.kolommen[section] if section < len(self.kolommen) else None
        return section + 1

    def flags(self, index: QModelIndex) -> Qt.ItemFlag:
        basis = Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable
        if self.bewerkbaar and index.column() not in self.alleen_lezen:
            basis |= Qt.ItemFlag.ItemIsEditable
        return basis

    def setData(self, index: QModelIndex, value: Any, role: int = Qt.ItemDataRole.EditRole) -> bool:
        if not index.isValid() or role != Qt.ItemDataRole.EditRole or not self.bewerkbaar:
            return False
        rij = self.rijen[index.row()]
        while len(rij) < len(self.kolommen):
            rij.append(None)
        oud = rij[index.column()]
        tekst = "" if value is None else str(value)
        if self.omzetter is not None:
            try:
                nieuw = self.omzetter(index.column(), tekst)
            except OmzetFout as e:
                self.fout_signal.emit(f"Kolom '{self.kolommen[index.column()]}': {e}")
                return False
        else:
            nieuw = behoud_type(oud, tekst)
        if gelijk(oud, nieuw) and type(oud) is type(nieuw):
            return True
        rij[index.column()] = nieuw
        self.dataChanged.emit(index, index)
        self._markeer()
        return True

    def _markeer(self) -> None:
        self.gewijzigd = True
        self.gewijzigd_signal.emit()

    def voeg_rij_toe(self) -> int:
        positie = len(self.rijen)
        self.beginInsertRows(QModelIndex(), positie, positie)
        self.rijen.append([None] * len(self.kolommen))
        self.endInsertRows()
        self._markeer()
        return positie

    def verwijder_rijen(self, rijen: list[int]) -> None:
        for rij in sorted(set(rijen), reverse=True):
            self.beginRemoveRows(QModelIndex(), rij, rij)
            del self.rijen[rij]
            self.endRemoveRows()
        if rijen:
            self._markeer()


class AccessTabel:
    """Een geladen Access-tabel met schema en de oorspronkelijke rijen (om wijzigingen te bepalen)."""

    def __init__(self, schema: Schema, kolommen: list[str], origineel: list[tuple[Any, ...]], totaal: int):
        self.schema = schema
        self.kolommen = kolommen
        self.origineel = origineel
        self.totaal = totaal


class EditorTab(QWidget):
    def __init__(self, log: Callable[[str], None], settings: QSettings, parent: QWidget | None = None):
        super().__init__(parent)
        self.log = log
        self.settings = settings
        self.bron_pad = ""
        self.bron_soort = ""  # "access" of "bestand"
        self.wachtwoord = ""
        self.modellen: dict[str, TabelModel] = {}
        self.access_tabellen: dict[str, AccessTabel] = {}
        self._status_bron = ""

        layout = QVBoxLayout(self)
        rij1 = QHBoxLayout()
        self.btn_open_access = QPushButton("🗄 Open Access-database...")
        self.btn_open_access.setToolTip("Bewerk tabellen direct; opslaan maakt altijd een nieuwe kopie van de database")
        self.btn_open_access.clicked.connect(self.kies_access)
        self.btn_open_bestand = QPushButton("📂 Open CSV/Excel-bestand...")
        self.btn_open_bestand.setToolTip("Bewerk een geëxporteerd bestand; opslaan maakt altijd een nieuw bestand")
        self.btn_open_bestand.clicked.connect(self.kies_bestand)
        self.lbl_bron = QLabel("Geen database of bestand geopend")
        self.lbl_bron.setObjectName("infoLabel")
        rij1.addWidget(self.btn_open_access)
        rij1.addWidget(self.btn_open_bestand)
        rij1.addWidget(self.lbl_bron, 1)
        layout.addLayout(rij1)

        rij2 = QHBoxLayout()
        lbl_tabel = QLabel("Tabel / blad:")
        self.cmb_tabel = QComboBox()
        self.cmb_tabel.setMinimumWidth(220)
        self.cmb_tabel.currentIndexChanged.connect(self.toon_tabel)
        lbl_tabel.setBuddy(self.cmb_tabel)
        self.txt_zoek = QLineEdit()
        self.txt_zoek.setPlaceholderText("🔍 Zoeken in alle kolommen... (Ctrl+F)")
        self.txt_zoek.setClearButtonEnabled(True)
        self.txt_zoek.textChanged.connect(self.zoek)
        self.btn_rij_toevoegen = QPushButton("➕ Rij toevoegen")
        self.btn_rij_toevoegen.clicked.connect(self.rij_toevoegen)
        self.btn_rijen_verwijderen = QPushButton("➖ Rijen verwijderen")
        self.btn_rijen_verwijderen.setToolTip("Verwijder de geselecteerde rijen (Delete)")
        self.btn_rijen_verwijderen.clicked.connect(self.rijen_verwijderen)
        self.btn_opslaan = QPushButton("💾 Opslaan als...")
        self.btn_opslaan.setObjectName("successButton")
        self.btn_opslaan.clicked.connect(self.opslaan_als)
        for widget in (lbl_tabel, self.cmb_tabel):
            rij2.addWidget(widget)
        rij2.addWidget(self.txt_zoek, 1)
        for widget in (self.btn_rij_toevoegen, self.btn_rijen_verwijderen, self.btn_opslaan):
            rij2.addWidget(widget)
        layout.addLayout(rij2)

        self.lbl_opmerking = QLabel("")
        self.lbl_opmerking.setWordWrap(True)
        self.lbl_opmerking.setVisible(False)
        layout.addWidget(self.lbl_opmerking)

        self.proxy = QSortFilterProxyModel(self)
        self.proxy.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.proxy.setFilterKeyColumn(-1)
        self.view = QTableView()
        self.view.setModel(self.proxy)
        self.view.setAlternatingRowColors(True)
        self.view.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.view.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked | QAbstractItemView.EditTrigger.EditKeyPressed
                                  | QAbstractItemView.EditTrigger.AnyKeyPressed)
        layout.addWidget(self.view, 1)

        self.lbl_status = QLabel("")
        layout.addWidget(self.lbl_status)

        QShortcut(QKeySequence.StandardKey.Find, self).activated.connect(self.txt_zoek.setFocus)
        QShortcut(QKeySequence.StandardKey.Delete, self.view).activated.connect(self.rijen_verwijderen)
        self.werk_knoppen_bij()

    # --- Hulpfuncties ---
    @property
    def model(self) -> TabelModel | None:
        bron = self.proxy.sourceModel()
        return bron if isinstance(bron, TabelModel) else None

    def heeft_wijzigingen(self) -> bool:
        return any(m.gewijzigd for m in self.modellen.values())

    def werk_knoppen_bij(self) -> None:
        model = self.model
        bewerkbaar = model is not None and model.bewerkbaar
        self.btn_rij_toevoegen.setEnabled(bewerkbaar)
        self.btn_rijen_verwijderen.setEnabled(bewerkbaar)
        self.btn_opslaan.setEnabled(bool(self.modellen))
        self.werk_status_bij()

    def werk_status_bij(self) -> None:
        model = self.model
        if model is None:
            self._status_bron = ""
        else:
            tekst = f"{len(model.rijen)} rijen · {len(model.kolommen)} kolommen"
            if self.heeft_wijzigingen():
                tekst += " · niet-opgeslagen wijzigingen"
            self._status_bron = tekst
        self.lbl_status.setText(tr(self._status_bron) if self._status_bron else "")

    def vertaal(self) -> None:
        self.werk_status_bij()
        self._zet_opmerking(str(self.lbl_opmerking.property("source_text") or ""))
        viewport = self.view.viewport()
        if viewport is not None:
            viewport.update()

    def _zet_opmerking(self, tekst: str) -> None:
        language_support.set_text(self.lbl_opmerking, tekst)
        self.lbl_opmerking.setVisible(bool(tekst))

    def mag_vervangen(self) -> bool:
        """Vraagt wat er met niet-opgeslagen wijzigingen moet gebeuren. False = annuleren."""
        if not self.heeft_wijzigingen():
            return True
        antwoord = QMessageBox.question(
            self, tr("Niet-opgeslagen wijzigingen"),
            tr(f"Wijzigingen in {os.path.basename(self.bron_pad)} opslaan?"),
            QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Discard | QMessageBox.StandardButton.Cancel)
        if antwoord == QMessageBox.StandardButton.Save:
            return self.opslaan_als()
        return antwoord == QMessageBox.StandardButton.Discard

    mag_sluiten = mag_vervangen

    # --- Openen ---
    def kies_access(self) -> None:
        startmap = str(self.settings.value("laatste_access_map", "") or "")
        pad, _ = QFileDialog.getOpenFileName(self, tr("Selecteer Access Database"), startmap,
                                             tr("Access Databases (*.accdb *.mdb);;Alle bestanden (*.*)"))
        if pad:
            self.open_pad(pad)

    def kies_bestand(self) -> None:
        startmap = str(self.settings.value("laatste_export_map", "") or "")
        pad, _ = QFileDialog.getOpenFileName(self, tr("Open bestand"), startmap,
                                             tr("Export-bestanden (*.xlsx *.csv);;Alle bestanden (*.*)"))
        if pad:
            self.open_pad(pad)

    def open_pad(self, pad: str, wachtwoord: str = "") -> None:
        if not self.mag_vervangen():
            return
        pad = os.path.normpath(pad)
        if is_access_bestand(pad):
            self.open_access(pad, wachtwoord)
        else:
            self.open_bestand(pad)

    def _reset(self, pad: str, soort: str) -> None:
        self.bron_pad, self.bron_soort = pad, soort
        self.modellen.clear()
        self.access_tabellen.clear()
        self.proxy.setSourceModel(None)
        self.txt_zoek.clear()
        self._zet_opmerking("")
        language_support.set_text(self.lbl_bron, pad)

    def open_bestand(self, pad: str) -> None:
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            tabellen = lees_export(pad)
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, tr("Fout"), tr(f"Kan bestand niet openen: {str(e).splitlines()[0]}"))
            return
        QApplication.restoreOverrideCursor()
        self.settings.setValue("laatste_export_map", os.path.dirname(pad))
        self._reset(pad, "bestand")
        for naam, (kolommen, rijen) in tabellen.items():
            self.modellen[naam] = self._nieuw_model(TabelModel(kolommen, rijen))
        self._vul_tabellen(list(tabellen))
        self.log(f"Geopend in editor: {os.path.basename(pad)}")

    def open_access(self, pad: str, wachtwoord: str = "") -> None:
        try:
            resultaat = met_wachtwoord(self, pad, lambda ww: haal_tabellen_op(pad, ww), wachtwoord)
        except Exception as e:
            QMessageBox.critical(self, tr("Fout"), tr(f"Kan bestand niet openen: {str(e).splitlines()[0]}"))
            return
        if resultaat is None:
            return
        tabellen, self.wachtwoord = resultaat
        self.settings.setValue("laatste_access_map", os.path.dirname(pad))
        self._reset(pad, "access")
        self._vul_tabellen(tabellen)
        self.log(f"Geopend in editor: {os.path.basename(pad)}")

    def _nieuw_model(self, model: TabelModel) -> TabelModel:
        model.gewijzigd_signal.connect(self.werk_status_bij)
        model.fout_signal.connect(lambda melding: self.log(f"⚠ {self.cmb_tabel.currentText()}: {melding}"))
        model.fout_signal.connect(lambda melding: QMessageBox.warning(self, tr("Ongeldige waarde"), tr(melding)))
        return model

    def _vul_tabellen(self, namen: list[str]) -> None:
        self.cmb_tabel.blockSignals(True)
        self.cmb_tabel.clear()
        self.cmb_tabel.addItems(namen)
        self.cmb_tabel.blockSignals(False)
        self.toon_tabel()

    def _laad_access_tabel(self, naam: str) -> TabelModel:
        with closing(verbind_access(self.bron_pad, self.wachtwoord)) as conn:
            schema = lees_schema(conn, naam)
            totaal = tel_rijen(conn, naam)
            kolommen, rijen = lees_rijen(conn, naam, MAX_EDITOR_RIJEN)
        self.access_tabellen[naam] = AccessTabel(schema, kolommen, rijen, totaal)
        per_kolom = [schema.kolom(k) for k in kolommen]
        alleen_lezen = {i for i, k in enumerate(per_kolom) if k is None or not k.ondersteund}

        def omzetter(kolom: int, tekst: str) -> Any:
            definitie = per_kolom[kolom]
            return zet_om(tekst, definitie) if definitie else tekst

        model = TabelModel(kolommen, [list(r) for r in rijen], bool(schema.sleutel), alleen_lezen, omzetter)
        return self._nieuw_model(model)

    def toon_tabel(self) -> None:
        naam = self.cmb_tabel.currentText()
        self._zet_opmerking("")
        if not naam:
            self.proxy.setSourceModel(None)
            self.werk_knoppen_bij()
            return
        if naam not in self.modellen and self.bron_soort == "access":
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            try:
                self.modellen[naam] = self._laad_access_tabel(naam)
            except Exception as e:
                QApplication.restoreOverrideCursor()
                self.log(f"❌ Fout bij laden van tabel '{naam}': {str(e).splitlines()[0]}")
                self.proxy.setSourceModel(None)
                self.werk_knoppen_bij()
                return
            QApplication.restoreOverrideCursor()
        model = self.modellen.get(naam)
        self.proxy.setSourceModel(model)
        if self.bron_soort == "access" and naam in self.access_tabellen:
            info = self.access_tabellen[naam]
            opmerkingen = []
            if not info.schema.sleutel:
                opmerkingen.append("Deze tabel heeft geen primaire sleutel en kan alleen bekeken worden.")
            elif info.totaal > len(info.origineel):
                # Veilig: niet-geladen rijen staan niet in het plan en worden dus nooit verwijderd
                opmerkingen.append(f"Alleen de eerste {len(info.origineel)} van {info.totaal} rijen zijn geladen.")
            self._zet_opmerking(opmerkingen[0] if opmerkingen else "")
        self.view.resizeColumnsToContents()
        self.werk_knoppen_bij()

    # --- Bewerken ---
    def zoek(self, tekst: str) -> None:
        self.proxy.setFilterFixedString(tekst)

    def rij_toevoegen(self) -> None:
        model = self.model
        if model is None or not model.bewerkbaar:
            return
        self.txt_zoek.clear()
        rij = model.voeg_rij_toe()
        index = self.proxy.mapFromSource(model.index(rij, 0))
        self.view.scrollTo(index)
        self.view.setCurrentIndex(index)

    def rijen_verwijderen(self) -> None:
        model = self.model
        selectie = self.view.selectionModel()
        if model is None or not model.bewerkbaar or selectie is None:
            return
        rijen = {self.proxy.mapToSource(i).row() for i in selectie.selectedIndexes()}
        model.verwijder_rijen(sorted(rijen))

    # --- Opslaan ---
    def opslaan_als(self) -> bool:
        if not self.modellen:
            return False
        if self.bron_soort == "access":
            return self._opslaan_access()
        return self._opslaan_bestand()

    def _opslaan_bestand(self) -> bool:
        basis, extensie = os.path.splitext(self.bron_pad)
        voorstel = f"{basis}_edited{extensie}"
        filters = tr("Excel-bestand (*.xlsx);;CSV-bestand (*.csv)")
        gekozen_filter = filters.split(";;")[1] if extensie.lower() == ".csv" else filters.split(";;")[0]
        pad, _ = QFileDialog.getSaveFileName(self, tr("Opslaan als"), voorstel, filters, gekozen_filter)
        if not pad:
            return False
        try:
            if pad.lower().endswith(".csv"):
                model = self.model
                if model is None:
                    return False
                schrijf_csv(pad, (model.kolommen, model.rijen))
                if len(self.modellen) > 1:
                    self.log(f"ℹ Alleen blad '{self.cmb_tabel.currentText()}' is als CSV opgeslagen.")
            else:
                schrijf_excel(pad, {naam: (m.kolommen, m.rijen) for naam, m in self.modellen.items()})
        except Exception as e:
            QMessageBox.critical(self, tr("Fout"), tr(f"Opslaan mislukt: {str(e).splitlines()[0]}"))
            return False
        for model in self.modellen.values():
            model.gewijzigd = False
        self.bron_pad = os.path.normpath(pad)
        language_support.set_text(self.lbl_bron, self.bron_pad)
        self.werk_status_bij()
        self.log(f"✅ Opgeslagen als: {os.path.basename(pad)}")
        return True

    def _opslaan_access(self) -> bool:
        gewijzigd = {naam: m for naam, m in self.modellen.items() if m.gewijzigd}
        if not gewijzigd:
            QMessageBox.information(self, tr("Info"), tr("Er zijn geen wijzigingen om op te slaan."))
            return True
        plannen = []
        fouten: list[str] = []
        for naam, model in gewijzigd.items():
            info = self.access_tabellen[naam]
            plan = maak_plan(info.schema, (model.kolommen, model.rijen), (info.kolommen, info.origineel),
                             MODUS_BIJWERKEN, verwijderen=True, eerste_rijnummer=1)
            fouten += [f"{naam}: {tr(f)}" for f in plan.fouten]
            plannen.append(plan)
        if fouten:
            QMessageBox.warning(self, tr("Kan niet opslaan"), "\n".join(fouten[:15]))
            return False

        pad, _ = QFileDialog.getSaveFileName(
            self, tr("Opslaan als nieuwe kopie van de database"), voorgestelde_kopienaam(self.bron_pad, "edited"),
            tr("Access Databases (*.accdb *.mdb)"))
        if not pad:
            return False
        pad = os.path.normpath(pad)
        if is_vergrendeld(self.bron_pad) and QMessageBox.question(
                self, tr("Database in gebruik"),
                tr("De database lijkt geopend te zijn in Access. Sluit Access voor een betrouwbare kopie. Toch doorgaan?")
        ) != QMessageBox.StandardButton.Yes:
            return False

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            zet_terug_in_kopie(self.bron_pad, pad, plannen, self.wachtwoord)
        except Exception as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, tr("Fout"), tr(f"Opslaan mislukt: {str(e).splitlines()[0]}"))
            return False
        QApplication.restoreOverrideCursor()
        aantal = sum(p.aantal_bewerkingen for p in plannen)
        self.log(f"✅ {aantal} wijzigingen opgeslagen in nieuwe kopie: {os.path.basename(pad)}")
        huidige_tabel = self.cmb_tabel.currentText()
        for model in self.modellen.values():
            model.gewijzigd = False
        # Verder werken in de nieuwe kopie
        self.open_access(pad, self.wachtwoord)
        index = self.cmb_tabel.findText(huidige_tabel)
        if index >= 0:
            self.cmb_tabel.setCurrentIndex(index)
        return True
