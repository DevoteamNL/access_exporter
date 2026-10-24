"""Dialogen: tabellen selecteren met voorbeeld, en wachtwoord vragen voor beveiligde databases."""

import os
from collections.abc import Callable, Iterator
from contextlib import closing
from typing import TypeVar

import pyodbc
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (
    QAbstractItemView, QApplication, QDialog, QDialogButtonBox, QHBoxLayout, QInputDialog, QLabel, QLineEdit,
    QListWidget, QListWidgetItem, QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from .core import PREVIEW_RIJEN, WachtwoordNodig, quote_naam, verbind_access
from .translations import language_support, tr

T = TypeVar("T")


def met_wachtwoord(parent: QWidget | None, pad: str, actie: Callable[[str], T],
                   wachtwoord: str = "") -> tuple[T, str] | None:
    """Voert `actie(wachtwoord)` uit en vraagt om een (nieuw) wachtwoord zolang de database dat vereist.
    Geeft (resultaat, wachtwoord) terug, of None als de gebruiker annuleert. Het wachtwoord wordt nooit bewaard."""
    eerste = not wachtwoord
    while True:
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            return actie(wachtwoord), wachtwoord
        except WachtwoordNodig:
            pass
        finally:
            QApplication.restoreOverrideCursor()
        naam = os.path.basename(pad)
        vraag = (f"Voer het wachtwoord in voor {naam}:" if eerste
                 else f"Onjuist wachtwoord. Voer het wachtwoord in voor {naam}:")
        eerste = False
        tekst, ok = QInputDialog.getText(parent, tr("Wachtwoord vereist"), tr(vraag), QLineEdit.EchoMode.Password)
        if not ok:
            return None
        wachtwoord = tekst


class TableSelectionDialog(QDialog):
    def __init__(self, db_path: str, tabellen: list[str], geselecteerde_tabellen: list[str],
                 parent: QWidget | None = None, wachtwoord: str = ""):
        super().__init__(parent)
        self.setWindowTitle("Tabellen Selecteren en Bekijken")
        self.resize(1000, 550)
        self.db_path = db_path
        self.wachtwoord = wachtwoord
        self.alle_tabellen = tabellen
        self.geselecteerde = list(geselecteerde_tabellen)
        self._conn: pyodbc.Connection | None = None  # Eén verbinding voor alle previews

        hoofd_layout = QVBoxLayout(self)
        layout = QHBoxLayout()
        hoofd_layout.addLayout(layout)

        # Links: lijst met vinkjes
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

        # Rechts: voorbeeld
        right_layout = QVBoxLayout()
        right_layout.addWidget(self._vet_label(f"2. Datavoorbeeld (Eerste {PREVIEW_RIJEN} rijen):"))
        self.preview_table = QTableWidget()
        self.preview_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.preview_table.setAlternatingRowColors(True)
        right_layout.addWidget(self.preview_table)
        layout.addLayout(right_layout, 3)

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
                self._conn = verbind_access(self.db_path, self.wachtwoord)
            with closing(self._conn.cursor()) as cursor:
                cursor.execute(f"SELECT TOP {PREVIEW_RIJEN} * FROM {quote_naam(huidig_item.text())}")
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
        self.geselecteerde = [item.text() for item in self._items() if item.checkState() == Qt.CheckState.Checked]
        self.accept()
