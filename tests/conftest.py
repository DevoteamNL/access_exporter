"""Gedeelde testhulpmiddelen: Qt zonder scherm, tijdelijke instellingen en een tijdelijke Access-database.
Shared test helpers: headless Qt, temporary settings and a temporary Access database."""

import ctypes
import os
import sys
import time
from collections.abc import Callable
from contextlib import closing
from datetime import datetime
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pyodbc  # noqa: E402
import pytest  # noqa: E402
from PyQt6.QtCore import QSettings  # noqa: E402
from PyQt6.QtWidgets import QApplication  # noqa: E402

from access_exporter.core import ACCESS_DRIVER, access_driver_beschikbaar  # noqa: E402

heeft_driver = sys.platform == "win32" and access_driver_beschikbaar()
vereist_driver = pytest.mark.skipif(not heeft_driver, reason="Microsoft Access ODBC-driver niet beschikbaar")


@pytest.fixture(scope="session")
def qapp() -> QApplication:
    app = QApplication.instance()
    return app if isinstance(app, QApplication) else QApplication([])


@pytest.fixture
def settings(tmp_path: Path) -> QSettings:
    return QSettings(str(tmp_path / "settings.ini"), QSettings.Format.IniFormat)


def wacht_tot(voorwaarde: Callable[[], bool], timeout: float = 30) -> None:
    einde = time.monotonic() + timeout
    while not voorwaarde():
        if time.monotonic() > einde:
            raise TimeoutError("Voorwaarde niet bereikt binnen de tijd")
        QApplication.processEvents()
        time.sleep(0.01)


def maak_database(pad: Path) -> None:
    attributen = f'CREATE_DBV12="{pad}" General\0\0'
    buffer = ctypes.create_unicode_buffer(attributen, len(attributen))
    if not ctypes.windll.ODBCCP32.SQLConfigDataSourceW(None, 1, ACCESS_DRIVER, buffer):  # type: ignore[attr-defined]
        raise RuntimeError("Kon geen Access-database aanmaken")


@pytest.fixture
def database(tmp_path: Path) -> Path:
    """Database met Klanten (autonummer-sleutel), Order_Regels (samengestelde sleutel), Log (geen sleutel)
    en een opgeslagen query."""
    if not heeft_driver:
        pytest.skip("Microsoft Access ODBC-driver niet beschikbaar")
    pad = tmp_path / "Voorbeeld.accdb"
    maak_database(pad)
    with closing(pyodbc.connect(f"DRIVER={{{ACCESS_DRIVER}}};DBQ={pad};", autocommit=True)) as conn:
        c = conn.cursor()
        c.execute("CREATE TABLE Klanten (ID COUNTER PRIMARY KEY, Naam TEXT(50), Prijs CURRENCY, Datum DATETIME, "
                  "Actief BIT, Aantal LONG, Factor DOUBLE, Notitie MEMO)")
        c.execute("CREATE TABLE Order_Regels (OrderID LONG NOT NULL, Regel LONG NOT NULL, Product TEXT(30), "
                  "CONSTRAINT PrimaryKey PRIMARY KEY (OrderID, Regel))")
        c.execute("CREATE TABLE Log (Tekst TEXT(100))")
        rijen = [
            ("Jansen", 12.5, datetime(2024, 1, 31), True, 3, 0.25, "Eerste klant"),
            ("de Vries", 1234.5, datetime(2023, 12, 1, 14, 30), False, 0, 1.5, None),
            ("Müller; \"Ltd\"", None, None, True, None, None, "Regel 1\nRegel 2"),
        ]
        c.executemany("INSERT INTO Klanten (Naam, Prijs, Datum, Actief, Aantal, Factor, Notitie) VALUES (?, ?, ?, ?, ?, ?, ?)", rijen)
        c.executemany("INSERT INTO Order_Regels VALUES (?, ?, ?)", [(1, 1, "Fiets"), (1, 2, "Bel"), (2, 1, "Slot")])
        c.execute("INSERT INTO Log VALUES ('gestart')")
        c.execute("CREATE VIEW DureKlanten AS SELECT Naam, Prijs FROM Klanten WHERE Prijs > 100")
    return pad
