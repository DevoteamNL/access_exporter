"""Unittests zonder Access-driver: hulpfuncties, bestanden, samenvoeglogica, vertalingen en thema's.
Unit tests without the Access driver: helpers, files, merge logic, translations and themes."""

import json
import re
from datetime import datetime
from decimal import Decimal
from pathlib import Path

import pytest

from access_exporter import themes
from access_exporter.core import (
    is_wachtwoordfout, lockbestand, maak_bladnaam, maak_csv_naam, maak_kopie, manifest_pad, schrijf_manifest,
    voorgestelde_kopienaam, zoek_manifest,
)
from access_exporter.file_io import lees_export, schrijf_csv, schrijf_excel
from access_exporter.merge import (
    MODUS_BIJWERKEN, MODUS_TOEVOEGEN, MODUS_VERVANGEN, Bron, Kolom, OmzetFout, Schema, gelijk, kies_sleutel, maak_plan,
    raad_tabel, voeg_bronnen_samen, zet_om,
)
from access_exporter.translations import language_support, tr


# --- core ---
def test_kopienaam_en_lockbestand():
    naam = voorgestelde_kopienaam(r"C:\map\Klanten.accdb", "merged")
    assert re.fullmatch(r"C:\\map\\Klanten_merged_\d{8}-\d{6}\.accdb", naam)
    assert lockbestand(r"C:\a\b.accdb") == r"C:\a\b.laccdb"
    assert lockbestand(r"C:\a\b.MDB") == r"C:\a\b.ldb"


def test_maak_kopie_weigert_origineel(tmp_path: Path):
    bron = tmp_path / "db.accdb"
    bron.write_bytes(b"x")
    with pytest.raises(ValueError):
        maak_kopie(str(bron), str(tmp_path / "." / "DB.accdb"))
    maak_kopie(str(bron), str(tmp_path / "kopie.accdb"))
    assert (tmp_path / "kopie.accdb").read_bytes() == b"x"


def test_wachtwoordfout_herkenning():
    assert is_wachtwoordfout(Exception("[Microsoft][ODBC] Not a valid password. (-1905)"))
    assert not is_wachtwoordfout(Exception("Could not find file"))


def test_bladnamen_en_csv_namen():
    gebruikt: set[str] = set()
    assert maak_bladnaam("a/b:c", gebruikt) == "a_b_c"
    assert maak_bladnaam("A/B:C", gebruikt) == "A_B_C~2"
    namen: set[str] = set()
    assert maak_csv_naam("db", "Order Regels?", namen) == "db_Order Regels.csv"


def test_manifest_excel_en_csv(tmp_path: Path):
    xlsx = tmp_path / "Export.xlsx"
    schrijf_manifest(manifest_pad(str(tmp_path), "Export", "xlsx"), "db.accdb", "xlsx", [{"table": "T", "sheets": ["T"]}])
    assert zoek_manifest(str(xlsx))["tables"][0]["table"] == "T"  # type: ignore[index]
    schrijf_manifest(manifest_pad(str(tmp_path), "Export", "csv"), "db.accdb", "csv", [{"table": "T 1", "file": "Export_T 1.csv"}])
    assert zoek_manifest(str(tmp_path / "Export_T 1.csv"))["tables"][0]["table"] == "T 1"  # type: ignore[index]
    assert zoek_manifest(str(tmp_path / "ander.csv")) is None
    (tmp_path / "kapot.xlsx.export.json").write_text("{niet json", "utf-8")
    assert zoek_manifest(str(tmp_path / "kapot.xlsx")) is None


# --- file_io ---
def test_csv_en_excel_rondgang(tmp_path: Path):
    tabel = (["Naam", "Getal", "Datum"], [["=SOM(A1)", 1.5, datetime(2024, 2, 29)], ["http://x", None, None]])
    schrijf_excel(str(tmp_path / "t.xlsx"), {"Blad1": tabel})
    gelezen = lees_export(str(tmp_path / "t.xlsx"))["Blad1"]
    assert gelezen[1][0] == ["=SOM(A1)", 1.5, datetime(2024, 2, 29)]
    assert gelezen[1][1] == ["http://x", None, None]
    schrijf_csv(str(tmp_path / "t.csv"), (["A", "B"], [["0012", "x;y"], [None, "z"]]))
    assert lees_export(str(tmp_path / "t.csv")) == {"t": (["A", "B"], [["0012", "x;y"], [None, "z"]])}


# --- merge: omzetten ---
TEKST = Kolom("Naam", "VARCHAR", True, 5)


@pytest.mark.parametrize(("waarde", "kolom", "verwacht"), [
    ("1.234,50", Kolom("P", "CURRENCY"), Decimal("1234.50")),
    ("1,234.50", Kolom("P", "CURRENCY"), Decimal("1234.50")),
    (2.5, Kolom("P", "DOUBLE"), 2.5),
    ("7", Kolom("N", "INTEGER"), 7),
    (7.0, Kolom("N", "INTEGER"), 7),
    ("31-01-2024", Kolom("D", "DATETIME"), datetime(2024, 1, 31)),
    ("2024-01-31 14:30:00", Kolom("D", "DATETIME"), datetime(2024, 1, 31, 14, 30)),
    ("Ja", Kolom("B", "BIT", False), True),
    ("faux", Kolom("B", "BIT", False), False),
    ("", TEKST, None),
    ("abc", TEKST, "abc"),
])
def test_zet_om(waarde, kolom, verwacht):
    assert zet_om(waarde, kolom) == verwacht


@pytest.mark.parametrize(("waarde", "kolom"), [
    ("abc", Kolom("P", "CURRENCY")), ("1,5", Kolom("N", "INTEGER")), ("32-13-2024", Kolom("D", "DATETIME")),
    ("misschien", Kolom("B", "BIT", False)), ("te lange tekst", TEKST),
])
def test_zet_om_fouten(waarde, kolom):
    with pytest.raises(OmzetFout):
        zet_om(waarde, kolom)


def test_gelijk():
    assert gelijk(Decimal("1.5000"), Decimal("1.5"))
    assert gelijk(None, "")
    assert not gelijk("a", "b")


def test_kies_sleutel_primarykey_en_terugval():
    kolommen = [Kolom("ID", "COUNTER", False), Kolom("Code", "VARCHAR", False), Kolom("Naam", "VARCHAR", True)]
    assert kies_sleutel({"PrimaryKey": [(1, "Code")], "Index_1": [(1, "ID")]}, kolommen) == ["Code"]
    assert kies_sleutel({"Index_X": [(1, "ID")], "Index_N": [(1, "Naam")]}, kolommen) == ["ID"]
    assert kies_sleutel({"Index_N": [(1, "Naam")]}, kolommen) == []


# --- merge: plannen ---
SCHEMA = Schema("Klanten", [Kolom("ID", "COUNTER", False), Kolom("Naam", "VARCHAR", True, 50),
                            Kolom("Prijs", "CURRENCY")], ["ID"])
ACCESS = (["ID", "Naam", "Prijs"], [(1, "A", Decimal("1.0000")), (2, "B", Decimal("2.0000")), (3, "C", None)])


def test_plan_bijwerken_toevoegen_verwijderen():
    bron = (["ID", "Naam", "Prijs", "Extra"], [[1, "A", "1,00", "x"], [2, "B2", "2", None], [None, "Nieuw", "5", None]])
    plan = maak_plan(SCHEMA, bron, ACCESS, MODUS_BIJWERKEN, verwijderen=True)
    assert plan.wijzigen == [((2,), {"Naam": "B2"})]
    assert plan.invoegen == [{"Naam": "Nieuw", "Prijs": Decimal("5")}]
    assert plan.verwijderen == [(3,)]
    assert plan.ongewijzigd == 1
    assert any("Extra" in o for o in plan.opmerkingen)


def test_plan_fouten_blokkeren_verwijderen():
    bron = (["ID", "Naam", "Prijs"], [[1, "A", "geen getal"], [2, "B", 2], [2, "B", 2]])
    plan = maak_plan(SCHEMA, bron, ACCESS, MODUS_BIJWERKEN, verwijderen=True)
    assert len(plan.fouten) == 2 and not plan.verwijderen
    assert plan.fouten[0].startswith("Rij 2, kolom 'Prijs'")
    assert plan.fouten[1] == "Rij 4: dubbele sleutelwaarde 2"


def test_plan_modi_zonder_sleutel():
    log = Schema("Log", [Kolom("Tekst", "VARCHAR", True, 100)], [])
    bron = (["Tekst"], [["a"], [None]])
    assert maak_plan(log, bron, (["Tekst"], []), MODUS_BIJWERKEN).geblokkeerd
    assert maak_plan(log, bron, (["Tekst"], []), MODUS_TOEVOEGEN).invoegen == [{"Tekst": "a"}]
    vervangen = maak_plan(log, bron, (["Tekst"], [("oud",)]), MODUS_VERVANGEN)
    assert vervangen.alles_verwijderen and vervangen.invoegen == [{"Tekst": "a"}]


def test_plan_toevoegen_slaat_bestaande_over():
    plan = maak_plan(SCHEMA, (["ID", "Naam"], [[1, "anders"], [9, "nieuw"]]), ACCESS, MODUS_TOEVOEGEN)
    assert plan.ongewijzigd == 1 and plan.invoegen == [{"ID": 9, "Naam": "nieuw"}] and not plan.wijzigen


def test_ontbrekende_sleutelkolom_blokkeert():
    plan = maak_plan(SCHEMA, (["Naam"], [["x"]]), ACCESS, MODUS_BIJWERKEN)
    assert plan.geblokkeerd and "Sleutelkolom 'ID'" in plan.fouten[0]


def test_raad_tabel_en_samenvoegen():
    tabellen = ["Klanten", "Order Regels", "Data[1]"]
    assert raad_tabel(Bron("x", "Klanten_2", [], []), tabellen) == "Klanten"
    assert raad_tabel(Bron("x", "Data_1_", [], []), tabellen) == "Data[1]"
    assert raad_tabel(Bron("x", "Export_Order Regels", [], []), tabellen) == "Order Regels"
    assert raad_tabel(Bron("x", "iets", [], [], tabel_hint="klanten"), tabellen) == "Klanten"
    assert raad_tabel(Bron("x", "Onbekend", [], []), tabellen) is None
    kolommen, rijen = voeg_bronnen_samen([Bron("a", "a", ["A", "B"], [[1, 2]]), Bron("b", "b", ["B", "C"], [[3, 4]])])
    assert kolommen == ["A", "B", "C"] and rijen == [[1, 2, None], [None, 3, 4]]


# --- vertalingen ---
@pytest.fixture
def taal():
    yield language_support.set_language
    language_support.set_language("nl")


def test_getallen_volgens_taal(taal, qapp):
    taal("nl")
    assert tr("Succesvol (1234567 rijen)") == "Succesvol (1.234.567 rijen)"
    taal("en")
    assert tr("Succesvol (1234567 rijen)") == "Successful (1,234,567 rows)"
    taal("fr")
    assert tr("Succesvol (1234567 rijen)") == "Réussi (1\u00a0234\u00a0567 lignes)"


def test_geneste_meldingen_worden_vertaald(taal, qapp):
    taal("en")
    assert tr("⚠ Klanten: Rij 5, kolom 'Prijs': 'abc' is geen getal") == "⚠ Klanten: Row 5, column 'Prijs': 'abc' is not a number"
    assert tr("⚠ Klanten: nog 1500 fouten") == "⚠ Klanten: 1,500 more errors"
    taal("fr")
    assert tr("Totaal 3 wijzigingen klaar om samen te voegen. 1 rijen of tabellen worden overgeslagen vanwege fouten (zie log).").startswith("3 modifications")


def test_vertaalcatalogus_heeft_geldige_plaatshouders():
    from access_exporter.translations import TRANSLATIONS
    for bron, (en, fr) in TRANSLATIONS.items():
        namen = set(re.findall(r"\{(\w+)\}", bron))
        assert set(re.findall(r"\{(\w+)\}", en)) == namen, bron
        assert set(re.findall(r"\{(\w+)\}", fr)) == namen, bron


# --- thema's ---
def test_themas(qapp):
    assert themes.los_op(themes.THEMA_LICHT) == "light"
    assert themes.los_op(themes.THEMA_DONKER) == "dark"
    assert themes.los_op(themes.THEMA_SYSTEEM) in ("light", "dark")
    assert themes.pas_toe(themes.THEMA_LICHT) == "light"
    licht = themes.kleur("tekst")
    assert themes.pas_toe(themes.THEMA_DONKER) == "dark"
    assert themes.kleur("tekst") != licht
    assert qapp.styleSheet() == themes.stylesheet(themes.THEMA_DONKER)


def test_manifest_json_is_leesbaar(tmp_path: Path):
    pad = tmp_path / "m.json"
    schrijf_manifest(str(pad), "C:/db.accdb", "xlsx", [])
    gegevens = json.loads(pad.read_text("utf-8"))
    assert gegevens["app"] == "Access Exporter" and gegevens["format"] == "xlsx"


# --- icoon / icon ---
def test_icoon_bestaat_en_laadt(qapp):
    from PyQt6.QtGui import QIcon

    import access_exporter_pyqt6

    pad = access_exporter_pyqt6.icoon_pad()
    assert pad.exists()
    icoon = QIcon(str(pad))
    assert not icoon.isNull()
    assert {16, 32, 48, 256} <= {s.width() for s in icoon.availableSizes()}
