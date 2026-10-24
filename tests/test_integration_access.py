"""Integratietests met een echte (tijdelijke) Access-database: exporteren, terugzetten en bewerken.
Integration tests with a real (temporary) Access database: export, merge back and edit."""

import hashlib
from contextlib import closing
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

import pytest

from access_exporter import editor_tab as editor_module
from access_exporter.core import (
    WachtwoordNodig, haal_tabellen_op, lees_rijen, verbind_access, zoek_manifest,
)
from access_exporter.editor_tab import EditorTab
from access_exporter.file_io import lees_export, schrijf_csv, schrijf_excel
from access_exporter.merge import (
    MODUS_BIJWERKEN, MODUS_TOEVOEGEN, MODUS_VERVANGEN, bereken_plannen, bronnen_uit_bestand, lees_schema, raad_tabel,
    zet_terug_in_kopie,
)
from access_exporter.worker import ConversionWorker
from tests.conftest import vereist_driver

pytestmark = vereist_driver


def hash_van(pad: Path) -> str:
    return hashlib.sha256(pad.read_bytes()).hexdigest()


def exporteer(database: Path, doelmap: Path, export_type: str, tabellen: list[str], decimaal: str = ",") -> list[tuple]:
    job = {"access_pad": str(database), "access_naam": database.name, "naam_zonder_ext": database.stem,
           "doelmap": str(doelmap), "export_type": export_type, "tabellen": tabellen, "csv_decimaal": decimaal}
    worker = ConversionWorker([(0, job)], open_map=False)
    updates: list[tuple] = []
    worker.update_table_signal.connect(lambda *a: updates.append(a))
    worker.run()
    return updates


def inhoud(pad: Path, tabel: str) -> list[tuple]:
    with closing(verbind_access(str(pad))) as conn:
        return sorted(lees_rijen(conn, tabel)[1], key=lambda r: tuple(str(x) for x in r))


def test_tabellen_queries_en_sleutels(database: Path):
    assert haal_tabellen_op(str(database)) == ["Klanten", "Log", "Order_Regels"]
    assert haal_tabellen_op(str(database), met_queries=True)[-1] == "DureKlanten"
    with closing(verbind_access(str(database))) as conn:
        assert lees_schema(conn, "Klanten").sleutel == ["ID"]
        assert lees_schema(conn, "Order_Regels").sleutel == ["OrderID", "Regel"]
        assert lees_schema(conn, "Log").sleutel == []


def test_wachtwoordfout_wordt_herkend(database: Path):
    import pyodbc
    with patch("access_exporter.core.pyodbc.connect", side_effect=pyodbc.Error("HY000", "Not a valid password. (-1905)")):
        with pytest.raises(WachtwoordNodig):
            verbind_access(str(database), "fout")


def test_export_excel_bewerken_en_terugzetten(database: Path, tmp_path: Path, qapp):
    origineel = hash_van(database)
    updates = exporteer(database, tmp_path, "xlsx", ["Klanten", "Order_Regels", "Log"])
    assert updates[-1][2] == "Klaar", updates[-1]
    excel = tmp_path / "Voorbeeld.xlsx"
    manifest = zoek_manifest(str(excel))
    assert manifest and [t["table"] for t in manifest["tables"]] == ["Klanten", "Order_Regels", "Log"]

    # Bewerk het bestand zoals een gebruiker in Excel zou doen
    bladen = lees_export(str(excel))
    kolommen, rijen = bladen["Klanten"]
    i_naam, i_prijs, i_id = kolommen.index("Naam"), kolommen.index("Prijs"), kolommen.index("ID")
    rijen[0][i_naam] = "Jansen BV"
    rijen[1][i_prijs] = "1.299,95"
    rijen = [r for r in rijen if r[i_naam] != "Müller; \"Ltd\""]
    nieuw = [None] * len(kolommen)
    nieuw[i_naam], nieuw[kolommen.index("Datum")], nieuw[kolommen.index("Actief")] = "Nieuwe klant", "15-03-2025", "ja"
    rijen.append(nieuw)
    bladen["Klanten"] = (kolommen, rijen)
    bewerkt = tmp_path / "Voorbeeld_bewerkt.xlsx"
    schrijf_excel(str(bewerkt), bladen)
    (tmp_path / "Voorbeeld_bewerkt.xlsx.export.json").write_text((tmp_path / "Voorbeeld.xlsx.export.json").read_text("utf-8"), "utf-8")

    bronnen = bronnen_uit_bestand(str(bewerkt))
    tabellen = haal_tabellen_op(str(database))
    koppelingen = {}
    for bron in bronnen:
        tabel = raad_tabel(bron, tabellen)
        assert tabel == bron.naam
        koppelingen.setdefault(tabel, []).append(bron)

    with closing(verbind_access(str(database))) as conn:
        plannen = bereken_plannen(conn, koppelingen, MODUS_BIJWERKEN, verwijderen=True)
    per_tabel = {p.tabel: p for p in plannen}
    klanten = per_tabel["Klanten"]
    assert not klanten.fouten, klanten.fouten
    assert (len(klanten.invoegen), len(klanten.wijzigen), len(klanten.verwijderen)) == (1, 2, 1)
    assert per_tabel["Order_Regels"].aantal_bewerkingen == 0
    assert per_tabel["Log"].geblokkeerd  # Geen primaire sleutel in modus bijwerken

    kopie = tmp_path / "Voorbeeld_merged.accdb"
    zet_terug_in_kopie(str(database), str(kopie), plannen)
    assert hash_van(database) == origineel
    with closing(verbind_access(str(kopie))) as conn:
        _, rijen_kopie = lees_rijen(conn, "Klanten")
    per_naam = {r[1]: r for r in rijen_kopie}
    assert set(per_naam) == {"Jansen BV", "de Vries", "Nieuwe klant"}
    assert per_naam["de Vries"][2] == Decimal("1299.95")
    assert per_naam["Nieuwe klant"][3] == datetime(2025, 3, 15) and per_naam["Nieuwe klant"][4] is True
    assert per_naam["Nieuwe klant"][0] > 3  # Nieuw autonummer


def test_csv_export_en_vervangen_of_toevoegen(database: Path, tmp_path: Path, qapp):
    exporteer(database, tmp_path, "csv", ["Klanten", "Order_Regels", "Log"], decimaal=",")
    csv_klanten = tmp_path / "Voorbeeld_Klanten.csv"
    tekst = csv_klanten.read_text("utf-8-sig")
    assert "1234,5" in tekst and "Regel 1\nRegel 2" in tekst.replace("\r\n", "\n")
    tabellen = haal_tabellen_op(str(database))
    bron_log = bronnen_uit_bestand(str(tmp_path / "Voorbeeld_Log.csv"))[0]
    assert raad_tabel(bron_log, tabellen) == "Log"

    # Log: rijen toevoegen (geen sleutel nodig); Order_Regels: inhoud vervangen
    bron_log.rijen.append(["tweede regel"])
    bron_regels = bronnen_uit_bestand(str(tmp_path / "Voorbeeld_Order_Regels.csv"))[0]
    assert raad_tabel(bron_regels, tabellen) == "Order_Regels"
    bron_regels.rijen = [r for r in bron_regels.rijen if r[2] != "Bel"]
    with closing(verbind_access(str(database))) as conn:
        plannen = bereken_plannen(conn, {"Log": [bron_log]}, MODUS_TOEVOEGEN, False)
        plannen += bereken_plannen(conn, {"Order_Regels": [bron_regels]}, MODUS_VERVANGEN, False)
    kopie = tmp_path / "kopie.accdb"
    zet_terug_in_kopie(str(database), str(kopie), plannen)
    assert [r[0] for r in inhoud(kopie, "Log")] == ["gestart", "gestart", "tweede regel"]
    assert [r[2] for r in inhoud(kopie, "Order_Regels")] == ["Fiets", "Slot"]


def test_fout_tijdens_uitvoeren_verwijdert_kopie(database: Path, tmp_path: Path):
    bron = (["OrderID", "Regel", "Product"], [[9, 1, "x" * 40]])  # Te lang voor TEXT(30) -> omzetfout
    from access_exporter.merge import maak_plan
    with closing(verbind_access(str(database))) as conn:
        schema = lees_schema(conn, "Order_Regels")
        plan = maak_plan(schema, bron, lees_rijen(conn, "Order_Regels"), MODUS_BIJWERKEN)
    assert plan.fouten and not plan.invoegen
    # Forceer een databasefout: dubbele sleutel bij invoegen
    plan.invoegen = [{"OrderID": 1, "Regel": 1, "Product": "dubbel"}]
    kopie = tmp_path / "mislukt.accdb"
    with pytest.raises(Exception):
        zet_terug_in_kopie(str(database), str(kopie), [plan])
    assert not kopie.exists()


def test_editor_bewerkt_access_en_slaat_op_als_kopie(database: Path, tmp_path: Path, qapp, settings):
    origineel = hash_van(database)
    tab = EditorTab(lambda _m: None, settings)
    tab.open_pad(str(database))
    assert tab.cmb_tabel.currentText() == "Klanten"
    model = tab.model
    assert model is not None and model.bewerkbaar
    kolom_naam = model.kolommen.index("Naam")
    assert model.setData(model.index(0, kolom_naam), "Jansen & Zn")
    with patch.object(editor_module.QMessageBox, "warning") as ongeldig:
        assert not model.setData(model.index(0, model.kolommen.index("Aantal")), "veel")  # Ongeldig getal
    ongeldig.assert_called_once()
    rij = model.voeg_rij_toe()
    model.setData(model.index(rij, kolom_naam), "Toegevoegd")
    tab.cmb_tabel.setCurrentText("Log")
    assert tab.model is not None and not tab.model.bewerkbaar  # Geen sleutel: alleen-lezen

    kopie = tmp_path / "bewerkt.accdb"
    with patch.object(editor_module.QFileDialog, "getSaveFileName", return_value=(str(kopie), "")), \
         patch.object(editor_module.QMessageBox, "warning") as waarschuwing:
        assert tab.opslaan_als()
    waarschuwing.assert_not_called()
    assert hash_van(database) == origineel
    namen = {r[1] for r in inhoud(kopie, "Klanten")}
    assert {"Jansen & Zn", "Toegevoegd", "de Vries"} <= namen and "Jansen" not in namen
    assert tab.bron_pad == str(kopie) and not tab.heeft_wijzigingen()


def test_editor_bestand_opslaan_als(tmp_path: Path, qapp, settings):
    bron = tmp_path / "data.csv"
    schrijf_csv(str(bron), (["A", "B"], [["1", "x"], ["2", None]]))
    tab = EditorTab(lambda _m: None, settings)
    tab.open_pad(str(bron))
    model = tab.model
    assert model is not None
    model.setData(model.index(1, 1), "y")
    model.verwijder_rijen([0])
    doel = tmp_path / "data_bewerkt.xlsx"
    with patch.object(editor_module.QFileDialog, "getSaveFileName", return_value=(str(doel), "")):
        assert tab.opslaan_als()
    assert lees_export(str(doel)) == {"data": (["A", "B"], [["2", "y"]])}
    assert bron.read_text("utf-8-sig").splitlines() == ["A;B", "1;x", "2;"]
