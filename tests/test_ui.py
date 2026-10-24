"""UI-tests (zonder scherm): hoofdvenster met tabbladen en thema, en het tabblad Terugzetten met een echte database.
UI tests (headless): main window with tabs and theme, and the Merge tab against a real database."""

from pathlib import Path
from unittest.mock import patch

import pytest

from access_exporter import main_window, merge_tab as merge_module, themes
from access_exporter.core import STATUS_KLAAR, STATUS_WACHTRIJ
from access_exporter.file_io import lees_export, schrijf_excel
from access_exporter.merge import MODUS_VERVANGEN
from access_exporter.merge_tab import MergeTab
from access_exporter.translations import language_support
from tests.conftest import vereist_driver, wacht_tot
from tests.test_integration_access import exporteer, inhoud


@pytest.fixture
def venster(qapp, settings):
    with patch.object(main_window, "QSettings", return_value=settings), \
         patch.object(main_window.AccessExporterGUI, "controleer_omgeving"):
        window = main_window.AccessExporterGUI()
    yield window
    window.close()
    language_support.set_language("nl")


def test_tabbladen_en_thema_wisselen(venster, settings):
    assert [venster.tabs.tabText(i) for i in range(3)] == ["📤 Exporteren", "✏️ Bewerken", "🔁 Terugzetten naar Access"]
    assert venster.theme_combo.currentData() == themes.THEMA_SYSTEEM  # Standaard: systeem
    venster.language_combo.setCurrentIndex(venster.language_combo.findData("en"))
    assert venster.tabs.tabText(2) == "🔁 Merge back into Access"
    assert venster.theme_combo.itemText(2) == "Dark"
    assert venster.editor_tab.btn_opslaan.text() == "💾 Save as..."
    assert venster.merge_tab.modus_knoppen[MODUS_VERVANGEN].text() == "Replace table contents"
    venster.theme_combo.setCurrentIndex(venster.theme_combo.findData(themes.THEMA_DONKER))
    assert settings.value("thema") == themes.THEMA_DONKER and themes.actief() == "dark"
    venster.theme_combo.setCurrentIndex(venster.theme_combo.findData(themes.THEMA_LICHT))
    assert themes.actief() == "light"


def test_wis_afgeronde_taken(venster):
    for csv in (False, True):
        venster.access_pad = "C:/tmp/voorbeeld.accdb"
        venster.zet_doelmap("C:/tmp")
        venster.huidige_db_tabellen, venster.geselecteerde_tabellen = ["Klanten"], ["Klanten"]
        venster.rb_csv.setChecked(csv)
        venster.voeg_toe()
    assert len(venster.jobs) == 2
    venster.update_table_row(0, "100%", STATUS_KLAAR, "Succesvol (10 rijen)")
    venster.wis_afgeronde_taken()
    assert len(venster.jobs) == 1 and venster.status_van_rij(0) == STATUS_WACHTRIJ


@vereist_driver
def test_terugzetten_via_tabblad(database: Path, tmp_path: Path, qapp, settings):
    exporteer(database, tmp_path, "xlsx", ["Klanten", "Log"])
    excel = tmp_path / "Voorbeeld.xlsx"
    bladen = lees_export(str(excel))
    kolommen, rijen = bladen["Klanten"]
    rijen[0][kolommen.index("Naam")] = "Gewijzigd"
    schrijf_excel(str(excel), bladen)

    tab = MergeTab(lambda _m: None, settings)
    tab.open_paden([str(database), str(excel)])
    assert tab.tbl_koppeling.rowCount() == 2
    assert [c.currentText() for c in tab._combos()] == ["Klanten", "Log"]
    tab._combos()[1].setCurrentIndex(0)  # Log overslaan (geen sleutel)
    tab.bereken_voorbeeld()
    wacht_tot(lambda: not tab.bezig() and bool(tab.plannen))
    assert tab.tbl_voorbeeld.item(0, 2).text() == "1"  # 1 wijziging
    assert tab.btn_samenvoegen.isEnabled()

    doel = tmp_path / "samengevoegd.accdb"
    with patch.object(merge_module.QFileDialog, "getSaveFileName", return_value=(str(doel), "")), \
         patch.object(merge_module.QMessageBox, "question", return_value=merge_module.QMessageBox.StandardButton.Yes), \
         patch.object(merge_module.QMessageBox, "information") as klaar:
        tab.samenvoegen()
        wacht_tot(lambda: klaar.called or not tab.bezig() and tab.worker is None)
        wacht_tot(lambda: klaar.called)
    assert "Gewijzigd" in [r[1] for r in inhoud(doel, "Klanten")]
    assert "Gewijzigd" not in [r[1] for r in inhoud(database, "Klanten")]
