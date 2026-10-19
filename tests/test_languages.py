import tempfile
import unittest
from unittest.mock import patch

from PyQt6.QtCore import QSettings
from PyQt6.QtWidgets import QApplication

import access_exporter_pyqt6 as exporter
from translations import TRANSLATIONS, language_support, tr


class LanguageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.settings = QSettings(
            self.directory.name + "/settings.ini", QSettings.Format.IniFormat
        )
        settings_patch = patch.object(exporter, "QSettings", return_value=self.settings)
        settings_patch.start()
        self.addCleanup(settings_patch.stop)
        environment_patch = patch.object(exporter.AccessExporterGUI, "controleer_omgeving")
        environment_patch.start()
        self.addCleanup(environment_patch.stop)
        language_support.set_language("nl")
        self.addCleanup(language_support.set_language, "nl")
        self.window = exporter.AccessExporterGUI()
        self.addCleanup(self.window.close)

    def switch(self, code):
        self.window.language_combo.setCurrentIndex(self.window.language_combo.findData(code))

    def add_job(self):
        self.window.access_pad = "/tmp/voorbeeld.accdb"
        self.window.zet_doelmap("/tmp/export")
        self.window.huidige_db_tabellen = ["Klanten", "Orders"]
        self.window.geselecteerde_tabellen = ["Klanten"]
        self.window.voeg_toe()

    def test_default_and_live_switching(self):
        self.assertEqual(self.window.language_combo.currentData(), "nl")
        self.add_job()
        for code, status, label in (
            ("en", "Queued", "Language"),
            ("fr", "En attente", "Langue"),
            ("nl", "In de wachtrij", "Taal"),
        ):
            with self.subTest(language=code):
                self.switch(code)
                self.assertEqual(self.window.table_queue.item(0, 3).text(), status)
                self.assertEqual(self.window.status_van_rij(0), exporter.STATUS_WACHTRIJ)
                self.assertEqual(tr("Taal"), label)
                self.assertEqual(self.window.lbl_doelmap_pad.text(), "/tmp/export")
                self.assertEqual(self.window.jobs[0]["tabellen"], ["Klanten"])

    def test_converting_translated_queue_and_live_progress(self):
        self.add_job()
        self.switch("en")
        with patch.object(exporter.ConversionWorker, "start") as start:
            self.window.start_conversie(True)
        start.assert_called_once()
        self.assertEqual(self.window.worker.taken[0][1]["tabellen"], ["Klanten"])
        self.window.update_table_row(0, "25%", exporter.STATUS_BEZIG, "Tabel 1/2: Klanten")
        self.switch("fr")
        self.assertEqual(self.window.status_van_rij(0), exporter.STATUS_BEZIG)
        self.assertEqual(self.window.table_queue.item(0, 3).text(), "En cours")
        self.assertEqual(self.window.table_queue.item(0, 4).text(), "Table 1/2 : Klanten")
        self.window.update_table_row(0, "100%", exporter.STATUS_KLAAR, "Succesvol (10 rijen)")
        self.assertEqual(self.window.table_queue.item(0, 4).text(), "Réussi (10 lignes)")

    def test_dialog_and_logs(self):
        self.add_job()
        self.switch("fr")
        with patch.object(exporter.TableSelectionDialog, "update_preview"):
            dialog = exporter.TableSelectionDialog("/tmp/voorbeeld.accdb", ["Klanten"], ["Klanten"], self.window)
        self.addCleanup(dialog.close)
        self.assertEqual(dialog.windowTitle(), "Sélectionner et prévisualiser les tables")
        self.assertEqual(dialog.lbl_teller.text(), "1 tables sur 1 sélectionnées")
        self.assertEqual(dialog.list_widget.item(0).text(), "Klanten")
        dialog.zet_alle_vinkjes(False)
        self.assertEqual(dialog.lbl_teller.text(), "0 tables sur 1 sélectionnées")
        self.assertIn("Ajouté à la file : voorbeeld.accdb au format XLSX", self.window.txt_log.toPlainText())
        self.switch("nl")
        self.assertIn("Toegevoegd aan wachtrij: voorbeeld.accdb als XLSX", self.window.txt_log.toPlainText())

    def test_saved_preference_and_invalid_language_fallback(self):
        self.switch("fr")
        restored = exporter.AccessExporterGUI()
        self.addCleanup(restored.close)
        self.assertEqual(restored.language_combo.currentData(), "fr")
        self.assertIn("vers Excel/CSV", restored.windowTitle())
        self.settings.setValue("taal", "unsupported")
        fallback = exporter.AccessExporterGUI()
        self.addCleanup(fallback.close)
        self.assertEqual(fallback.language_combo.currentData(), "nl")

    def test_catalog_and_formatted_messages(self):
        for code, index in (("en", 0), ("fr", 1)):
            language_support.set_language(code)
            for source, targets in TRANSLATIONS.items():
                with self.subTest(language=code, source=source):
                    self.assertEqual(tr(source), targets[index])
            self.assertNotIn("{", tr("Tabel 1/2: Klanten"))
            self.assertEqual(tr("Klanten"), "Klanten")
        self.assertEqual(tr("Doelbestand is in gebruik of alleen-lezen: test.xlsx"),
                         "Le fichier de destination est utilisé ou en lecture seule : test.xlsx")


if __name__ == "__main__":
    unittest.main()