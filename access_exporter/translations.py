import re

from PyQt6.QtCore import QLibraryInfo, QLocale, QTranslator
from PyQt6.QtWidgets import (
    QAbstractButton, QApplication, QGroupBox, QLabel, QLineEdit, QWidget,
)


TRANSLATIONS = {
    "Taal": ("Language", "Langue"),
    "Access naar Excel/CSV Exporter v{version}": ("Access to Excel/CSV Exporter v{version}", "Exportateur Access vers Excel/CSV v{version}"),
    "Stap 1: Bestanden, tabellen en formaat instellen": ("Step 1: Configure files, tables and format", "Étape 1 : Configurer les fichiers, tables et format"),
    "📂 1. Kies Access-bestand": ("📂 1. Choose Access file", "📂 1. Choisir un fichier Access"),
    "⚙ 2. Kies Exportformaat": ("⚙ 2. Choose export format", "⚙ 2. Choisir le format d'export"),
    "📊 1 Excel-bestand (.xlsx)": ("📊 One Excel file (.xlsx)", "📊 Un fichier Excel (.xlsx)"),
    "📄 Aparte CSV's per tabel (.csv)": ("📄 Separate CSV files per table (.csv)", "📄 Un fichier CSV par table (.csv)"),
    "Geen Access-bestand geselecteerd": ("No Access file selected", "Aucun fichier Access sélectionné"),
    "📋 3. Kies te exporteren tabellen": ("📋 3. Choose tables to export", "📋 3. Choisir les tables à exporter"),
    "Geen tabellen geladen": ("No tables loaded", "Aucune table chargée"),
    "📁 4. Kies Doelmap": ("📁 4. Choose output folder", "📁 4. Choisir le dossier de destination"),
    "➕ 5. Voeg toe aan wachtrij": ("➕ 5. Add to queue", "➕ 5. Ajouter à la file"),
    "Geen map geselecteerd": ("No folder selected", "Aucun dossier sélectionné"),
    "Stap 2: Wachtrij": ("Step 2: Queue", "Étape 2 : File d'attente"),
    "Access-bestand": ("Access file", "Fichier Access"),
    "Doelbestand(en)": ("Output file(s)", "Fichier(s) de destination"),
    "Voortgang": ("Progress", "Progression"),
    "Status": ("Status", "État"),
    "Melding (Statusreden)": ("Message (status reason)", "Message (motif de l'état)"),
    "Logboek - Live status": ("Log - Live status", "Journal - État en direct"),
    "🗑️ Verwijder Selectie": ("🗑️ Remove selection", "🗑️ Supprimer la sélection"),
    "Verwijder de geselecteerde rijen uit de wachtrij (Delete)": ("Remove selected rows from the queue (Delete)", "Supprimer les lignes sélectionnées de la file (Suppr)"),
    "Map openen na afloop": ("Open folder when finished", "Ouvrir le dossier à la fin"),
    "⏹ Stop": ("⏹ Stop", "⏹ Arrêter"),
    "Stop de lopende conversie (gedeeltelijke bestanden worden verwijderd)": ("Stop the current conversion (partial files will be removed)", "Arrêter la conversion en cours (les fichiers partiels seront supprimés)"),
    "▶️ Converteer Selectie": ("▶️ Convert selection", "▶️ Convertir la sélection"),
    "🚀 Converteer Hele Wachtrij": ("🚀 Convert entire queue", "🚀 Convertir toute la file"),
    "Verwerkt alle taken met status 'In de wachtrij', 'Fout' of 'Gestopt'": ("Process all jobs marked 'Queued', 'Error' or 'Stopped'", "Traiter les tâches 'En attente', 'Erreur' ou 'Arrêté'"),
    "Applicatie opgestart en klaar voor gebruik.": ("Application started and ready to use.", "Application démarrée et prête à l'emploi."),
    "In de wachtrij": ("Queued", "En attente"),
    "Bezig": ("Running", "En cours"),
    "Klaar": ("Done", "Terminé"),
    "Fout": ("Error", "Erreur"),
    "Gestopt": ("Stopped", "Arrêté"),
    "Tabellen Selecteren en Bekijken": ("Select and preview tables", "Sélectionner et prévisualiser les tables"),
    "1. Selecteer tabellen (vink aan):": ("1. Select tables (check boxes):", "1. Sélectionner les tables (cocher) :"),
    "🔍 Filter op tabelnaam...": ("🔍 Filter by table name...", "🔍 Filtrer par nom de table..."),
    "☑ Alles aan": ("☑ Select all", "☑ Tout sélectionner"),
    "Vink alle (gefilterde) tabellen aan": ("Select all (filtered) tables", "Cocher toutes les tables (filtrées)"),
    "☐ Alles uit": ("☐ Deselect all", "☐ Tout désélectionner"),
    "Vink alle (gefilterde) tabellen uit": ("Deselect all (filtered) tables", "Décocher toutes les tables (filtrées)"),
    "2. Datavoorbeeld (Eerste {count} rijen):": ("2. Data preview (first {count} rows):", "2. Aperçu des données ({count} premières lignes) :"),
    "✅ Bevestig Selectie": ("✅ Confirm selection", "✅ Confirmer la sélection"),
    "Annuleren": ("Cancel", "Annuler"),
    "{count} van {total} tabellen geselecteerd": ("{count} of {total} tables selected", "{count} tables sur {total} sélectionnées"),
    "Foutmelding": ("Error message", "Message d'erreur"),
    "Alle ({count}) tabellen geselecteerd": ("All ({count}) tables selected", "Toutes les tables ({count}) sélectionnées"),
    "Geen tabellen gevonden in dit bestand": ("No tables found in this file", "Aucune table trouvée dans ce fichier"),
    "Fout bij laden van tabellen": ("Error loading tables", "Erreur de chargement des tables"),
    "Let op: 0 tabellen geselecteerd!": ("Warning: 0 tables selected!", "Attention : aucune table sélectionnée !"),
    "{count} geselecteerd ({names})": ("{count} selected ({names})", "{count} sélectionnées ({names})"),
    "Let op": ("Warning", "Attention"),
    "Kies eerst een geldig Access-bestand.": ("Choose a valid Access file first.", "Choisissez d'abord un fichier Access valide."),
    "Selecteer Access Database": ("Select Access database", "Sélectionner une base Access"),
    "Access Databases (*.accdb *.mdb);;Alle bestanden (*.*)": ("Access databases (*.accdb *.mdb);;All files (*.*)", "Bases Access (*.accdb *.mdb);;Tous les fichiers (*.*)"),
    "Kies de doelmap": ("Choose output folder", "Choisir le dossier de destination"),
    "Incompleet": ("Incomplete", "Incomplet"),
    "Selecteer eerst een Access-bestand en een doelmap.": ("Select an Access file and an output folder first.", "Sélectionnez d'abord un fichier Access et un dossier de destination."),
    "Je hebt geen tabellen geselecteerd.": ("You have not selected any tables.", "Vous n'avez sélectionné aucune table."),
    "Al in wachtrij": ("Already queued", "Déjà dans la file"),
    "Dit bestand staat al met hetzelfde formaat en dezelfde doelmap in de wachtrij.\nToch nogmaals toevoegen?": ("This file is already queued with the same format and output folder.\nAdd it again?", "Ce fichier est déjà dans la file avec le même format et dossier de destination.\nL'ajouter à nouveau ?"),
    "Meerdere CSV's ({name}_...)": ("Multiple CSV files ({name}_...)", "Plusieurs fichiers CSV ({name}_...)"),
    "{count} tabellen geselecteerd": ("{count} tables selected", "{count} tables sélectionnées"),
    "Info": ("Info", "Information"),
    "Er zijn geen bestanden in de wachtrij om te verwerken.": ("There are no queued files to process.", "Aucun fichier en attente à traiter."),
    "Geen selectie": ("No selection", "Aucune sélection"),
    "Selecteer eerst een rij in de wachtrij.": ("Select a row in the queue first.", "Sélectionnez d'abord une ligne dans la file."),
    "Conversie bezig": ("Conversion running", "Conversion en cours"),
    "Er loopt nog een conversie. Stoppen en afsluiten?": ("A conversion is still running. Stop and quit?", "Une conversion est en cours. Arrêter et quitter ?"),
    "Onverwachte fout": ("Unexpected error", "Erreur inattendue"),
    "Installeer xlsxwriter: pip install xlsxwriter": ("Install xlsxwriter: pip install xlsxwriter", "Installez xlsxwriter : pip install xlsxwriter"),
    "⚠ Geen Excel-bibliotheek gevonden; alleen CSV-export beschikbaar (pip install xlsxwriter).": ("⚠ No Excel library found; only CSV export is available (pip install xlsxwriter).", "⚠ Aucune bibliothèque Excel trouvée ; seul l'export CSV est disponible (pip install xlsxwriter)."),
    "ℹ Excel-export via {engine}. Installeer xlsxwriter voor snellere export.": ("ℹ Excel export via {engine}. Install xlsxwriter for faster export.", "ℹ Export Excel via {engine}. Installez xlsxwriter pour un export plus rapide."),
    "Access-driver ontbreekt": ("Access driver missing", "Pilote Access manquant"),
    "❌ De ODBC-driver '{driver}' is niet gevonden.": ("❌ The ODBC driver '{driver}' was not found.", "❌ Le pilote ODBC '{driver}' est introuvable."),
    "De ODBC-driver '{driver}' is niet gevonden.": ("The ODBC driver '{driver}' was not found.", "Le pilote ODBC '{driver}' est introuvable."),
    "De ODBC-driver '{driver}' is niet gevonden.\n\nInstalleer de Microsoft Access Database Engine ({bits}-bit, gelijk aan deze Python-installatie).": ("The ODBC driver '{driver}' was not found.\n\nInstall Microsoft Access Database Engine ({bits}-bit, matching this Python installation).", "Le pilote ODBC '{driver}' est introuvable.\n\nInstallez Microsoft Access Database Engine ({bits} bits, comme cette installation Python)."),
    "Bestand geselecteerd: {name}": ("File selected: {name}", "Fichier sélectionné : {name}"),
    "Doelmap ingesteld op map van het Access-bestand: {path}": ("Output folder set to the Access file's folder: {path}", "Dossier de destination défini sur celui du fichier Access : {path}"),
    "❌ Fout bij ophalen tabellen: {error}": ("❌ Error retrieving tables: {error}", "❌ Erreur de récupération des tables : {error}"),
    "Doelmap ingesteld: {path}": ("Output folder set: {path}", "Dossier de destination défini : {path}"),
    "Toegevoegd aan wachtrij: {name} als {format}": ("Added to queue: {name} as {format}", "Ajouté à la file : {name} au format {format}"),
    "Verwijderd uit wachtrij: {name}": ("Removed from queue: {name}", "Supprimé de la file : {name}"),
    "⏹ Stoppen aangevraagd...": ("⏹ Stop requested...", "⏹ Arrêt demandé..."),
    "Alle geselecteerde taken in wachtrij voltooid.": ("All selected queue jobs finished.", "Toutes les tâches sélectionnées sont terminées."),
    "⏹ Conversie gestopt; resterende taken blijven in de wachtrij.": ("⏹ Conversion stopped; remaining jobs stay queued.", "⏹ Conversion arrêtée ; les tâches restantes restent en attente."),
    "Verbinding maken...": ("Connecting...", "Connexion..."),
    "Start verwerking: {name} (Formaat: {format})": ("Processing started: {name} (Format: {format})", "Début du traitement : {name} (Format : {format})"),
    "Rijen tellen...": ("Counting rows...", "Comptage des lignes..."),
    "Onbekend exportformaat: {format}": ("Unknown export format: {format}", "Format d'export inconnu : {format}"),
    "Door gebruiker gestopt": ("Stopped by user", "Arrêté par l'utilisateur"),
    "⏹ Gestopt: {name} (gedeeltelijke bestanden verwijderd)": ("⏹ Stopped: {name} (partial files removed)", "⏹ Arrêté : {name} (fichiers partiels supprimés)"),
    "❌ Fout bij {name}: {error}": ("❌ Error in {name}: {error}", "❌ Erreur dans {name} : {error}"),
    "Succesvol ({count} rijen)": ("Successful ({count} rows)", "Réussi ({count} lignes)"),
    "✅ Conversie succesvol afgerond: {name}": ("✅ Conversion completed successfully: {name}", "✅ Conversion terminée avec succès : {name}"),
    "Geen Excel-bibliotheek gevonden. Installeer met: pip install xlsxwriter": ("No Excel library found. Install with: pip install xlsxwriter", "Aucune bibliothèque Excel trouvée. Installez avec : pip install xlsxwriter"),
    "-> Inlezen tabel '{name}' ({index}/{total})...": ("-> Reading table '{name}' ({index}/{total})...", "-> Lecture de la table '{name}' ({index}/{total})..."),
    "   (Te groot, opsplitsen in {count} delen...)": ("   (Too large, splitting into {count} parts...)", "   (Trop volumineux, division en {count} parties...)"),
    "Tabel {index}/{total}: {name}": ("Table {index}/{total}: {name}", "Table {index}/{total} : {name}"),
    "Tabel {index}/{total} ingelezen": ("Table {index}/{total} read", "Table {index}/{total} lue"),
    "Excel-bestand opslaan...": ("Saving Excel file...", "Enregistrement du fichier Excel..."),
    "-> Opslaan van Excel-bestand...": ("-> Saving Excel file...", "-> Enregistrement du fichier Excel..."),
    "-> Schrijven CSV: '{name}' ({index}/{total})...": ("-> Writing CSV: '{name}' ({index}/{total})...", "-> Écriture CSV : '{name}' ({index}/{total})..."),
    "CSV {index}/{total}: {name}": ("CSV {index}/{total}: {name}", "CSV {index}/{total} : {name}"),
    "CSV {index}/{total} geschreven": ("CSV {index}/{total} written", "CSV {index}/{total} écrit"),
    "Doelbestand is in gebruik of alleen-lezen: {name}": ("Output file is in use or read-only: {name}", "Le fichier de destination est utilisé ou en lecture seule : {name}"),
}


class LanguageSupport:
    def __init__(self):
        self.language = "nl"
        self.qt_translator = QTranslator()
        self.patterns = []
        for source, targets in TRANSLATIONS.items():
            if "{" in source:
                parts = re.split(r"(\{\w+\})", source)
                pattern = "".join(
                    f"(?P<{part[1:-1]}>.+?)" if part.startswith("{") else re.escape(part)
                    for part in parts
                )
                self.patterns.append((re.compile(pattern, re.DOTALL), targets))

    def set_language(self, language: str) -> None:
        self.language = language if language in ("nl", "en", "fr") else "nl"
        app = QApplication.instance()
        if app is not None:
            app.removeTranslator(self.qt_translator)
            QLocale.setDefault(QLocale(self.language))
            if self.qt_translator.load(
                f"qtbase_{self.language}",
                QLibraryInfo.path(QLibraryInfo.LibraryPath.TranslationsPath),
            ):
                app.installTranslator(self.qt_translator)

    def text(self, source: str) -> str:
        if self.language == "nl":
            return source
        index = 0 if self.language == "en" else 1
        if source in TRANSLATIONS:
            return TRANSLATIONS[source][index]
        for pattern, targets in self.patterns:
            match = pattern.fullmatch(source)
            if match:
                values = match.groupdict()
                if "error" in values:
                    values["error"] = self.text(values["error"])
                return targets[index].format(**values)
        return source

    def set_text(self, widget: QWidget, source: str) -> None:
        widget.setProperty("source_text", source)
        widget.setText(self.text(source))

    def refresh(self, root: QWidget) -> None:
        for widget in [root, *root.findChildren(QWidget)]:
            properties = ["toolTip", "windowTitle"]
            if isinstance(widget, (QLabel, QAbstractButton)):
                properties.append("text")
            if isinstance(widget, QGroupBox):
                properties.append("title")
            if isinstance(widget, QLineEdit):
                properties.append("placeholderText")
            for name in properties:
                source_key = f"source_{name}"
                source = widget.property(source_key)
                if source is None:
                    source = widget.property(name)
                    widget.setProperty(source_key, source)
                if source:
                    widget.setProperty(name, self.text(source))


language_support = LanguageSupport()
tr = language_support.text