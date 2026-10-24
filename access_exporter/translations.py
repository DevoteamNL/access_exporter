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
    # --- Thema en tabbladen ---
    "Thema": ("Theme", "Thème"),
    "Systeem": ("System", "Système"),
    "Licht": ("Light", "Clair"),
    "Donker": ("Dark", "Sombre"),
    "📤 Exporteren": ("📤 Export", "📤 Exporter"),
    "✏️ Bewerken": ("✏️ Edit", "✏️ Modifier"),
    "🔁 Terugzetten naar Access": ("🔁 Merge back into Access", "🔁 Réintégrer dans Access"),
    # --- Exporteren: nieuwe opties ---
    "Of sleep een .accdb/.mdb-bestand naar dit venster": ("Or drag an .accdb/.mdb file onto this window", "Ou faites glisser un fichier .accdb/.mdb sur cette fenêtre"),
    "Decimaalteken in CSV-bestanden": ("Decimal separator in CSV files", "Séparateur décimal dans les fichiers CSV"),
    "Decimale komma (1,5)": ("Decimal comma (1,5)", "Virgule décimale (1,5)"),
    "Decimale punt (1.5)": ("Decimal point (1.5)", "Point décimal (1.5)"),
    "Ook opgeslagen query's": ("Include saved queries", "Inclure les requêtes enregistrées"),
    "Toon ook opgeslagen selectiequery's in de tabellijst, zodat je hun resultaat kunt exporteren": ("Also list saved select queries, so you can export their results", "Afficher aussi les requêtes de sélection enregistrées pour pouvoir exporter leur résultat"),
    "Dubbelklik op een afgeronde Excel-taak om het bestand in de editor te openen": ("Double-click a finished Excel job to open the file in the editor", "Double-cliquez sur une tâche Excel terminée pour ouvrir le fichier dans l'éditeur"),
    "🧹 Wis afgeronde taken": ("🧹 Clear finished jobs", "🧹 Effacer les tâches terminées"),
    "Verwijder alle taken met status 'Klaar' uit de wachtrij": ("Remove all jobs with status 'Done' from the queue", "Retirer de la file toutes les tâches au statut 'Terminé'"),
    "Open bestand": ("Open file", "Ouvrir un fichier"),
    "CSV-bestanden (*.csv)": ("CSV files (*.csv)", "Fichiers CSV (*.csv)"),
    # --- Wachtwoord ---
    "Wachtwoord vereist": ("Password required", "Mot de passe requis"),
    "Voer het wachtwoord in voor {name}:": ("Enter the password for {name}:", "Saisissez le mot de passe de {name} :"),
    "Onjuist wachtwoord. Voer het wachtwoord in voor {name}:": ("Incorrect password. Enter the password for {name}:", "Mot de passe incorrect. Saisissez le mot de passe de {name} :"),
    "De database is met een wachtwoord beveiligd, of het opgegeven wachtwoord klopt niet.": ("The database is password protected, or the password is incorrect.", "La base de données est protégée par un mot de passe, ou le mot de passe est incorrect."),
    "De kopie mag niet hetzelfde bestand zijn als het origineel.": ("The copy must not be the same file as the original.", "La copie ne peut pas être le même fichier que l'original."),
    "Niet-ondersteund bestandstype: {name}": ("Unsupported file type: {name}", "Type de fichier non pris en charge : {name}"),
    # --- Editor ---
    "Ja": ("Yes", "Oui"),
    "Nee": ("No", "Non"),
    "🗄 Open Access-database...": ("🗄 Open Access database...", "🗄 Ouvrir une base Access..."),
    "Bewerk tabellen direct; opslaan maakt altijd een nieuwe kopie van de database": ("Edit tables directly; saving always creates a new copy of the database", "Modifiez les tables directement ; l'enregistrement crée toujours une nouvelle copie de la base"),
    "📂 Open CSV/Excel-bestand...": ("📂 Open CSV/Excel file...", "📂 Ouvrir un fichier CSV/Excel..."),
    "Bewerk een geëxporteerd bestand; opslaan maakt altijd een nieuw bestand": ("Edit an exported file; saving always creates a new file", "Modifiez un fichier exporté ; l'enregistrement crée toujours un nouveau fichier"),
    "Geen database of bestand geopend": ("No database or file open", "Aucune base ou fichier ouvert"),
    "Tabel / blad:": ("Table / sheet:", "Table / feuille :"),
    "🔍 Zoeken in alle kolommen... (Ctrl+F)": ("🔍 Search all columns... (Ctrl+F)", "🔍 Rechercher dans toutes les colonnes... (Ctrl+F)"),
    "➕ Rij toevoegen": ("➕ Add row", "➕ Ajouter une ligne"),
    "➖ Rijen verwijderen": ("➖ Delete rows", "➖ Supprimer les lignes"),
    "Verwijder de geselecteerde rijen (Delete)": ("Delete the selected rows (Delete)", "Supprimer les lignes sélectionnées (Suppr)"),
    "💾 Opslaan als...": ("💾 Save as...", "💾 Enregistrer sous..."),
    "{count} rijen · {columns} kolommen": ("{count} rows · {columns} columns", "{count} lignes · {columns} colonnes"),
    "{count} rijen · {columns} kolommen · niet-opgeslagen wijzigingen": ("{count} rows · {columns} columns · unsaved changes", "{count} lignes · {columns} colonnes · modifications non enregistrées"),
    "Niet-opgeslagen wijzigingen": ("Unsaved changes", "Modifications non enregistrées"),
    "Wijzigingen in {name} opslaan?": ("Save changes to {name}?", "Enregistrer les modifications de {name} ?"),
    "Export-bestanden (*.xlsx *.csv);;Alle bestanden (*.*)": ("Exported files (*.xlsx *.csv);;All files (*.*)", "Fichiers exportés (*.xlsx *.csv);;Tous les fichiers (*.*)"),
    "Kan bestand niet openen: {error}": ("Cannot open file: {error}", "Impossible d'ouvrir le fichier : {error}"),
    "Geopend in editor: {name}": ("Opened in editor: {name}", "Ouvert dans l'éditeur : {name}"),
    "Ongeldige waarde": ("Invalid value", "Valeur non valide"),
    "❌ Fout bij laden van tabel '{name}': {error}": ("❌ Error loading table '{name}': {error}", "❌ Erreur lors du chargement de la table '{name}' : {error}"),
    "Deze tabel heeft geen primaire sleutel en kan alleen bekeken worden.": ("This table has no primary key and is read-only.", "Cette table n'a pas de clé primaire et est en lecture seule."),
    "Alleen de eerste {count} van {total} rijen zijn geladen.": ("Only the first {count} of {total} rows are loaded.", "Seules les {count} premières lignes sur {total} sont chargées."),
    "Excel-bestand (*.xlsx);;CSV-bestand (*.csv)": ("Excel file (*.xlsx);;CSV file (*.csv)", "Fichier Excel (*.xlsx);;Fichier CSV (*.csv)"),
    "Opslaan als": ("Save as", "Enregistrer sous"),
    "ℹ Alleen blad '{name}' is als CSV opgeslagen.": ("ℹ Only sheet '{name}' was saved as CSV.", "ℹ Seule la feuille '{name}' a été enregistrée en CSV."),
    "Opslaan mislukt: {error}": ("Saving failed: {error}", "Échec de l'enregistrement : {error}"),
    "✅ Opgeslagen als: {name}": ("✅ Saved as: {name}", "✅ Enregistré sous : {name}"),
    "Er zijn geen wijzigingen om op te slaan.": ("There are no changes to save.", "Il n'y a aucune modification à enregistrer."),
    "Kan niet opslaan": ("Cannot save", "Impossible d'enregistrer"),
    "Opslaan als nieuwe kopie van de database": ("Save as a new copy of the database", "Enregistrer comme nouvelle copie de la base"),
    "Access Databases (*.accdb *.mdb)": ("Access Databases (*.accdb *.mdb)", "Bases de données Access (*.accdb *.mdb)"),
    "Database in gebruik": ("Database in use", "Base de données en cours d'utilisation"),
    "De database lijkt geopend te zijn in Access. Sluit Access voor een betrouwbare kopie. Toch doorgaan?": ("The database appears to be open in Access. Close Access for a reliable copy. Continue anyway?", "La base semble ouverte dans Access. Fermez Access pour une copie fiable. Continuer quand même ?"),
    "✅ {count} wijzigingen opgeslagen in nieuwe kopie: {name}": ("✅ {count} changes saved to new copy: {name}", "✅ {count} modifications enregistrées dans une nouvelle copie : {name}"),
    # --- Omzetfouten en opmerkingen bij samenvoegen ---
    "'{value}' is geen getal": ("'{value}' is not a number", "'{value}' n'est pas un nombre"),
    "'{value}' is geen geheel getal": ("'{value}' is not a whole number", "'{value}' n'est pas un nombre entier"),
    "'{value}' is geen geldige datum": ("'{value}' is not a valid date", "'{value}' n'est pas une date valide"),
    "'{value}' is geen Ja/Nee-waarde": ("'{value}' is not a Yes/No value", "'{value}' n'est pas une valeur Oui/Non"),
    "tekst is langer dan {count} tekens": ("text is longer than {count} characters", "le texte dépasse {count} caractères"),
    "Kolom '{column}' bestaat niet in Access en wordt overgeslagen": ("Column '{column}' does not exist in Access and is skipped", "La colonne '{column}' n'existe pas dans Access et est ignorée"),
    "Kolom '{column}' ({type}) wordt niet ondersteund en overgeslagen": ("Column '{column}' ({type}) is not supported and is skipped", "La colonne '{column}' ({type}) n'est pas prise en charge et est ignorée"),
    "Kolom '{column}' ontbreekt in het bestand en blijft ongewijzigd": ("Column '{column}' is missing from the file and stays unchanged", "La colonne '{column}' est absente du fichier et reste inchangée"),
    "Kolom '{column}': {error}": ("Column '{column}': {error}", "Colonne '{column}' : {error}"),
    "Tabel '{name}' heeft geen primaire sleutel; kies 'Alleen nieuwe rijen toevoegen' of 'Tabelinhoud vervangen'": ("Table '{name}' has no primary key; choose 'Only add new rows' or 'Replace table contents'", "La table '{name}' n'a pas de clé primaire ; choisissez 'Ajouter uniquement les nouvelles lignes' ou 'Remplacer le contenu de la table'"),
    "Sleutelkolom '{column}' ontbreekt in het bestand": ("Key column '{column}' is missing from the file", "La colonne clé '{column}' est absente du fichier"),
    "Rij {row}, kolom '{column}': {error}": ("Row {row}, column '{column}': {error}", "Ligne {row}, colonne '{column}' : {error}"),
    "Rij {row}: sleutelwaarde ontbreekt": ("Row {row}: key value is missing", "Ligne {row} : valeur de clé manquante"),
    "Rij {row}: dubbele sleutelwaarde {value}": ("Row {row}: duplicate key value {value}", "Ligne {row} : valeur de clé en double {value}"),
    "Er worden geen rijen verwijderd omdat er fouten zijn": ("No rows are deleted because there are errors", "Aucune ligne n'est supprimée car il y a des erreurs"),
    # --- Terugzetten naar Access ---
    "Zet bewerkte CSV/Excel-exports terug in Access. Er wordt altijd een nieuwe kopie van de database gemaakt; het origineel blijft ongewijzigd.": ("Merge edited CSV/Excel exports back into Access. A new copy of the database is always created; the original stays unchanged.", "Réintégrez des exports CSV/Excel modifiés dans Access. Une nouvelle copie de la base est toujours créée ; l'original reste inchangé."),
    "1. Originele Access-database": ("1. Original Access database", "1. Base Access d'origine"),
    "🗄 Kies database...": ("🗄 Choose database...", "🗄 Choisir la base..."),
    "Geen database geselecteerd": ("No database selected", "Aucune base sélectionnée"),
    "2. Bewerkte bestanden en doeltabellen": ("2. Edited files and target tables", "2. Fichiers modifiés et tables cibles"),
    "📂 Bestanden toevoegen...": ("📂 Add files...", "📂 Ajouter des fichiers..."),
    "Bestanden toevoegen": ("Add files", "Ajouter des fichiers"),
    "🗑 Lijst wissen": ("🗑 Clear list", "🗑 Vider la liste"),
    "Bestand / blad": ("File / sheet", "Fichier / feuille"),
    "Rijen": ("Rows", "Lignes"),
    "Doeltabel in Access": ("Target table in Access", "Table cible dans Access"),
    "(overslaan)": ("(skip)", "(ignorer)"),
    "3. Hoe samenvoegen?": ("3. How to merge?", "3. Comment fusionner ?"),
    "Bijwerken en toevoegen (op primaire sleutel)": ("Update and add (by primary key)", "Mettre à jour et ajouter (par clé primaire)"),
    "Bestaande rijen worden bijgewerkt op basis van de primaire sleutel; nieuwe rijen worden toegevoegd": ("Existing rows are updated by primary key; new rows are added", "Les lignes existantes sont mises à jour par clé primaire ; les nouvelles lignes sont ajoutées"),
    "Alleen nieuwe rijen toevoegen": ("Only add new rows", "Ajouter uniquement les nouvelles lignes"),
    "Rijen waarvan de sleutel al bestaat worden overgeslagen; er wordt niets gewijzigd of verwijderd": ("Rows whose key already exists are skipped; nothing is changed or deleted", "Les lignes dont la clé existe déjà sont ignorées ; rien n'est modifié ni supprimé"),
    "Tabelinhoud vervangen": ("Replace table contents", "Remplacer le contenu de la table"),
    "Alle rijen in de gekoppelde tabellen worden verwijderd en vervangen door de rijen uit het bestand": ("All rows in the linked tables are deleted and replaced by the rows from the file", "Toutes les lignes des tables liées sont supprimées et remplacées par celles du fichier"),
    "Rijen die niet meer in het bestand staan ook verwijderen": ("Also delete rows that are no longer in the file", "Supprimer aussi les lignes qui ne sont plus dans le fichier"),
    "Alleen bij 'Bijwerken en toevoegen'. Er wordt niets verwijderd als er fouten zijn.": ("Only with 'Update and add'. Nothing is deleted if there are errors.", "Uniquement avec 'Mettre à jour et ajouter'. Rien n'est supprimé en cas d'erreurs."),
    "4. Controleren en samenvoegen": ("4. Review and merge", "4. Vérifier et fusionner"),
    "🔍 Voorbeeld berekenen": ("🔍 Calculate preview", "🔍 Calculer l'aperçu"),
    "Toont wat er per tabel verandert, zonder iets op te slaan": ("Shows what changes per table, without saving anything", "Montre ce qui change par table, sans rien enregistrer"),
    "💾 Samenvoegen in nieuwe kopie...": ("💾 Merge into new copy...", "💾 Fusionner dans une nouvelle copie..."),
    "Tabel": ("Table", "Table"),
    "Toevoegen": ("Add", "Ajouter"),
    "Wijzigen": ("Change", "Modifier"),
    "Verwijderen": ("Delete", "Supprimer"),
    "Ongewijzigd": ("Unchanged", "Inchangées"),
    "Fouten": ("Errors", "Erreurs"),
    "alles": ("all", "tout"),
    "Wacht tot het samenvoegen klaar is of druk op Stop.": ("Wait until the merge has finished or press Stop.", "Attendez la fin de la fusion ou appuyez sur Arrêter."),
    "⏹ Samenvoegen gestopt; er is niets opgeslagen.": ("⏹ Merge stopped; nothing was saved.", "⏹ Fusion arrêtée ; rien n'a été enregistré."),
    "❌ Samenvoegen mislukt: {error}": ("❌ Merge failed: {error}", "❌ Échec de la fusion : {error}"),
    "Samenvoegen mislukt: {error}": ("Merge failed: {error}", "Échec de la fusion : {error}"),
    "Voorbeeld berekenen voor {name}...": ("Calculating preview for {name}...", "Calcul de l'aperçu pour {name}..."),
    "⚠ {name}: nog {count} fouten": ("⚠ {name}: {count} more errors", "⚠ {name} : {count} erreurs supplémentaires"),
    "⚠ {name}: {error}": ("⚠ {name}: {error}", "⚠ {name} : {error}"),
    "ℹ {name}: {error}": ("ℹ {name}: {error}", "ℹ {name} : {error}"),
    "Totaal {count} wijzigingen klaar om samen te voegen.": ("{count} changes ready to merge in total.", "{count} modifications prêtes à fusionner au total."),
    "Totaal {count} wijzigingen klaar om samen te voegen. {errors} rijen of tabellen worden overgeslagen vanwege fouten (zie log).": ("{count} changes ready to merge in total. {errors} rows or tables are skipped because of errors (see log).", "{count} modifications prêtes à fusionner au total. {errors} lignes ou tables sont ignorées en raison d'erreurs (voir le journal)."),
    "Samengevoegde database opslaan als": ("Save merged database as", "Enregistrer la base fusionnée sous"),
    "Kies een andere naam: het origineel wordt nooit overschreven.": ("Choose another name: the original is never overwritten.", "Choisissez un autre nom : l'original n'est jamais écrasé."),
    "Bevestigen": ("Confirm", "Confirmer"),
    "{count} wijzigingen samenvoegen in {name}?": ("Merge {count} changes into {name}?", "Fusionner {count} modifications dans {name} ?"),
    "{count} wijzigingen samenvoegen in {name}? Let op: dit omvat {deletes} verwijderingen.": ("Merge {count} changes into {name}? Note: this includes {deletes} deletions.", "Fusionner {count} modifications dans {name} ? Attention : cela comprend {deletes} suppressions."),
    "Samenvoegen in nieuwe kopie: {name}...": ("Merging into new copy: {name}...", "Fusion dans une nouvelle copie : {name}..."),
    "✅ Samengevoegd in nieuwe kopie: {name}": ("✅ Merged into new copy: {name}", "✅ Fusionné dans une nouvelle copie : {name}"),
    "Samengevoegd in nieuwe kopie: {name}": ("Merged into new copy: {name}", "Fusionné dans une nouvelle copie : {name}"),
}

# Plaatshouders met een aantal krijgen het duizendtalscheidingsteken van de taal
NUMERIEKE_PLAATSHOUDERS = {"count", "total", "errors", "deletes"}
DUIZENDTALLEN = {"nl": ".", "en": ",", "fr": "\u00a0"}


def formatteer_getal(tekst: str, taal: str) -> str:
    if not tekst.isdigit() or len(tekst) <= 3:
        return tekst
    return f"{int(tekst):,}".replace(",", DUIZENDTALLEN.get(taal, ","))


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
                self.patterns.append((re.compile(pattern, re.DOTALL), source, targets))

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
        if source in TRANSLATIONS:
            return source if self.language == "nl" else TRANSLATIONS[source][0 if self.language == "en" else 1]
        for pattern, template, targets in self.patterns:
            match = pattern.fullmatch(source)
            if match:
                # Ook in het Nederlands: getallen opmaken; ingebedde meldingen (bijv. fouten) mee vertalen
                values = {
                    name: formatteer_getal(value, self.language) if name in NUMERIEKE_PLAATSHOUDERS else self.text(value)
                    for name, value in match.groupdict().items()
                }
                if self.language != "nl":
                    template = targets[0 if self.language == "en" else 1]
                return template.format(**values)
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