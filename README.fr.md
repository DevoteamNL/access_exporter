# Access Exporter

<img src="access_exporter/assets/access-exporter.png" alt="Icône d'Access Exporter" width="96" align="right">

[English](README.md) | [Nederlands](README.nl.md) | **Français**

## À quoi sert-il ?

Access Exporter est un petit programme Windows qui permet de **sortir** les données d'une base Microsoft Access
(`.accdb` / `.mdb`) et, si vous le souhaitez, de les **réintégrer**. Vous n'avez besoin ni de Microsoft Access, ni de
connaissances techniques.

- **Exporter** – convertir les tables d'une ou plusieurs bases Access en un classeur Excel ou en fichiers CSV, par
  exemple pour vérifier, nettoyer ou partager les données avec un client.
- **Modifier** – ouvrir une base Access ou un fichier Excel/CSV exporté et modifier les valeurs directement dans le
  programme, comme dans un tableur.
- **Réintégrer dans Access** – remettre un export Excel/CSV modifié dans la base.

**Votre base d'origine n'est jamais modifiée.** Chaque modification ou fusion est enregistrée dans une *nouvelle
copie* de la base, avec la date et l'heure dans son nom (par exemple `Clients_merged_20260520-143000.accdb`). Si le
résultat ne convient pas, supprimez simplement la copie et recommencez.

Usage typique : exporter la base d'un client vers Excel, laisser le client corriger les données, réintégrer le
fichier corrigé dans une copie de la base et livrer cette copie.

## Sommaire

- [Installation](#installation)
- [Exporter des tables](#exporter-des-tables)
- [Modifier des données](#modifier-des-données)
- [Réintégrer dans Access](#réintégrer-dans-access)
- [Apparence et langue](#apparence-et-langue)
- [Bon à savoir](#bon-à-savoir)
- [Dépannage](#dépannage)
- [Pour les développeurs](#pour-les-développeurs)
- [Licence](#licence)

## Installation

1. Téléchargez `AccessExporter.exe` depuis la
   [dernière version](https://github.com/DevoteamNL/access_exporter/releases/latest). Aucune installation n'est
   nécessaire ; vous pouvez le lancer depuis n'importe quel dossier. Windows peut afficher un avertissement
   SmartScreen la première fois : choisissez **Informations complémentaires → Exécuter quand même**.
2. Le programme a besoin du **Microsoft Access Database Engine** (64 bits), gratuit, la partie d'Access qui lit les
   fichiers de base de données. De nombreuses installations d'Office l'incluent déjà. S'il manque, le programme vous
   le signale au démarrage ; installez-le depuis
   [Microsoft](https://www.microsoft.com/fr-fr/download/details.aspx?id=54920) (`accessdatabaseengine_X64.exe`).

Pour vérifier le téléchargement, comparez sa somme de contrôle SHA256 avec le fichier `AccessExporter.exe.sha256` de
la version : `Get-FileHash .\AccessExporter.exe` dans PowerShell.

## Exporter des tables

Utilisez l'onglet **📤 Exporter**.

1. **📂 1. Choisir un fichier Access** – ou faites glisser un fichier `.accdb`/`.mdb` sur la fenêtre. Les tables sont
   chargées et toutes sélectionnées. Cochez **Inclure les requêtes enregistrées** pour exporter aussi les requêtes
   enregistrées de la base.
2. **⚙ 2. Choisir le format d'export** – **📊 Un fichier Excel (.xlsx)** avec une feuille par table, ou
   **📄 Un fichier CSV par table (.csv)**. Pour le CSV, choisissez **Virgule décimale (1,5)** (Excel français/européen)
   ou **Point décimal (1.5)**.
3. **📋 3. Choisir les tables à exporter** (facultatif) – filtrez, cochez et prévisualisez les premières lignes de
   chaque table.
4. **📁 4. Choisir le dossier de destination** – par défaut le dossier du fichier Access.
5. **➕ 5. Ajouter à la file**, puis **▶️ Convertir la sélection** ou **🚀 Convertir toute la file**.

Vous pouvez mettre plusieurs bases dans la file, chacune avec ses propres réglages. La progression et les messages
s'affichent dans **Journal - État en direct**. **🧹 Effacer les tâches terminées** nettoie la file ;
<kbd>Suppr</kbd> supprime les lignes sélectionnées. Cochez **Ouvrir le dossier à la fin** pour ouvrir
automatiquement le dossier de destination.

À côté de chaque export, le programme écrit un petit fichier `.export.json`. Il indique de quelle base et de quelle
table provient chaque feuille ou fichier CSV, afin que les bonnes tables soient déjà proposées lors de la
réintégration. Gardez-le à côté de l'export si vous comptez réintégrer les données.

Détails : les tables dépassant la limite d'Excel (1 048 576 lignes) sont réparties sur plusieurs feuilles
(`Table_1`, `Table_2`, …) ; les fichiers CSV utilisent `;` et l'UTF-8 et s'ouvrent donc correctement dans Excel ; un
texte comme `=SUM(A1)` reste du texte.

## Modifier des données

Utilisez l'onglet **✏️ Modifier**.

1. Cliquez sur **🗄 Ouvrir une base Access...** ou **📂 Ouvrir un fichier CSV/Excel...** (ou faites glisser un
   fichier sur l'onglet).
2. Choisissez une table ou une feuille dans **Table / feuille :** et modifiez les cellules en double-cliquant
   dessus. Utilisez **🔍 Rechercher dans toutes les colonnes... (Ctrl+F)** pour trouver des lignes, et
   **➕ Ajouter une ligne** et **➖ Supprimer les lignes** (ou <kbd>Suppr</kbd>). Les valeurs qui ne correspondent pas
   à la colonne (par exemple du texte dans une colonne numérique) sont refusées avec un message.
3. Cliquez sur **💾 Enregistrer sous...** et choisissez un nom. Un nom avec la date et l'heure est déjà proposé.
   - Pour une base Access, toutes les modifications sont écrites dans une **nouvelle copie** de la base, qui est
     ensuite ouverte dans l'éditeur.
   - Pour un fichier Excel, toutes les feuilles sont enregistrées ; pour un fichier CSV, la feuille affichée.

Si vous fermez le programme ou ouvrez un autre fichier alors que des modifications ne sont pas enregistrées, le
programme vous demande si vous voulez les abandonner.

## Réintégrer dans Access

Utilisez l'onglet **🔁 Réintégrer dans Access**.

1. **1. Base Access d'origine** – **🗄 Choisir la base...** (la base dont provient l'export).
2. **2. Fichiers modifiés et tables cibles** – **📂 Ajouter des fichiers...** avec vos fichiers Excel/CSV modifiés.
   Vérifiez pour chaque feuille ou fichier la **Table cible dans Access** (remplie automatiquement si possible) ou
   choisissez *(ignorer)*.
3. **3. Comment fusionner ?**
   - **Mettre à jour et ajouter (par clé primaire)** – les lignes modifiées sont mises à jour, les nouvelles lignes
     ajoutées. Recommandé.
   - **Ajouter uniquement les nouvelles lignes** – les lignes existantes ne sont pas touchées.
   - **Remplacer le contenu de la table** – la table est vidée puis remplie avec le fichier.
   - **Supprimer aussi les lignes qui ne sont plus dans le fichier** – facultatif, uniquement avec *Mettre à jour et
     ajouter*.
4. **4. Vérifier et fusionner** – **🔍 Calculer l'aperçu** indique pour chaque table combien de lignes seront
   ajoutées, modifiées, supprimées ou inchangées, ainsi que les éventuelles erreurs. Cliquez ensuite sur
   **💾 Fusionner dans une nouvelle copie...**, confirmez le nom de la nouvelle copie et le résumé.

Toutes les modifications sont écrites en une seule fois : en cas d'erreur ou si vous cliquez sur **⏹ Arrêter**, rien
n'est conservé et la copie inachevée est supprimée.

## Apparence et langue

En haut de la fenêtre :

- **Thème** – **Système** (suit le mode clair/sombre de Windows, par défaut), **Clair** ou **Sombre**.
- **Langue** – **Nederlands** (par défaut), **English** ou **Français**. Les textes, messages et le journal changent
  immédiatement ; vos données ne sont jamais traduites.

Ces deux choix sont mémorisés, avec vos derniers dossiers et réglages d'export (enregistrés dans le registre Windows
sous `HKEY_CURRENT_USER\Software\AccessExporter\AccessExporterPyQt6`).

## Bon à savoir

- **Bases protégées par mot de passe** – le mot de passe vous est demandé. Il n'est conservé que tant que le
  programme est ouvert et n'est jamais enregistré.
- **Les tables sans clé primaire** sont affichées en lecture seule dans l'éditeur. Pour la réintégration, utilisez
  *Ajouter uniquement les nouvelles lignes* ou *Remplacer le contenu de la table*.
- **Grandes tables** – l'éditeur charge au maximum 100 000 lignes par table ; les lignes au-delà ne sont pas
  touchées lors de l'enregistrement.
- **Les pièces jointes, images et autres colonnes binaires** ne sont pas exportées et restent inchangées lors de la
  réintégration.
- **Les colonnes NuméroAuto** sont remplies par Access pour les nouvelles lignes.
- Fermez la base dans Microsoft Access avant de réintégrer ou d'enregistrer ; le programme vous avertit si elle
  semble ouverte.

## Dépannage

| Message ou problème | Que faire |
|---------------------|-----------|
| *Pilote Access manquant* au démarrage | Installez le Microsoft Access Database Engine 64 bits (voir [Installation](#installation)). Depuis le code source, son architecture (32/64 bits) doit correspondre à celle de Python. |
| *Base de données en cours d'utilisation* | Fermez la base dans Access (un fichier `.laccdb`/`.ldb` à côté signifie qu'elle est ouverte) et réessayez. |
| Nombres incorrects dans Excel après un export CSV | Choisissez l'autre réglage décimal à l'étape 2. |
| Une ligne affiche des erreurs dans l'aperçu de réintégration | La valeur ne correspond pas au type de la colonne (par exemple du texte dans une colonne de date). Corrigez-la dans le fichier ou dans l'onglet **✏️ Modifier** et recalculez l'aperçu. |

## Pour les développeurs

Prérequis : Windows, Python 3.10+, l'Access Database Engine avec la même architecture que Python.

```powershell
git clone https://github.com/DevoteamNL/access_exporter.git
cd access_exporter
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt pytest
.venv\Scripts\python.exe access_exporter_pyqt6.py        # lancer le programme
.venv\Scripts\python.exe -m pytest -q tests              # lancer les tests
powershell -ExecutionPolicy Bypass -File .\build.ps1   # construire dist\AccessExporter.exe
```

- Le code se trouve dans le paquet `access_exporter` : `core.py` (accès à la base, sans Qt), `file_io.py`
  (Excel/CSV), `merge.py` (comparaison et réécriture), `main_window.py`, `editor_tab.py`, `merge_tab.py`,
  `themes.py`, `translations.py`. `access_exporter_pyqt6.py` est le point d'entrée.
- Les tests qui nécessitent le pilote Access créent une base temporaire et sont ignorés si le pilote est absent ;
  les autres utilisent la plateforme offscreen de Qt.
- Les textes de l'interface sont écrits en néerlandais ; l'anglais et le français se trouvent dans `TRANSLATIONS`
  dans `translations.py`.
- L'icône est dessinée dans `access_exporter/assets/access-exporter.svg` ; lancez
  `python tools/generate_app_icon.py` pour régénérer le `.ico` et le `.png`.
- La publication d'une version GitHub construit l'exe et le téléverse avec sa somme de contrôle
  (`.github/workflows/release-assets.yml`). Les modifications sont listées dans [CHANGELOG.md](CHANGELOG.md).

## Licence

[GNU General Public License v3.0](LICENSE) (GPL-3.0-only), conformément à la licence de PyQt6, inclus dans le
programme.
