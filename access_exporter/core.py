"""Database- en bestandshulpfuncties zonder Qt (Access via pyodbc, pandas)."""

import json
import os
import re
import shutil
import subprocess
from collections.abc import Iterator
from contextlib import closing
from datetime import datetime
from importlib.util import find_spec
from typing import Any

import pandas as pd
import pyodbc

from . import __version__

ACCESS_DRIVER = "Microsoft Access Driver (*.mdb, *.accdb)"
ACCESS_EXTENSIES = (".accdb", ".mdb")
EXCEL_ENGINE = next((m for m in ("xlsxwriter", "openpyxl") if find_spec(m)), None)
EXCEL_MAX_DATA_RIJEN = 1_048_575  # 1.048.576 rijen per blad, min de kopregel
EXCEL_MAX_BLADNAAM = 31
BLOKGROOTTE = 50_000  # Aantal rijen dat per keer in het geheugen wordt geladen
PREVIEW_RIJEN = 10
CSV_SCHEIDINGSTEKEN = ";"
CSV_ENCODING = "utf-8-sig"
MANIFEST_ACHTERVOEGSEL = ".export.json"

STATUS_WACHTRIJ = "In de wachtrij"
STATUS_BEZIG = "Bezig"
STATUS_KLAAR = "Klaar"
STATUS_FOUT = "Fout"
STATUS_GESTOPT = "Gestopt"

_ONGELDIGE_BLADTEKENS = re.compile(r"[\[\]:*?/\\]")


class WachtwoordNodig(Exception):
    """De database is met een wachtwoord beveiligd, of het opgegeven wachtwoord klopt niet."""


def access_driver_beschikbaar() -> bool:
    return ACCESS_DRIVER in pyodbc.drivers()


def is_access_bestand(pad: str) -> bool:
    return pad.lower().endswith(ACCESS_EXTENSIES)


def is_wachtwoordfout(fout: Exception) -> bool:
    tekst = str(fout).lower()
    return "-1905" in tekst or "password" in tekst or "wachtwoord" in tekst


def _odbc_waarde(waarde: str) -> str:
    # Accolades zorgen dat ; en = in een wachtwoord geen deel van de verbindingsreeks worden
    return "{" + waarde.replace("}", "}}") + "}"


def verbind_access(pad: str, wachtwoord: str = "", alleen_lezen: bool = True) -> pyodbc.Connection:
    """Opent een Access-database. Standaard alleen-lezen: de bron-database wordt nooit gewijzigd.
    Schrijven gebeurt alleen op een kopie (zie maak_kopie), in één transactie (autocommit uit)."""
    verbinding = f"DRIVER={{{ACCESS_DRIVER}}};DBQ={pad};"
    if wachtwoord:
        verbinding += f"PWD={_odbc_waarde(wachtwoord)};"
    try:
        return pyodbc.connect(verbinding, readonly=alleen_lezen, autocommit=False)
    except pyodbc.Error as e:
        if is_wachtwoordfout(e):
            raise WachtwoordNodig(str(e)) from e
        raise


def quote_naam(naam: str) -> str:
    # Backticks i.p.v. [] zodat een '?' in de naam niet als ODBC-parameter wordt gezien
    return f"`{naam}`"


quote_tabel = quote_naam


def haal_tabellen_op(pad: str, wachtwoord: str = "", met_queries: bool = False) -> list[str]:
    """Tabellen (en desgewenst opgeslagen selectiequery's) van de database, systeemtabellen uitgezonderd."""
    soorten = ("TABLE", "VIEW") if met_queries else ("TABLE",)
    with closing(verbind_access(pad, wachtwoord)) as conn:
        objecten = [(t.table_type, t.table_name) for t in conn.cursor().tables()
                    if t.table_type in soorten and not t.table_name.startswith(("MSys", "~"))]
    # Eerst tabellen, dan query's
    return [naam for soort in soorten for s, naam in objecten if s == soort]


def tel_rijen(conn: pyodbc.Connection, tabel: str) -> int:
    rij = conn.cursor().execute(f"SELECT COUNT(*) FROM {quote_naam(tabel)}").fetchone()
    return int(rij[0]) if rij else 0


def lees_in_blokken(conn: pyodbc.Connection, tabel: str, blokgrootte: int = BLOKGROOTTE) -> Iterator[pd.DataFrame]:
    """Leest een tabel in blokken in, zodat grote tabellen niet in één keer in het geheugen staan.
    Een lege tabel levert één leeg DataFrame op (met kolomnamen)."""
    with closing(conn.cursor()) as cursor:
        cursor.execute(f"SELECT * FROM {quote_naam(tabel)}")
        kolommen = [c[0] for c in cursor.description]
        leeg = True
        while rijen := cursor.fetchmany(blokgrootte):
            leeg = False
            df = pd.DataFrame.from_records([tuple(r) for r in rijen], columns=kolommen, coerce_float=True)
            # Nullable types: een getalkolom met lege waarden blijft in elk blok een gehele-getallenkolom
            yield df.convert_dtypes()
        if leeg:
            yield pd.DataFrame(columns=kolommen)


def lees_rijen(conn: pyodbc.Connection, tabel: str, limiet: int | None = None) -> tuple[list[str], list[tuple[Any, ...]]]:
    """Leest de ruwe waarden (zonder typeconversie) van een tabel, optioneel de eerste `limiet` rijen."""
    top = f"TOP {int(limiet)} " if limiet else ""
    with closing(conn.cursor()) as cursor:
        cursor.execute(f"SELECT {top}* FROM {quote_naam(tabel)}")
        kolommen = [c[0] for c in cursor.description]
        return kolommen, [tuple(r) for r in cursor.fetchall()]


def maak_bladnaam(naam: str, gebruikt: set[str], achtervoegsel: str = "") -> str:
    """Geldige, unieke Excel-bladnaam (max. 31 tekens, geen []:*?/\\, niet hoofdlettergevoelig uniek)."""
    basis = _ONGELDIGE_BLADTEKENS.sub("_", naam).strip("'") or "Blad"
    if basis.lower() == "history":  # Gereserveerd door Excel
        basis += "_"
    kandidaat = basis[:EXCEL_MAX_BLADNAAM - len(achtervoegsel)] + achtervoegsel
    teller = 2
    while kandidaat.lower() in gebruikt:
        extra = f"~{teller}{achtervoegsel}"
        kandidaat = basis[:EXCEL_MAX_BLADNAAM - len(extra)] + extra
        teller += 1
    gebruikt.add(kandidaat.lower())
    return kandidaat


def veilige_bestandsnaam(tekst: str) -> str:
    return "".join(x for x in tekst if x.isalnum() or x in " _-").strip()


def maak_csv_naam(prefix: str, tabel: str, gebruikt: set[str]) -> str:
    basis = f"{prefix}_{veilige_bestandsnaam(tabel) or 'tabel'}"
    naam, teller = basis, 2
    while naam.lower() in gebruikt:
        naam = f"{basis}_{teller}"
        teller += 1
    gebruikt.add(naam.lower())
    return f"{naam}.csv"


def controleer_schrijfbaar(pad: str) -> None:
    """Faalt direct als een bestaand doelbestand vergrendeld is (bijv. geopend in Excel)."""
    if os.path.exists(pad):
        try:
            open(pad, "r+b").close()
        except PermissionError:
            raise PermissionError(f"Doelbestand is in gebruik of alleen-lezen: {os.path.basename(pad)}") from None


def open_in_verkenner(pad: str) -> None:
    pad = os.path.normpath(pad)
    if os.path.isfile(pad):
        subprocess.Popen(f'explorer /select,"{pad}"')
    else:
        subprocess.Popen(["explorer", pad])


def tijdstempel() -> str:
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def voorgestelde_kopienaam(origineel: str, label: str) -> str:
    """Bijv. C:\\map\\Klanten.accdb -> C:\\map\\Klanten_merged_20261024-031500.accdb"""
    basis, extensie = os.path.splitext(origineel)
    return f"{basis}_{label}_{tijdstempel()}{extensie}"


def lockbestand(pad: str) -> str:
    basis, extensie = os.path.splitext(pad)
    return basis + (".ldb" if extensie.lower() == ".mdb" else ".laccdb")


def is_vergrendeld(pad: str) -> bool:
    """True als Access (of een ander programma) de database open heeft."""
    return os.path.exists(lockbestand(pad))


def controleer_kopiedoel(origineel: str, doel: str) -> None:
    """Weigert een doel dat hetzelfde bestand is als het origineel (ook via 8.3-namen of links) of niet schrijfbaar is."""
    zelfde = os.path.normcase(os.path.abspath(origineel)) == os.path.normcase(os.path.abspath(doel))
    if zelfde or (os.path.exists(doel) and os.path.samefile(origineel, doel)):
        raise ValueError("De kopie mag niet hetzelfde bestand zijn als het origineel.")
    controleer_schrijfbaar(doel)


def maak_kopie(origineel: str, doel: str) -> None:
    """Kopieert de database naar een nieuw bestand; het origineel blijft ongewijzigd."""
    controleer_kopiedoel(origineel, doel)
    shutil.copyfile(origineel, doel)


# ---------------------------------------------------------
# Manifest: legt vast welke tabel in welk blad/bestand staat, zodat terugzetten betrouwbaar werkt
# ---------------------------------------------------------
def manifest_pad(doelmap: str, naam_zonder_ext: str, export_type: str) -> str:
    return os.path.join(doelmap, f"{naam_zonder_ext}.{export_type}{MANIFEST_ACHTERVOEGSEL}")


def schrijf_manifest(pad: str, database: str, export_type: str, tabellen: list[dict[str, Any]],
                     csv_decimaal: str = ",") -> None:
    gegevens = {
        "app": "Access Exporter",
        "version": __version__,
        "database": os.path.abspath(database),
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "format": export_type,
        "csv_decimal": csv_decimaal,
        "tables": tabellen,
    }
    with open(pad, "w", encoding="utf-8") as f:
        json.dump(gegevens, f, ensure_ascii=False, indent=2)


def lees_manifest(pad: str) -> dict[str, Any] | None:
    try:
        with open(pad, encoding="utf-8") as f:
            gegevens = json.load(f)
    except (OSError, ValueError):
        return None
    return gegevens if isinstance(gegevens, dict) and isinstance(gegevens.get("tables"), list) else None


def zoek_manifest(bestand: str) -> dict[str, Any] | None:
    """Zoekt het manifest bij een geëxporteerd bestand (Excel: <bestand>.export.json, CSV: in dezelfde map)."""
    if bestand.lower().endswith(".xlsx"):
        return lees_manifest(bestand + MANIFEST_ACHTERVOEGSEL) if os.path.exists(bestand + MANIFEST_ACHTERVOEGSEL) else None
    map_ = os.path.dirname(os.path.abspath(bestand))
    naam = os.path.basename(bestand).lower()
    try:
        kandidaten = [f for f in os.listdir(map_) if f.lower().endswith(".csv" + MANIFEST_ACHTERVOEGSEL)]
    except OSError:
        return None
    for kandidaat in kandidaten:
        gegevens = lees_manifest(os.path.join(map_, kandidaat))
        if gegevens and any(str(t.get("file", "")).lower() == naam for t in gegevens["tables"]):
            return gegevens
    return None
