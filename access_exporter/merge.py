"""Wijzigingen uit geëxporteerde/bewerkte gegevens terugzetten in een KOPIE van een Access-database (zonder Qt).

Werkwijze: schema lezen -> bestandswaarden omzetten naar Access-typen -> vergelijken op primaire sleutel
(plan met toevoegen/wijzigen/verwijderen) -> origineel kopiëren -> plan in één transactie uitvoeren op de kopie.
"""

import math
import os
import re
from collections.abc import Callable
from contextlib import closing
from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any

import pyodbc

from .core import (
    _ONGELDIGE_BLADTEKENS, EXCEL_MAX_BLADNAAM, lees_rijen, maak_kopie, quote_naam, veilige_bestandsnaam,
    verbind_access, zoek_manifest,
)
from .file_io import Tabel, lees_export

MODUS_BIJWERKEN = "bijwerken"   # Bestaande rijen bijwerken en nieuwe toevoegen (op primaire sleutel)
MODUS_TOEVOEGEN = "toevoegen"   # Alleen rijen toevoegen die nog niet bestaan
MODUS_VERVANGEN = "vervangen"   # Tabelinhoud volledig vervangen door het bestand

TEKST_TYPES = {"VARCHAR", "CHAR", "LONGCHAR", "WVARCHAR", "WCHAR", "WLONGVARCHAR", "TEXT", "MEMO", "GUID"}
GEHEEL_TYPES = {"COUNTER", "INTEGER", "SMALLINT", "BYTE", "LONG", "TINYINT", "BIGINT"}
DECIMAAL_TYPES = {"CURRENCY", "DECIMAL", "NUMERIC", "MONEY"}
DRIJVEND_TYPES = {"DOUBLE", "REAL", "FLOAT", "SINGLE"}
DATUM_TYPES = {"DATETIME", "DATE", "TIME", "TIMESTAMP"}
JANEE_TYPES = {"BIT", "YESNO"}
BINAIR_TYPES = {"LONGBINARY", "BINARY", "VARBINARY", "LONGVARBINARY", "OLEOBJECT", "IMAGE"}

WAAR = {"1", "-1", "true", "waar", "ja", "j", "yes", "y", "vrai", "oui", "x", "✓"}
ONWAAR = {"0", "false", "onwaar", "nee", "n", "no", "faux", "non"}
DATUM_FORMATEN = [f + t for f in ("%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y") for t in ("", " %H:%M", " %H:%M:%S")]


class OmzetFout(ValueError):
    pass


class Gestopt(Exception):
    pass


@dataclass
class Kolom:
    naam: str
    type: str
    nullable: bool = True
    grootte: int | None = None

    @property
    def autonummer(self) -> bool:
        return self.type == "COUNTER"

    @property
    def ondersteund(self) -> bool:
        return self.type not in BINAIR_TYPES


@dataclass
class Schema:
    tabel: str
    kolommen: list[Kolom]
    sleutel: list[str]

    def kolom(self, naam: str) -> Kolom | None:
        zoek = naam.strip().casefold()
        return next((k for k in self.kolommen if k.naam.casefold() == zoek), None)


@dataclass
class TabelPlan:
    tabel: str
    modus: str
    sleutel: list[str] = field(default_factory=list)
    invoegen: list[dict[str, Any]] = field(default_factory=list)
    wijzigen: list[tuple[tuple[Any, ...], dict[str, Any]]] = field(default_factory=list)
    verwijderen: list[tuple[Any, ...]] = field(default_factory=list)
    alles_verwijderen: bool = False
    ongewijzigd: int = 0
    fouten: list[str] = field(default_factory=list)
    opmerkingen: list[str] = field(default_factory=list)
    geblokkeerd: bool = False  # De hele tabel kan niet worden verwerkt (zie fouten)

    @property
    def aantal_bewerkingen(self) -> int:
        return len(self.invoegen) + len(self.wijzigen) + len(self.verwijderen) + int(self.alles_verwijderen)


# ---------------------------------------------------------
# Schema
# ---------------------------------------------------------
def kies_sleutel(indexen: dict[str, list[tuple[int, str]]], kolommen: list[Kolom]) -> list[str]:
    """Primaire sleutel: index 'PrimaryKey' (Access), anders de kleinste unieke index zonder lege kolommen."""
    def kolommen_van(naam: str) -> list[str]:
        return [k for _, k in sorted(indexen[naam])]

    for naam in indexen:
        if naam.casefold() == "primarykey":
            return kolommen_van(naam)
    verplicht = {k.naam.casefold() for k in kolommen if not k.nullable}
    kandidaten = [n for n in indexen if all(k.casefold() in verplicht for k in kolommen_van(n))]
    if not kandidaten:
        return []
    return kolommen_van(min(kandidaten, key=lambda n: (len(indexen[n]), n)))


def lees_schema(conn: pyodbc.Connection, tabel: str) -> Schema:
    with closing(conn.cursor()) as cursor:
        # Catalogusfuncties gebruiken patronen ('_' is een jokerteken): daarom exact filteren op naam
        kolommen = [Kolom(r.column_name, (r.type_name or "").upper(), bool(r.nullable), r.column_size)
                    for r in cursor.columns(table=tabel) if r.table_name == tabel]
    indexen: dict[str, list[tuple[int, str]]] = {}
    with closing(conn.cursor()) as cursor:
        for s in cursor.statistics(tabel, unique=True):
            if s.table_name == tabel and s.index_name and s.column_name and not s.non_unique:
                indexen.setdefault(s.index_name, []).append((s.ordinal_position or 0, s.column_name))
    return Schema(tabel, kolommen, kies_sleutel(indexen, kolommen))


# ---------------------------------------------------------
# Omzetten van bestandswaarden naar Access-typen
# ---------------------------------------------------------
def _getaltekst(tekst: str) -> str:
    """'1.234,5' / '1,234.5' / '1234,5' -> '1234.5' (het laatste scheidingsteken is het decimaalteken)."""
    s = tekst.strip().replace(" ", "").replace("\u00a0", "").replace("\u202f", "").replace("'", "")
    if "," in s and "." in s:
        decimaal = "," if s.rfind(",") > s.rfind(".") else "."
        duizend = "." if decimaal == "," else ","
        s = s.replace(duizend, "").replace(decimaal, ".")
    elif s.count(",") == 1:
        s = s.replace(",", ".")
    elif s.count(",") > 1 or s.count(".") > 1:
        s = s.replace(",", "").replace(".", "")
    return s


def _naar_decimal(waarde: Any) -> Decimal:
    if isinstance(waarde, bool):
        return Decimal(int(waarde))
    if isinstance(waarde, Decimal):
        return waarde
    if isinstance(waarde, int):
        return Decimal(waarde)
    if isinstance(waarde, float):
        return Decimal(repr(waarde))
    try:
        getal = Decimal(_getaltekst(str(waarde)))
    except InvalidOperation:
        raise OmzetFout(f"'{waarde}' is geen getal") from None
    if not getal.is_finite():
        raise OmzetFout(f"'{waarde}' is geen getal")
    return getal


def _naar_geheel(waarde: Any) -> int:
    if isinstance(waarde, int):
        return int(waarde)
    getal = _naar_decimal(waarde)
    if getal != getal.to_integral_value():
        raise OmzetFout(f"'{waarde}' is geen geheel getal")
    return int(getal)


def _naar_datum(waarde: Any) -> datetime:
    if isinstance(waarde, datetime):
        return waarde.replace(tzinfo=None)
    if isinstance(waarde, date):
        return datetime(waarde.year, waarde.month, waarde.day)
    tekst = str(waarde).strip()
    try:
        return datetime.fromisoformat(tekst).replace(tzinfo=None)
    except ValueError:
        pass
    for formaat in [*DATUM_FORMATEN, "%H:%M", "%H:%M:%S"]:
        try:
            resultaat = datetime.strptime(tekst, formaat)
        except ValueError:
            continue
        # Alleen een tijd: Access gebruikt 30-12-1899 als datum
        return resultaat.replace(1899, 12, 30) if formaat.startswith("%H") else resultaat
    raise OmzetFout(f"'{waarde}' is geen geldige datum")


def _naar_janee(waarde: Any) -> bool:
    if isinstance(waarde, bool):
        return waarde
    if isinstance(waarde, (int, float, Decimal)):
        return waarde != 0
    tekst = str(waarde).strip().casefold()
    if tekst in WAAR:
        return True
    if tekst in ONWAAR:
        return False
    raise OmzetFout(f"'{waarde}' is geen Ja/Nee-waarde")


def _naar_tekst(waarde: Any, kolom: Kolom) -> str:
    if isinstance(waarde, str):
        tekst = waarde
    elif isinstance(waarde, float) and waarde.is_integer():
        tekst = str(int(waarde))
    elif isinstance(waarde, datetime):
        tekst = waarde.isoformat(" ")
    else:
        tekst = str(waarde)
    if kolom.type not in ("LONGCHAR", "MEMO", "WLONGVARCHAR") and kolom.grootte and len(tekst) > kolom.grootte:
        raise OmzetFout(f"tekst is langer dan {kolom.grootte} tekens")
    return tekst


def is_leeg(waarde: Any) -> bool:
    if waarde is None:
        return True
    if isinstance(waarde, float) and math.isnan(waarde):
        return True
    return isinstance(waarde, str) and waarde.strip() == ""


def zet_om(waarde: Any, kolom: Kolom) -> Any:
    """Zet een waarde uit een bestand of de editor om naar het type van de Access-kolom.
    Lege waarden worden NULL (Ja/Nee-velden: Nee)."""
    if is_leeg(waarde):
        return False if kolom.type in JANEE_TYPES else None
    if kolom.type in TEKST_TYPES:
        return _naar_tekst(waarde, kolom)
    if kolom.type in GEHEEL_TYPES:
        return _naar_geheel(waarde)
    if kolom.type in DRIJVEND_TYPES:
        return float(_naar_decimal(waarde))
    if kolom.type in DECIMAAL_TYPES:
        return _naar_decimal(waarde)
    if kolom.type in DATUM_TYPES:
        return _naar_datum(waarde)
    if kolom.type in JANEE_TYPES:
        return _naar_janee(waarde)
    return waarde


def gelijk(a: Any, b: Any) -> bool:
    """Vergelijkt een Access-waarde met een omgezette bestandswaarde (lege tekst = NULL, getallen numeriek)."""
    if is_leeg(a) and is_leeg(b):
        return True
    if is_leeg(a) or is_leeg(b):
        return False
    getallen = (int, float, Decimal)
    if isinstance(a, getallen) and isinstance(b, getallen) and not isinstance(a, bool) and not isinstance(b, bool):
        return math.isclose(float(a), float(b), rel_tol=1e-9, abs_tol=1e-9)
    if isinstance(a, datetime) and isinstance(b, datetime):
        return a.replace(microsecond=0, tzinfo=None) == b.replace(microsecond=0, tzinfo=None)
    return a == b


def _sleutelwaarde(waarde: Any) -> Any:
    if isinstance(waarde, bool):
        return int(waarde)
    if isinstance(waarde, (float, Decimal)):
        getal = float(waarde)
        return int(getal) if getal.is_integer() else getal
    if isinstance(waarde, str):
        return waarde.strip().casefold()  # Access vergelijkt tekst niet hoofdlettergevoelig
    if isinstance(waarde, datetime):
        return waarde.replace(microsecond=0)
    return waarde


# ---------------------------------------------------------
# Bronnen (bestanden/bladen) en koppeling aan Access-tabellen
# ---------------------------------------------------------
@dataclass
class Bron:
    label: str              # Weergave, bijv. 'Export.xlsx › Klanten_1'
    naam: str               # Blad- of bestandsnaam (zonder extensie)
    kolommen: list[str]
    rijen: list[list[Any]]
    tabel_hint: str | None = None  # Tabelnaam uit het exportmanifest
    pad: str = ""


def bronnen_uit_bestand(pad: str) -> list[Bron]:
    manifest = zoek_manifest(pad)
    bestandsnaam = os.path.basename(pad)
    hints: dict[str, str] = {}
    if manifest:
        for item in manifest["tables"]:
            tabel = str(item.get("table", ""))
            for blad in item.get("sheets", []) or []:
                hints[str(blad).casefold()] = tabel
            if item.get("file"):
                hints[str(item["file"]).casefold()] = tabel
    bronnen = []
    for naam, (kolommen, rijen) in lees_export(pad).items():
        is_csv = pad.lower().endswith(".csv")
        hint = hints.get(bestandsnaam.casefold() if is_csv else naam.casefold())
        label = bestandsnaam if is_csv else f"{bestandsnaam} › {naam}"
        bronnen.append(Bron(label, naam, kolommen, rijen, hint, pad))
    return bronnen


def _vergelijkbaar(naam: str) -> str:
    return naam.strip().casefold()


def raad_tabel(bron: Bron, tabellen: list[str]) -> str | None:
    """Zoekt de Access-tabel bij een bron: manifest, dan (opgeschoonde) naam, dan CSV-voorvoegsel."""
    per_naam = {_vergelijkbaar(t): t for t in tabellen}
    if bron.tabel_hint and _vergelijkbaar(bron.tabel_hint) in per_naam:
        return per_naam[_vergelijkbaar(bron.tabel_hint)]
    kandidaten = [bron.naam, re.sub(r"(~\d+)?_\d+$", "", bron.naam), re.sub(r"~\d+$", "", bron.naam)]
    for kandidaat in map(_vergelijkbaar, kandidaten):
        for tabel in tabellen:
            varianten = {
                _vergelijkbaar(tabel),
                _vergelijkbaar(_ONGELDIGE_BLADTEKENS.sub("_", tabel).strip("'")[:EXCEL_MAX_BLADNAAM]),
                _vergelijkbaar(veilige_bestandsnaam(tabel)),
            }
            if kandidaat in varianten:
                return tabel
    # CSV-export: '<database>_<tabel>' (eventueel met '_2' bij dubbele namen)
    naam = _vergelijkbaar(re.sub(r"_\d+$", "", bron.naam))
    passend = [t for t in tabellen if veilige_bestandsnaam(t) and naam.endswith("_" + _vergelijkbaar(veilige_bestandsnaam(t)))]
    return max(passend, key=len) if passend else None


def voeg_bronnen_samen(bronnen: list[Bron]) -> Tabel:
    """Voegt bronnen voor dezelfde tabel samen (bijv. opgesplitste Excel-bladen), uitgelijnd op kolomnaam."""
    kolommen: list[str] = []
    for bron in bronnen:
        for kolom in bron.kolommen:
            if kolom.casefold() not in {k.casefold() for k in kolommen}:
                kolommen.append(kolom)
    positie = {k.casefold(): i for i, k in enumerate(kolommen)}
    rijen: list[list[Any]] = []
    for bron in bronnen:
        doel = [positie[k.casefold()] for k in bron.kolommen]
        for rij in bron.rijen:
            nieuw: list[Any] = [None] * len(kolommen)
            for i, waarde in zip(doel, rij):
                nieuw[i] = waarde
            rijen.append(nieuw)
    return kolommen, rijen


# ---------------------------------------------------------
# Plan maken
# ---------------------------------------------------------
def maak_plan(schema: Schema, bron: Tabel, access: tuple[list[str], list[tuple[Any, ...]]],
              modus: str = MODUS_BIJWERKEN, verwijderen: bool = False, eerste_rijnummer: int = 2) -> TabelPlan:
    """Vergelijkt bronrijen met de Access-rijen en bepaalt wat er moet gebeuren.
    Rijen met fouten worden overgeslagen en gemeld; er wordt dan niets verwijderd."""
    plan = TabelPlan(schema.tabel, modus, list(schema.sleutel))
    bron_kolommen, bron_rijen = bron
    gekoppeld: list[tuple[int, Kolom]] = []
    gebruikt: set[str] = set()
    for i, naam in enumerate(bron_kolommen):
        kolom = schema.kolom(str(naam))
        if kolom is None:
            plan.opmerkingen.append(f"Kolom '{naam}' bestaat niet in Access en wordt overgeslagen")
        elif not kolom.ondersteund:
            plan.opmerkingen.append(f"Kolom '{naam}' ({kolom.type}) wordt niet ondersteund en overgeslagen")
        elif kolom.naam.casefold() not in gebruikt:
            gebruikt.add(kolom.naam.casefold())
            gekoppeld.append((i, kolom))
    for kolom in schema.kolommen:
        if kolom.ondersteund and kolom.naam.casefold() not in gebruikt:
            plan.opmerkingen.append(f"Kolom '{kolom.naam}' ontbreekt in het bestand en blijft ongewijzigd")

    sleutel = schema.sleutel if modus != MODUS_VERVANGEN else []
    if modus == MODUS_BIJWERKEN and not schema.sleutel:
        plan.geblokkeerd = True
        plan.fouten.append(f"Tabel '{schema.tabel}' heeft geen primaire sleutel; kies 'Alleen nieuwe rijen toevoegen' of 'Tabelinhoud vervangen'")
        return plan
    ontbrekend = [k for k in sleutel if k.casefold() not in gebruikt]
    autonummer_sleutel = len(sleutel) == 1 and bool((k := schema.kolom(sleutel[0])) and k.autonummer)
    if ontbrekend and not (modus == MODUS_TOEVOEGEN and autonummer_sleutel):
        plan.geblokkeerd = True
        plan.fouten.append(f"Sleutelkolom '{ontbrekend[0]}' ontbreekt in het bestand")
        return plan

    access_kolommen, access_rijen = access
    access_pos = {k.casefold(): i for i, k in enumerate(access_kolommen)}
    access_per_sleutel: dict[tuple[Any, ...], tuple[Any, ...]] = {}
    if sleutel:
        for rij in access_rijen:
            access_per_sleutel[tuple(_sleutelwaarde(rij[access_pos[k.casefold()]]) for k in sleutel)] = rij
    sleutelset = {k.casefold() for k in sleutel}
    gezien: set[tuple[Any, ...]] = set()

    def zonder_leeg_autonummer(waarden: dict[str, Any]) -> dict[str, Any]:
        return {k: v for k, v in waarden.items() if not (v is None and (kol := schema.kolom(k)) and kol.autonummer)}

    for n, rij in enumerate(bron_rijen):
        rijnummer = n + eerste_rijnummer
        if all(is_leeg(w) for w in rij):
            continue
        waarden: dict[str, Any] = {}
        fout = False
        for i, kolom in gekoppeld:
            ruw = rij[i] if i < len(rij) else None
            try:
                waarden[kolom.naam] = zet_om(ruw, kolom)
            except OmzetFout as e:
                plan.fouten.append(f"Rij {rijnummer}, kolom '{kolom.naam}': {e}")
                fout = True
        if fout:
            continue
        if not sleutel:
            plan.invoegen.append(zonder_leeg_autonummer(waarden))
            continue
        sleutelwaarden = [waarden.get(k) for k in sleutel]
        if any(v is None for v in sleutelwaarden):
            if autonummer_sleutel:
                plan.invoegen.append(zonder_leeg_autonummer(waarden))
            else:
                plan.fouten.append(f"Rij {rijnummer}: sleutelwaarde ontbreekt")
            continue
        sleutelwaarde = tuple(_sleutelwaarde(v) for v in sleutelwaarden)
        if sleutelwaarde in gezien:
            plan.fouten.append(f"Rij {rijnummer}: dubbele sleutelwaarde {', '.join(map(str, sleutelwaarden))}")
            continue
        gezien.add(sleutelwaarde)
        bestaand = access_per_sleutel.get(sleutelwaarde)
        if bestaand is None:
            plan.invoegen.append(waarden)
        elif modus == MODUS_TOEVOEGEN:
            plan.ongewijzigd += 1
        else:
            gewijzigd = {k: v for k, v in waarden.items()
                         if k.casefold() not in sleutelset and not gelijk(bestaand[access_pos[k.casefold()]], v)}
            if gewijzigd:
                plan.wijzigen.append((tuple(bestaand[access_pos[k.casefold()]] for k in sleutel), gewijzigd))
            else:
                plan.ongewijzigd += 1

    if modus == MODUS_VERVANGEN:
        plan.alles_verwijderen = True
    elif modus == MODUS_BIJWERKEN and verwijderen:
        if plan.fouten:
            plan.opmerkingen.append("Er worden geen rijen verwijderd omdat er fouten zijn")
        else:
            plan.verwijderen = [tuple(rij[access_pos[k.casefold()]] for k in sleutel)
                                for s, rij in access_per_sleutel.items() if s not in gezien]
    return plan


def bereken_plannen(conn: pyodbc.Connection, koppelingen: dict[str, list[Bron]], modus: str, verwijderen: bool,
                    stop: Callable[[], bool] = lambda: False) -> list[TabelPlan]:
    plannen = []
    for tabel, bronnen in koppelingen.items():
        if stop():
            raise Gestopt()
        try:
            schema = lees_schema(conn, tabel)
            plannen.append(maak_plan(schema, voeg_bronnen_samen(bronnen), lees_rijen(conn, tabel), modus, verwijderen))
        except pyodbc.Error as e:
            plannen.append(TabelPlan(tabel, modus, fouten=[str(e).split("\n")[0]], geblokkeerd=True))
    return plannen


# ---------------------------------------------------------
# Uitvoeren
# ---------------------------------------------------------
def _waar(sleutel: list[str]) -> str:
    return " AND ".join(f"{quote_naam(k)} = ?" for k in sleutel)


def voer_uit(conn: pyodbc.Connection, plannen: list[TabelPlan], stop: Callable[[], bool] = lambda: False,
             voortgang: Callable[[int, int], None] = lambda klaar, totaal: None) -> None:
    """Voert alle plannen uit in één transactie; bij een fout of stop wordt alles teruggedraaid."""
    uit_te_voeren = [p for p in plannen if not p.geblokkeerd]
    totaal = sum(p.aantal_bewerkingen for p in uit_te_voeren)
    klaar = 0
    cursor = conn.cursor()

    def stap() -> None:
        nonlocal klaar
        klaar += 1
        if klaar % 100 == 0 or klaar == totaal:
            if stop():
                raise Gestopt()
            voortgang(klaar, totaal)

    try:
        for plan in uit_te_voeren:
            tabel = quote_naam(plan.tabel)
            if plan.alles_verwijderen:
                cursor.execute(f"DELETE FROM {tabel}")
                stap()
            for sleutelwaarden in plan.verwijderen:
                cursor.execute(f"DELETE FROM {tabel} WHERE {_waar(plan.sleutel)}", *sleutelwaarden)
                stap()
            for sleutelwaarden, waarden in plan.wijzigen:
                zet = ", ".join(f"{quote_naam(k)} = ?" for k in waarden)
                cursor.execute(f"UPDATE {tabel} SET {zet} WHERE {_waar(plan.sleutel)}", *waarden.values(), *sleutelwaarden)
                stap()
            for waarden in plan.invoegen:
                kolommen = ", ".join(quote_naam(k) for k in waarden)
                plaatsen = ", ".join("?" for _ in waarden)
                cursor.execute(f"INSERT INTO {tabel} ({kolommen}) VALUES ({plaatsen})", *waarden.values())
                stap()
        conn.commit()
    except BaseException:
        conn.rollback()
        raise
    finally:
        cursor.close()


def zet_terug_in_kopie(origineel: str, doel: str, plannen: list[TabelPlan], wachtwoord: str = "",
                       stop: Callable[[], bool] = lambda: False,
                       voortgang: Callable[[int, int], None] = lambda klaar, totaal: None) -> None:
    """Kopieert de originele database naar `doel` en voert de plannen uit op die kopie.
    Mislukt het, dan wordt de (half bijgewerkte) kopie verwijderd. Het origineel wordt nooit gewijzigd."""
    maak_kopie(origineel, doel)
    try:
        with closing(verbind_access(doel, wachtwoord, alleen_lezen=False)) as conn:
            voer_uit(conn, plannen, stop, voortgang)
    except BaseException:
        try:
            os.remove(doel)
        except OSError:
            pass
        raise
