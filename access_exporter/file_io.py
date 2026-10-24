"""Lezen en schrijven van geëxporteerde CSV- en Excel-bestanden (zonder Qt)."""

import math
import os
import tempfile
from typing import Any

import pandas as pd

from .core import CSV_ENCODING, CSV_SCHEIDINGSTEKEN, EXCEL_ENGINE, controleer_schrijfbaar, maak_bladnaam

Tabel = tuple[list[str], list[list[Any]]]  # (kolommen, rijen)


def _schoon(waarde: Any) -> Any:
    """NaN/NaT/pd.NA -> None, pandas-tijdstempels -> datetime."""
    if waarde is None:
        return None
    if isinstance(waarde, float) and math.isnan(waarde):
        return None
    if waarde is pd.NaT or waarde is pd.NA:
        return None
    if isinstance(waarde, pd.Timestamp):
        return waarde.to_pydatetime()
    return waarde


def _frame_naar_tabel(df: pd.DataFrame) -> Tabel:
    kolommen = [str(k) for k in df.columns]
    rijen = [[_schoon(w) for w in rij] for rij in df.itertuples(index=False, name=None)]
    return kolommen, rijen


def lees_export(pad: str) -> dict[str, Tabel]:
    """Leest een CSV (één tabel, naam = bestandsnaam) of Excel-bestand (één tabel per blad).
    CSV-waarden blijven tekst (voorloopnullen blijven behouden); lege cellen worden None."""
    if pad.lower().endswith(".csv"):
        df = pd.read_csv(pad, sep=CSV_SCHEIDINGSTEKEN, encoding=CSV_ENCODING, dtype=str,
                         keep_default_na=False, na_filter=False)
        kolommen, rijen = _frame_naar_tabel(df)
        rijen = [[None if w == "" else w for w in rij] for rij in rijen]
        return {os.path.splitext(os.path.basename(pad))[0]: (kolommen, rijen)}
    if pad.lower().endswith((".xlsx", ".xlsm")):
        bladen = pd.read_excel(pad, sheet_name=None, dtype=object, engine="openpyxl")
        return {str(naam): _frame_naar_tabel(df) for naam, df in bladen.items()}
    raise ValueError(f"Niet-ondersteund bestandstype: {os.path.basename(pad)}")


def _atomair_schrijven(pad: str, schrijf) -> None:
    """Schrijft eerst naar een tijdelijk bestand in dezelfde map en vervangt dan het doel.
    Bij een fout blijft een bestaand doelbestand ongewijzigd."""
    controleer_schrijfbaar(pad)
    map_ = os.path.dirname(os.path.abspath(pad))
    fd, tijdelijk = tempfile.mkstemp(prefix=".~", suffix=os.path.splitext(pad)[1], dir=map_)
    os.close(fd)
    try:
        schrijf(tijdelijk)
        os.replace(tijdelijk, pad)
    except BaseException:
        if os.path.exists(tijdelijk):
            os.remove(tijdelijk)
        raise


def schrijf_csv(pad: str, tabel: Tabel) -> None:
    kolommen, rijen = tabel
    df = pd.DataFrame(rijen, columns=kolommen, dtype=object)
    _atomair_schrijven(pad, lambda doel: df.to_csv(doel, index=False, sep=CSV_SCHEIDINGSTEKEN,
                                                    encoding=CSV_ENCODING, lineterminator="\r\n"))


def schrijf_excel(pad: str, tabellen: dict[str, Tabel]) -> None:
    if EXCEL_ENGINE is None:
        raise RuntimeError("Geen Excel-bibliotheek gevonden. Installeer met: pip install xlsxwriter")
    # Tekst als '=...' of 'http://...' moet als tekst blijven, niet als formule/hyperlink
    engine_kwargs = {"options": {"strings_to_formulas": False, "strings_to_urls": False}} if EXCEL_ENGINE == "xlsxwriter" else {}

    def schrijf(doel: str) -> None:
        gebruikt: set[str] = set()
        with pd.ExcelWriter(doel, engine=EXCEL_ENGINE, engine_kwargs=engine_kwargs) as writer:
            for naam, (kolommen, rijen) in tabellen.items():
                pd.DataFrame(rijen, columns=kolommen, dtype=object).to_excel(
                    writer, sheet_name=maak_bladnaam(naam, gebruikt), index=False)

    _atomair_schrijven(pad, schrijf)
