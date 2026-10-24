"""Achtergrondtaken (QThread): de export-conversie en een algemene taakuitvoerder."""

import math
import os
from collections.abc import Callable
from contextlib import closing
from typing import Any

import pandas as pd
import pyodbc
from PyQt6.QtCore import QThread, pyqtSignal

from .core import (
    EXCEL_ENGINE, EXCEL_MAX_DATA_RIJEN, STATUS_BEZIG, STATUS_FOUT, STATUS_GESTOPT, STATUS_KLAAR,
    controleer_schrijfbaar, lees_in_blokken, maak_bladnaam, maak_csv_naam, manifest_pad, open_in_verkenner,
    schrijf_manifest, tel_rijen, verbind_access,
)


class ConversieGestopt(Exception):
    pass


class ConversionWorker(QThread):
    """Voert de zware conversies uit op de achtergrond en communiceert via signalen met de GUI."""
    update_table_signal = pyqtSignal(int, str, str, str)  # rij, voortgang, status, melding
    log_signal = pyqtSignal(str)
    job_done_signal = pyqtSignal()

    def __init__(self, taken: list[tuple[int, dict[str, Any]]], open_map: bool):
        super().__init__()
        self.taken = taken
        self.open_map = open_map
        self._rij = 0
        self._rijen_totaal = 0
        self._rijen_klaar = 0
        self._laatste_pct = -1
        self._max_pct = 100

    def run(self) -> None:
        for rij, job in self.taken:
            if self.isInterruptionRequested():
                self.log_signal.emit("⏹ Conversie gestopt; resterende taken blijven in de wachtrij.")
                break
            self._verwerk_job(rij, job)
            self.job_done_signal.emit()

    def _controleer_stop(self) -> None:
        if self.isInterruptionRequested():
            raise ConversieGestopt()

    def _voortgang(self, extra_rijen: int, melding: str) -> None:
        self._rijen_klaar += extra_rijen
        pct = self._max_pct if not self._rijen_totaal else int(self._rijen_klaar / self._rijen_totaal * self._max_pct)
        if pct != self._laatste_pct:
            self._laatste_pct = pct
            self.update_table_signal.emit(self._rij, f"{pct}%", STATUS_BEZIG, melding)

    def _verwerk_job(self, rij: int, job: dict[str, Any]) -> None:
        naam = job["access_naam"]
        export_type = job["export_type"]
        tabellen: list[str] = job["tabellen"]
        aangemaakt: list[str] = []

        self._rij, self._rijen_klaar, self._laatste_pct = rij, 0, -1
        self.update_table_signal.emit(rij, "0%", STATUS_BEZIG, "Verbinding maken...")
        self.log_signal.emit(f"Start verwerking: {naam} (Formaat: {export_type.upper()})")

        try:
            with closing(verbind_access(job["access_pad"], job.get("wachtwoord", ""))) as conn:
                self.update_table_signal.emit(rij, "0%", STATUS_BEZIG, "Rijen tellen...")
                tellingen = {t: tel_rijen(conn, t) for t in tabellen}
                self._rijen_totaal = sum(tellingen.values())

                if export_type == "xlsx":
                    resultaat, inhoud = self._schrijf_excel(conn, job, tellingen, aangemaakt)
                elif export_type == "csv":
                    resultaat, inhoud = self._schrijf_csv(conn, job, aangemaakt)
                else:
                    raise ValueError(f"Onbekend exportformaat: {export_type}")
            # Manifest: maakt terugzetten naar Access betrouwbaar (welke tabel staat in welk blad/bestand)
            schrijf_manifest(manifest_pad(job["doelmap"], job["naam_zonder_ext"], export_type),
                             job["access_pad"], export_type, inhoud, job.get("csv_decimaal", ","))

        except ConversieGestopt:
            self._ruim_op(aangemaakt)
            self.update_table_signal.emit(rij, "0%", STATUS_GESTOPT, "Door gebruiker gestopt")
            self.log_signal.emit(f"⏹ Gestopt: {naam} (gedeeltelijke bestanden verwijderd)")
            return
        except Exception as e:
            self._ruim_op(aangemaakt)
            foutreden = str(e).split("\n")[0]
            self.update_table_signal.emit(rij, "0%", STATUS_FOUT, foutreden)
            self.log_signal.emit(f"❌ Fout bij {naam}: {foutreden}")
            return

        # Het aantal wordt bij het tonen opgemaakt volgens de gekozen taal (1.000 / 1,000)
        self.update_table_signal.emit(rij, "100%", STATUS_KLAAR, f"Succesvol ({self._rijen_totaal} rijen)")
        self.log_signal.emit(f"✅ Conversie succesvol afgerond: {naam}")
        if self.open_map:
            open_in_verkenner(resultaat)

    def _schrijf_excel(self, conn: pyodbc.Connection, job: dict[str, Any], tellingen: dict[str, int],
                       aangemaakt: list[str]) -> tuple[str, list[dict[str, Any]]]:
        if EXCEL_ENGINE is None:
            raise RuntimeError("Geen Excel-bibliotheek gevonden. Installeer met: pip install xlsxwriter")

        excel_pad = os.path.join(job["doelmap"], f"{job['naam_zonder_ext']}.xlsx")
        controleer_schrijfbaar(excel_pad)
        tabellen: list[str] = job["tabellen"]
        totaal = len(tabellen)
        gebruikte_bladen: set[str] = set()
        inhoud: list[dict[str, Any]] = []
        # Tekst als '=...' of 'http://...' moet als tekst blijven, niet als formule/hyperlink
        engine_kwargs = {"options": {"strings_to_formulas": False, "strings_to_urls": False}} if EXCEL_ENGINE == "xlsxwriter" else {}
        self._max_pct = 95  # Laatste stap is het opslaan van het bestand

        aangemaakt.append(excel_pad)
        with pd.ExcelWriter(excel_pad, engine=EXCEL_ENGINE, engine_kwargs=engine_kwargs) as writer:
            for index, tabel in enumerate(tabellen, 1):
                self._controleer_stop()
                self.log_signal.emit(f"-> Inlezen tabel '{tabel}' ({index}/{totaal})...")
                aantal_delen = max(1, math.ceil(tellingen[tabel] / EXCEL_MAX_DATA_RIJEN))
                if aantal_delen > 1:
                    self.log_signal.emit(f"   (Te groot, opsplitsen in {aantal_delen} delen...)")

                bladen: list[str] = []
                blad_idx = 0
                rijen_in_blad = 0
                for df in lees_in_blokken(conn, tabel):
                    self._controleer_stop()
                    if df.empty:  # Lege tabel: alleen de kopregel wegschrijven
                        bladen.append(maak_bladnaam(tabel, gebruikte_bladen))
                        df.to_excel(writer, sheet_name=bladen[-1], index=False)
                        continue
                    start = 0
                    while start < len(df):
                        if rijen_in_blad == EXCEL_MAX_DATA_RIJEN:
                            blad_idx += 1
                            rijen_in_blad = 0
                        if blad_idx == len(bladen):
                            splitsen = aantal_delen > 1 or blad_idx > 0
                            bladen.append(maak_bladnaam(tabel, gebruikte_bladen, f"_{blad_idx + 1}" if splitsen else ""))
                        deel = df.iloc[start:start + EXCEL_MAX_DATA_RIJEN - rijen_in_blad]
                        deel.to_excel(
                            writer, sheet_name=bladen[blad_idx], index=False,
                            header=rijen_in_blad == 0, startrow=rijen_in_blad + 1 if rijen_in_blad else 0
                        )
                        rijen_in_blad += len(deel)
                        start += len(deel)
                    self._voortgang(len(df), f"Tabel {index}/{totaal}: {tabel}")
                inhoud.append({"table": tabel, "sheets": bladen})
                self._voortgang(0, f"Tabel {index}/{totaal} ingelezen")

            self.update_table_signal.emit(self._rij, "95%", STATUS_BEZIG, "Excel-bestand opslaan...")
            self.log_signal.emit("-> Opslaan van Excel-bestand...")
        return excel_pad, inhoud

    def _schrijf_csv(self, conn: pyodbc.Connection, job: dict[str, Any], aangemaakt: list[str]) -> tuple[str, list[dict[str, Any]]]:
        tabellen: list[str] = job["tabellen"]
        totaal = len(tabellen)
        gebruikte_namen: set[str] = set()
        inhoud: list[dict[str, Any]] = []
        decimaal = job.get("csv_decimaal", ",")
        self._max_pct = 100

        for index, tabel in enumerate(tabellen, 1):
            self._controleer_stop()
            bestandsnaam = maak_csv_naam(job["naam_zonder_ext"], tabel, gebruikte_namen)
            csv_pad = os.path.join(job["doelmap"], bestandsnaam)
            controleer_schrijfbaar(csv_pad)
            self.log_signal.emit(f"-> Schrijven CSV: '{tabel}' ({index}/{totaal})...")
            with open(csv_pad, "w", encoding="utf-8-sig", newline="") as f:
                aangemaakt.append(csv_pad)
                for blok_nr, df in enumerate(lees_in_blokken(conn, tabel)):
                    self._controleer_stop()
                    df.to_csv(f, index=False, sep=";", header=blok_nr == 0, decimal=decimaal)
                    self._voortgang(len(df), f"CSV {index}/{totaal}: {tabel}")
            inhoud.append({"table": tabel, "file": bestandsnaam})
            self._voortgang(0, f"CSV {index}/{totaal} geschreven")
        return job["doelmap"], inhoud

    def _ruim_op(self, paden: list[str]) -> None:
        for pad in paden:
            try:
                os.remove(pad)
            except OSError:
                pass


class TaakWorker(QThread):
    """Voert één functie uit op de achtergrond. De functie krijgt `stop` (bool-functie) en
    `voortgang(klaar, totaal)` mee; het resultaat of de fout komt terug via signalen."""
    voortgang_signal = pyqtSignal(int, int)
    klaar_signal = pyqtSignal(object)
    fout_signal = pyqtSignal(object)

    def __init__(self, functie: Callable[..., Any]):
        super().__init__()
        self.functie = functie

    def run(self) -> None:
        try:
            resultaat = self.functie(stop=self.isInterruptionRequested,
                                     voortgang=lambda klaar, totaal: self.voortgang_signal.emit(klaar, totaal))
        except BaseException as e:  # noqa: BLE001 - elke fout moet terug naar de GUI
            self.fout_signal.emit(e)
        else:
            self.klaar_signal.emit(resultaat)
