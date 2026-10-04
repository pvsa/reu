"""Stufe 2: Auslagen aus den monatlichen ODS-Dateien.

Liest das Blatt 'Auslagen' je Monat des Zeitraums (Pfad aus [auslagen] mit
{year}/{month} ersetzt), prüft Pflichtspalten und Freigabe-Signal
(Meta!A1 == 'yes') und gruppiert die Belege nach Kunde.

Bei Mehrmonats-Zeiträumen (Monatsbereich/Quartal) werden die Monatsdateien
zusammengefasst. Fehlende Monatsdateien sind erlaubt. Existiert KEINE
Monatsdatei des Zeitraums, wird kein Fehler geworfen, sondern
"ohne_dateien": True geliefert – das CLI fragt dann den Nutzer um
Bestätigung, dass ohne Auslagen abgerechnet wird. Freigabe erfordert 'yes'
in ALLEN vorhandenen Monatsdateien.

Rückgabe:
    {
      "freigegeben": bool,
      "auslagen": {kunde: [ {datum, art, bezeichnung, betrag_netto, belegnr} ]},
      "ohne_dateien": bool,   # True: keine Auslagen-Datei im Zeitraum vorhanden
    }
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Any

from .config import AuslagenConf
from .ods import blatt_als_dicts, zelle
from .util import Zeitraum, date_to_iso, dezimal

PFLICHTSPALTEN_DEFAULT = ["datum", "kunde", "art", "bezeichnung", "betrag_netto", "belegnr"]


class AuslagenError(Exception):
    pass


def _aufgeloester_pfad(datei_template: str, year: int, month: int) -> str:
    return datei_template.format(year=year, month=month)


def _lies_monatsdatei(pfad: str, conf: AuslagenConf) -> tuple[list[dict[str, Any]], bool]:
    """Liest eine Monats-ODS, prüft Pflichtspalten, liefert (saetze, freigegeben)."""
    saetze = blatt_als_dicts(pfad, conf.blatt)
    pflichtspalten = conf.pflichtspalten or PFLICHTSPALTEN_DEFAULT
    if saetze:
        vorhandene = set(saetze[0].keys())
        fehlt = [s for s in pflichtspalten if s not in vorhandene]
        if fehlt:
            raise AuslagenError(
                f"Pflichtspalten fehlen in {pfad}!{conf.blatt}: {', '.join(fehlt)}"
            )
    freigabe = str(zelle(pfad, "", conf.freigabe_zelle)).strip().lower() == "yes"
    return saetze, freigabe


def auslagen_dateien_vorhanden(conf: AuslagenConf, zeitraum: Zeitraum) -> bool:
    """True, wenn mindestens eine Auslagen-Datei des Zeitraums existiert."""
    return any(
        Path(_aufgeloester_pfad(conf.datei, zeitraum.jahr, monat)).is_file()
        for monat in zeitraum.monate
    )


def lade_auslagen(
    conf: AuslagenConf, zeitraum: Zeitraum, bekannte_kunden: set[str]
) -> dict[str, Any]:
    """Lädt die Auslagen-ODS je Monat des Zeitraums und fasst sie zusammen."""
    saetze_alle: list[dict[str, Any]] = []
    freigegeben = True
    vorhandene_dateien = 0

    for monat in zeitraum.monate:
        pfad = _aufgeloester_pfad(conf.datei, zeitraum.jahr, monat)
        if not Path(pfad).is_file():
            continue  # Monat ohne Auslagen-Datei ist erlaubt
        vorhandene_dateien += 1
        saetze, freigabe = _lies_monatsdatei(pfad, conf)
        if not freigabe:
            freigegeben = False
        saetze_alle.extend(saetze)

    if vorhandene_dateien == 0:
        # Keine Auslagen-Datei im gesamten Zeitraum: kein Abbruch,
        # das CLI holt eine Nutzer-Bestätigung ein (siehe run-reu.py).
        return {"freigegeben": False, "auslagen": {}, "ohne_dateien": True}

    gruppe: dict[str, list[dict[str, Any]]] = {}
    for satz in saetze_alle:
        kunde = str(satz.get("kunde", "")).strip()
        if not kunde:
            continue
        if kunde not in bekannte_kunden:
            raise AuslagenError(
                f"Auslagenzeile für unbekannten Kunden '{kunde}' – "
                f"in kunden.ods nicht definiert"
            )
        betrag = dezimal(satz.get("betrag_netto"))
        if betrag == 0 and str(satz.get("betrag_netto", "")).strip() not in {"", "0", "0,00", "0.00"}:
            raise AuslagenError(
                f"betrag_netto für Beleg {satz.get('belegnr','?')} (Kunde {kunde}) "
                f"ist nicht numerisch: {satz.get('betrag_netto')!r}"
            )
        gruppe.setdefault(kunde, []).append(
            {
                "datum": date_to_iso(satz.get("datum")) or "",
                "art": str(satz.get("art", "")).strip(),
                "bezeichnung": str(satz.get("bezeichnung", "")).strip(),
                "betrag_netto": betrag,
                "belegnr": str(satz.get("belegnr", "")).strip(),
            }
        )

    # je Kunde nach Datum sortieren
    for kunde in gruppe:
        gruppe[kunde].sort(key=lambda e: (e["datum"], e["belegnr"]))
    return {"freigegeben": freigegeben, "auslagen": gruppe, "ohne_dateien": False}


def auslagen_summe(zeilen: list[dict[str, Any]]) -> Decimal:
    return sum((z["betrag_netto"] for z in zeilen), Decimal("0"))
