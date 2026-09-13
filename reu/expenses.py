"""Stufe 2: Auslagen aus der monatlichen ODS-Datei.

Liest das Blatt 'Auslagen' der datei (Pfad aus [auslagen] mit {year}/{month}
ersetzt), prüft Pflichtspalten und Freigabe-Signal (Meta!A1 == 'yes') und
gruppiert die Belege nach Kunde.

Rückgabe:
    {
      "freigegeben": bool,
      "auslagen": {kunde: [ {datum, art, bezeichnung, betrag_netto, belegnr} ]},
    }
"""
from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from typing import Any

from .config import AuslagenConf
from .ods import blatt_als_dicts, zelle
from .util import date_to_iso, dezimal

PFLICHTSPALTEN_DEFAULT = ["datum", "kunde", "art", "bezeichnung", "betrag_netto", "belegnr"]


class AuslagenError(Exception):
    pass


def _aufgeloester_pfad(datei_template: str, year: int, month: int) -> str:
    return datei_template.format(year=year, month=month)


def lade_auslagen(
    conf: AuslagenConf, year: int, month: int, bekannte_kunden: set[str]
) -> dict[str, Any]:
    pfad = _aufgeloester_pfad(conf.datei, year, month)
    p = Path(pfad)
    if not p.is_absolute() and not p.exists():
        # relativer Pfad: relativ zur Basis (Konf-Verzeichnis-Oberhalb)
        # Wir versuchen den Pfad wie angegeben; ansonsten Fehler.
        pass
    if not Path(pfad).is_file():
        raise AuslagenError(f"Auslagen-ODS nicht gefunden: {pfad}")

    saetze = blatt_als_dicts(pfad, conf.blatt)
    pflichtspalten = conf.pflichtspalten or PFLICHTSPALTEN_DEFAULT
    if saetze:
        vorhandene = set(saetze[0].keys())
        fehlt = [s for s in pflichtspalten if s not in vorhandene]
        if fehlt:
            raise AuslagenError(
                f"Pflichtspalten fehlen in {pfad}!{conf.blatt}: {', '.join(fehlt)}"
            )

    # Freigabe
    freigabe = str(zelle(pfad, "", conf.freigabe_zelle)).strip().lower() == "yes"
    if not freigabe:
        # trotzdem laden, damit der Entwurf die Werte anzeigen kann
        pass

    gruppe: dict[str, list[dict[str, Any]]] = {}
    for satz in saetze:
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
        datum = satz.get("datum")
        gruppe.setdefault(kunde, []).append(
            {
                "datum": date_to_iso(datum) or "",
                "art": str(satz.get("art", "")).strip(),
                "bezeichnung": str(satz.get("bezeichnung", "")).strip(),
                "betrag_netto": betrag,
                "belegnr": str(satz.get("belegnr", "")).strip(),
            }
        )

    # je Kunde nach Datum sortieren
    for kunde in gruppe:
        gruppe[kunde].sort(key=lambda e: (e["datum"], e["belegnr"]))
    return {"freigegeben": freigabe, "auslagen": gruppe}


def auslagen_summe(zeilen: list[dict[str, Any]]) -> Decimal:
    return sum((z["betrag_netto"] for z in zeilen), Decimal("0"))
