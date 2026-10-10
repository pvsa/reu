"""Stufe 2: Auslagen aus der zentralen <user>-auslagen.ods.

Liest das Blatt 'Auslagen' der EINEN Auslagen-Datei je User
([auslagen] datei, z.B. conf/alice-auslagen.ods). Die Zuordnung zur
Abrechnung erfolgt über die Spalte 'datum': Nur Zeilen, deren Datum im
Abrechnungszeitraum liegt, werden berücksichtigt – Zeilen anderer
Monate/Jahre bleiben unberührt in der Datei liegen.

Freigabe zeilenweise über die Spalte 'freigabe' ('yes'/'ja'): Nur
freigegebene Zeilen im Zeitraum werden abgerechnet; nicht freigegebene
werden übersprungen und in 'nicht_freigegeben' gemeldet. Die Rechnung
gilt als freigegeben (finale Re-Nr, ZUGFeRD, Versand), sobald MINDESTENS
EINE Zeile des Zeitraums freigegeben ist.

Spaltentitel (Header) sind groß-/kleinschreibungsagnostisch – 'Freigabe'
wie 'freigabe', 'Datum' wie 'datum' (wie beim Services-Blatt).

Existiert die Auslagen-Datei gar nicht, liefert lade_auslagen ein
leeres Ergebnis ('datei_vorhanden': False) – das CLI fragt dann zu
Beginn um Bestätigung und weist ansonsten generell auf die
Auslagenlage hin (siehe run-reu.py).

Rückgabe von lade_auslagen:
    {
      "datei_vorhanden": bool,  # False: [auslagen] datei existiert nicht
      "freigegeben": bool,      # True, sobald eine Zeitraum-Zeile yes/ja hat
      "auslagen": {kunde: [ {datum, art, bezeichnung, betrag_netto, belegnr} ]},
      "nicht_freigegeben": [ {kunde, bezeichnung, belegnr} | {hinweis} ],
      "offene_zeilen": {kunde: anzahl},  # nicht freigegebene Zeilen je Kunde
      "zeitraum_zeilen": int,   # Zeilen im Zeitraum (abgerechnet + übersprungen)
    }
"""
from __future__ import annotations

import datetime as _dt
from decimal import Decimal
from pathlib import Path
from typing import Any

from .config import AuslagenConf
from .ods import blatt_als_dicts
from .util import Zeitraum, dezimal, letzter_des_monats

PFLICHTSPALTEN_DEFAULT = ["datum", "kunde", "art", "bezeichnung", "betrag_netto", "belegnr"]

_DATUM_FORMATE = ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y")


class AuslagenError(Exception):
    pass


def _als_datum(wert: object) -> _dt.date:
    """Parst ein Datum: date/datetime (ODS-Datumzelle), ISO oder TT.MM.JJJJ."""
    if isinstance(wert, _dt.datetime):
        return wert.date()
    if isinstance(wert, _dt.date):
        return wert
    text = str(wert or "").strip()
    for fmt in _DATUM_FORMATE:
        try:
            return _dt.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise AuslagenError(
        f"Auslagen-Datum unlesbar: {wert!r} "
        "(erwartet JJJJ-MM-TT oder TT.MM.JJJJ)"
    )


def _leere_zeile(satz: dict[str, Any]) -> bool:
    return all(v is None or not str(v).strip() for v in satz.values())


def _norm_satz(satz: dict[str, Any]) -> dict[str, Any]:
    """Normalisiert Header-Schlüssel: 'Freigabe' -> 'freigabe' etc.

    Spaltentitel sind damit groß-/kleinschreibungsagnostisch – wie beim
    Services-Blatt ('Kunde' wie 'kunde'). Leerzeichen am Rand entfallen.
    """
    aus: dict[str, Any] = {}
    for schluessel, wert in satz.items():
        k = str(schluessel).strip().lower()
        if k and k not in aus:
            aus[k] = wert
    return aus


def auslagen_datei_vorhanden(conf: AuslagenConf) -> bool:
    """True, wenn die Auslagen-ODS ([auslagen] datei) existiert."""
    return Path(conf.datei).is_file()


def lade_auslagen(
    conf: AuslagenConf, zeitraum: Zeitraum, bekannte_kunden: set[str]
) -> dict[str, Any]:
    """Lädt die zentrale Auslagen-ODS, filtert nach Zeitraum und Freigabe.

    Die Spalte 'datum' entscheidet über die Zuordnung zum Zeitraum:
    Nur Zeilen mit Datum innerhalb des Abrechnungszeitraums
    (erster Tag des ersten Monats bis letzter Tag des letzten Monats)
    werden berücksichtigt.
    """
    pfad = Path(conf.datei)
    leer = {
        "datei_vorhanden": False,
        "freigegeben": False,
        "auslagen": {},
        "nicht_freigegeben": [],
        "offene_zeilen": {},
        "zeitraum_zeilen": 0,
    }
    if not pfad.is_file():
        return leer
    # Header normalisieren ('Freigabe' == 'freigabe', 'Datum' == 'datum')
    saetze = [_norm_satz(s) for s in blatt_als_dicts(pfad, conf.blatt)]
    pflichtspalten = [str(p).strip().lower() for p in (conf.pflichtspalten or PFLICHTSPALTEN_DEFAULT)]
    if saetze:
        vorhandene = set(saetze[0].keys())
        fehlt = [s for s in pflichtspalten if s not in vorhandene]
        if fehlt:
            raise AuslagenError(
                f"Pflichtspalten fehlen in {pfad}!{conf.blatt}: {', '.join(fehlt)} "
                f"(Groß-/Kleinschreibung der Spaltentitel ist egal; vorhanden: "
                f"{', '.join(sorted(vorhandene))})"
            )
    spalte = str(conf.freigabe_spalte).strip().lower()
    hat_spalte = bool(saetze) and spalte in saetze[0]
    von = _dt.date(zeitraum.jahr, zeitraum.erster_monat, 1)
    bis = letzter_des_monats(zeitraum.jahr, zeitraum.letzter_monat)

    nicht_freigegeben: list[dict[str, Any]] = []
    offene_zeilen: dict[str, int] = {}
    gruppe: dict[str, list[dict[str, Any]]] = {}
    freigegeben = False
    zeitraum_zeilen = 0
    for satz in saetze:
        if _leere_zeile(satz):
            continue
        datum = _als_datum(satz.get("datum"))
        if not (von <= datum <= bis):
            continue  # anderer Monat/Jahr – bleibt unberührt
        zeitraum_zeilen += 1
        kunde = str(satz.get("kunde", "")).strip()
        if not kunde:
            raise AuslagenError(
                f"Auslagenzeile ohne Kunde in {pfad}!{conf.blatt}: {satz!r}"
            )
        if kunde not in bekannte_kunden:
            raise AuslagenError(
                f"Auslagenzeile für unbekannten Kunden '{kunde}' – "
                f"in kunden.ods nicht definiert"
            )
        roh_betrag = satz.get("betrag_netto")
        betrag = dezimal(roh_betrag)
        if (
            betrag == 0
            and str(roh_betrag or "").strip() not in {"", "0", "0,00", "0.00"}
        ):
            raise AuslagenError(
                f"betrag_netto für Beleg {satz.get('belegnr', '?')} (Kunde {kunde}) "
                f"ist nicht numerisch: {roh_betrag!r}"
            )
        if (
            not hat_spalte
            or str(satz.get(spalte, "")).strip().lower() not in {"yes", "ja"}
        ):
            nicht_freigegeben.append(
                {
                    "kunde": kunde,
                    "bezeichnung": str(satz.get("bezeichnung", "")).strip(),
                    "belegnr": str(satz.get("belegnr", "")).strip(),
                }
            )
            offene_zeilen[kunde] = offene_zeilen.get(kunde, 0) + 1
            continue
        gruppe.setdefault(kunde, []).append(
            {
                "datum": datum.isoformat(),
                "art": str(satz.get("art", "")).strip(),
                "bezeichnung": str(satz.get("bezeichnung", "")).strip(),
                "betrag_netto": betrag,
                "belegnr": str(satz.get("belegnr", "")).strip(),
            }
        )
        freigegeben = True

    if not hat_spalte and zeitraum_zeilen:
        nicht_freigegeben.insert(
            0,
            {"hinweis": f"Spalte '{spalte}' fehlt – keine Zeile freigegeben"},
        )

    # je Kunde nach Datum sortieren
    for kunde in gruppe:
        gruppe[kunde].sort(key=lambda e: (e["datum"], e["belegnr"]))
    return {
        "datei_vorhanden": True,
        "freigegeben": freigegeben,
        "auslagen": gruppe,
        "nicht_freigegeben": nicht_freigegeben,
        "offene_zeilen": offene_zeilen,
        "zeitraum_zeilen": zeitraum_zeilen,
    }


def auslagen_summe(zeilen: list[dict[str, Any]]) -> Decimal:
    return sum((z["betrag_netto"] for z in zeilen), Decimal("0"))
