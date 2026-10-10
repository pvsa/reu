"""Stufe 2: Auslagen aus der zentralen <user>-auslagen.ods.

Liest das Blatt 'Auslagen' der EINEN Auslagen-Datei je User
([auslagen] datei, z.B. conf/alice-auslagen.ods). Die Zuordnung zur
Abrechnung läuft ALLEIN über die Spalte 'freigabe': Jede Zeile mit
'yes'/'ja' fließt in die nächste Rechnung – unabhängig vom
Abrechnungszeitraum. Die Spalte 'datum' ist reine Beleginformation
(wird angezeigt, aber nie als Filterkriterium genutzt).

Nicht freigegebene Zeilen (leer/'no'/'nein') werden übersprungen und in
'nicht_freigegeben' gemeldet. Die Rechnung gilt als freigegeben (finale
Re-Nr, ZUGFeRD, Versand), sobald MINDESTENS EINE Zeile der Datei
'yes'/'ja' hat.

Spaltentitel (Header) sind groß-/kleinschreibungsagnostisch – 'Freigabe'
wie 'freigabe', 'Datum' wie 'datum' (wie beim Services-Blatt).

Existiert die Auslagen-Datei gar nicht, liefert lade_auslagen ein
leeres Ergebnis ('datei_vorhanden': False) – das CLI fragt dann zu
Beginn um Bestätigung und weist ansonsten generell auf die
Auslagenlage hin (siehe run-reu.py).

Rückgabe von lade_auslagen:
    {
      "datei_vorhanden": bool,  # False: [auslagen] datei existiert nicht
      "freigegeben": bool,      # True, sobald IRGENDWO eine Zeile yes/ja hat
      "auslagen": {kunde: [ {datum, art, bezeichnung, betrag_netto, belegnr} ]},
      "nicht_freigegeben": [ {kunde, bezeichnung, belegnr} | {hinweis} ],
      "offene_zeilen": {kunde: anzahl},  # nicht freigegebene Zeilen je Kunde
      "erledigt": [ {kunde, bezeichnung, belegnr, vermerk} ],  # bereits abgerechnet (z.B. Re-Nr)
      "zeilen_gesamt": int,     # alle Zeilen der Datei (abgerechnet + übersprungen + erledigt)
    }


Erledigt-Vermerk: Nach einer finalen Rechnung schreibt
vermerke_rechnungsnummern() die Rechnungsnummer in die Freigabe-Spalte
ALLER abgerechneten Zeilen ('ja' -> '2026-001', datum-unabhängig). Diese
Zeilen gelten bei künftigen Läufen als bereits abgerechnet: Sie werden
nicht erneut abgerechnet und nicht als 'nicht freigegeben' angemeckert.
"""
from __future__ import annotations

import datetime as _dt
from decimal import Decimal
from pathlib import Path
from typing import Any

from .config import AuslagenConf
from .ods import OdsError, aendere_spalte, blatt_als_dicts
from .util import dezimal

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


def _als_datum_opt(wert: object) -> _dt.date | None:
    """Parst das Belegdatum tolerant; None bei leer/unlesbar.

    Das Datum ist nur Beleginformation (kein Filterkriterium) – eine
    leere oder unlesbare Zelle darf die Abrechnung nicht blockieren.
    """
    try:
        return _als_datum(wert)
    except AuslagenError:
        return None


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


def lade_auslagen(conf: AuslagenConf, bekannte_kunden: set[str]) -> dict[str, Any]:
    """Lädt die zentrale Auslagen-ODS; Zuordnung NUR über die Freigabe-Spalte.

    Jede Zeile mit freigabe='yes'/'ja' wird abgerechnet – egal welches
    Datum (oder welcher Monat) in der Spalte 'datum' steht.
    """
    pfad = Path(conf.datei)
    leer = {
        "datei_vorhanden": False,
        "freigegeben": False,
        "auslagen": {},
        "nicht_freigegeben": [],
        "offene_zeilen": {},
        "erledigt": [],
        "zeilen_gesamt": 0,
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

    nicht_freigegeben: list[dict[str, Any]] = []
    offene_zeilen: dict[str, int] = {}
    erledigt: list[dict[str, Any]] = []
    gruppe: dict[str, list[dict[str, Any]]] = {}
    freigegeben = False
    zeilen_gesamt = 0
    for satz in saetze:
        if _leere_zeile(satz):
            continue
        # Datum ist nur Beleginformation: wird gelesen (für Anzeige/
        # Sortierung), entscheidet aber NICHT über die Zuordnung.
        datum = _als_datum_opt(satz.get("datum"))
        zeilen_gesamt += 1
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
        wert = str(satz.get(spalte, "")).strip() if hat_spalte else ""
        wert_klein = wert.lower()
        if (
            hat_spalte
            and wert_klein not in {"yes", "ja", "no", "nein"}
            and wert != ""
        ):
            # bereits abgerechnet: In der Freigabe-Spalte steht z.B. die
            # Rechnungsnummer ('ja' -> '2026-001', siehe
            # vermerke_rechnungsnummern). Nicht erneut abgerechnet,
            # keine WARNUNG.
            erledigt.append(
                {
                    "kunde": kunde,
                    "bezeichnung": str(satz.get("bezeichnung", "")).strip(),
                    "belegnr": str(satz.get("belegnr", "")).strip(),
                    "vermerk": wert,
                }
            )
            continue
        if wert_klein not in {"yes", "ja"}:
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
                "datum": datum.isoformat() if datum else "",
                "art": str(satz.get("art", "")).strip(),
                "bezeichnung": str(satz.get("bezeichnung", "")).strip(),
                "betrag_netto": betrag,
                "belegnr": str(satz.get("belegnr", "")).strip(),
            }
        )
        freigegeben = True

    if not hat_spalte and zeilen_gesamt:
        nicht_freigegeben.insert(
            0,
            {"hinweis": f"Spalte '{spalte}' fehlt – keine Zeile freigegeben"},
        )

    # je Kunde nach Datum/Beleg sortieren
    for kunde in gruppe:
        gruppe[kunde].sort(key=lambda e: (e["datum"], e["belegnr"]))
    return {
        "datei_vorhanden": True,
        "freigegeben": freigegeben,
        "auslagen": gruppe,
        "nicht_freigegeben": nicht_freigegeben,
        "offene_zeilen": offene_zeilen,
        "erledigt": erledigt,
        "zeilen_gesamt": zeilen_gesamt,
    }


def vermerke_rechnungsnummern(
    conf: AuslagenConf, renr_je_kunde: dict[str, str]
) -> int:
    """Schreibt die Rechnungsnummer in die Freigabe-Spalte abgerechneter Zeilen.

    Für jeden Kunden in renr_je_kunde ({kunde: rechnungsnummer}) werden ALLE
    Zeilen mit freigabe='yes'/'ja' in der Auslagen-ODS als 'erledigt'
    markiert – datum-unabhängig. In der Freigabe-Spalte steht danach die
    Rechnungsnummer statt 'ja'. Die Zeilen werden bei künftigen Läufen als
    bereits abgerechnet erkannt (keine erneute Abrechnung, keine WARNUNG).
    Gibt die Anzahl der markierten Zeilen zurück.
    """
    if not renr_je_kunde:
        return 0
    spalte = str(conf.freigabe_spalte).strip().lower()

    def _neuer_wert(satz: dict[str, Any]) -> str | None:
        wert = str(satz.get(spalte, "")).strip()
        if wert.lower() not in {"yes", "ja"}:
            return None  # nicht freigegeben oder bereits erledigt
        kunde = str(satz.get("kunde", "")).strip()
        if kunde not in renr_je_kunde:
            return None
        return renr_je_kunde[kunde]

    try:
        return aendere_spalte(Path(conf.datei), conf.blatt, conf.freigabe_spalte, _neuer_wert)
    except OdsError as exc:
        raise AuslagenError(str(exc)) from exc


def auslagen_summe(zeilen: list[dict[str, Any]]) -> Decimal:
    return sum((z["betrag_netto"] for z in zeilen), Decimal("0"))
