"""Kundenstammdaten aus der pro-User ODS-Datei (z.B. conf/alice-kunden.ods).

Schema des Blatts 'Kunden' (eine Zeile je Kunde):
    kunde | name | strasse | plz | ort | land | ust_id | leitweg_id | stundensatz | zahlungsziel_tage

'stundensatz' als DE-Dezimal '95,00' oder EN '95.00' – beides parsebar.
'0,00' ist erlaubt (Kunde ohne Honorar, z.B. nur Auslagen/Services);
negativ ist unzulässig.
'ust_id' ist OPTIONAL (leer = kein Ausweis in PDF-Fuß und ZUGFeRD-XML).
'zahlungsziel_tage' ist OPTIONAL (Spaltentitel groß-/kleinschreibungsagnostisch,
jede Variante wie 'Zahlungsziel' oder 'zahlungsziel-tage' wird erkannt):
Anzahl Tage ab Rechnungsdatum. LEER = kein Zahlungsziel (weder PDF-Fuß
noch ZUGFeRD-BT-9) – die Rechnung spezifiziert dann keine Fälligkeit.
Fehlt die Spalte komplett: [erechnung] zahlungsziel_tage aus der conf
(leer dort = ebenfalls kein Zahlungsziel).

Spaltentitel sind generell groß-/kleinschreibungsagnostisch ('Kunde' wie
'kunde').
"""
from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Any

from .ods import blatt_als_dicts
from .util import dezimal

PFLICHTSPALTEN = ["kunde", "name", "plz", "ort", "land", "stundensatz"]
OPTIONAL_SPALTEN = ["strasse", "ust_id", "leitweg_id", "zahlungsziel_tage"]

# Spaltenpräfix für das Zahlungsziel: 'zahlungsziel_tage', 'Zahlungsziel',
# 'zahlungsziel (tage)' usw. werden alle erkannt.
_ZAHUNGSZIEL_PRAEFIX = "zahlungsziel"


class KundenError(Exception):
    pass


def _normalisiere_satz(satz: dict[str, Any]) -> dict[str, Any]:
    """Spaltentitel kleingeschrieben (groß-/kleinschreibungsagnostisch)."""
    aus: dict[str, Any] = {}
    for schluessel, wert in satz.items():
        k = str(schluessel).strip().lower()
        if k and k not in aus:
            aus[k] = wert
    return aus


def _zahlungsziel_tage(satz: dict[str, Any], kuerzel: str, spalte: str) -> int | None:
    """Zahlungsziel aus der (optionalen) Spalte; leer = kein Zahlungsziel.

    Wirft KundenError bei nicht-ganzzahliger Angabe – das Zahlungsziel
    darf nicht stillschweigend fallen gelassen werden.
    """
    wert = satz.get(spalte, "")
    roh = str(wert).strip() if wert is not None else ""
    if not roh:
        return None
    try:
        zahl = float(roh.replace(",", "."))
    except ValueError as exc:
        raise KundenError(
            f"Zahlungsziel für Kunde '{kuerzel}' ist nicht numerisch: {wert!r}"
        ) from exc
    if not zahl.is_integer():
        raise KundenError(
            f"Zahlungsziel für Kunde '{kuerzel}' ist nicht ganzzahlig: {wert!r}"
        )
    tage = int(zahl)
    if tage < 0:
        raise KundenError(
            f"Zahlungsziel für Kunde '{kuerzel}' ist negativ: {wert!r}"
        )
    return tage


@dataclass
class Kunde:
    kunde: str
    name: str
    strasse: str
    plz: str
    ort: str
    land: str
    ust_id: str
    leitweg_id: str
    stundensatz: Decimal
    # Zahlungsziel (Tage ab Rechnungsdatum); None = kein Zahlungsziel.
    # zahlungsziel_spalte: True, wenn die kunden.ods die Spalte führt –
    # dann entscheidet sie allein (leere Zelle = kein Zahlungsziel);
    # sonst gilt [erechnung] zahlungsziel_tage aus der conf.
    zahlungsziel_tage: int | None = None
    zahlungsziel_spalte: bool = False

    def faelligkeit(self, re_datum: _dt.date) -> _dt.date | None:
        """Fälligkeitsdatum (BT-9/Fuß) – None = nicht spezifiziert."""
        if self.zahlungsziel_tage is None:
            return None
        return re_datum + _dt.timedelta(days=self.zahlungsziel_tage)


def lade_kunden(datei: str | Path, blatt: str = "Kunden") -> dict[str, Kunde]:
    """Liefert {kunde_kuerzel: Kunde}.

    Wirft KundenError bei fehlenden Pflichtspalten, Duplikaten, nicht-
    numerischem oder negativem Stundensatz. '0,00' ist erlaubt.
    Die USt-IdNr. muss NICHT gesetzt sein.
    """
    saetze = blatt_als_dicts(datei, blatt)
    if not saetze:
        raise KundenError(f"Kundenblatt '{blatt}' in {datei} ist leer")
    # Spalten prüfen (Header aus erster gelesener Zeile, normalisiert)
    normalisiert = [_normalisiere_satz(s) for s in saetze]
    vorhandene = set(normalisiert[0].keys())
    fehlt = [s for s in PFLICHTSPALTEN if s not in vorhandene]
    if fehlt:
        raise KundenError(
            f"Pflichtspalten fehlen in {datei}!{blatt}: {', '.join(fehlt)}"
        )
    # Optionale Zahlungsziel-Spalte erkennen (Header-Präfix)
    zahlungsziel_key = ""
    for schluessel in vorhandene:
        if schluessel.startswith(_ZAHUNGSZIEL_PRAEFIX):
            zahlungsziel_key = schluessel
            break
    hat_zahlungsziel_spalte = bool(zahlungsziel_key)
    kunden: dict[str, Kunde] = {}
    for satz in normalisiert:
        kuerzel = str(satz.get("kunde", "")).strip()
        if not kuerzel:
            continue
        if kuerzel in kunden:
            raise KundenError(f"Kunde '{kuerzel}' mehrfach in {datei}!{blatt}")
        stundensatz = dezimal(satz.get("stundensatz"))
        if stundensatz < 0:
            raise KundenError(
                f"Stundensatz für Kunde '{kuerzel}' ist negativ: {satz.get('stundensatz')!r}"
            )
        tage = _zahlungsziel_tage(satz, kuerzel, zahlungsziel_key) if zahlungsziel_key else None
        kunden[kuerzel] = Kunde(
            kunde=kuerzel,
            name=str(satz.get("name", "")).strip(),
            strasse=str(satz.get("strasse", "")).strip(),
            plz=str(satz.get("plz", "")).strip(),
            ort=str(satz.get("ort", "")).strip(),
            land=str(satz.get("land", "DE")).strip().upper(),
            ust_id=str(satz.get("ust_id", "")).strip(),
            leitweg_id=str(satz.get("leitweg_id", "")).strip(),
            stundensatz=stundensatz,
            zahlungsziel_tage=tage,
            zahlungsziel_spalte=hat_zahlungsziel_spalte,
        )
    if not kunden:
        raise KundenError(f"Keine Kunden in {datei}!{blatt} gefunden")
    return kunden
