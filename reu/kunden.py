"""Kundenstammdaten aus der pro-User ODS-Datei (z.B. conf/alice-kunden.ods).

Schema des Blatts 'Kunden' (eine Zeile je Kunde):
    kunde | name | strasse | plz | ort | land | ust_id | leitweg_id | stundensatz

`stundensatz` als DE-Dezimal '95,00' oder EN '95.00' – beides parsebar.
"""
from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path

from .ods import blatt_als_dicts
from .util import dezimal

PFLICHTSPALTEN = ["kunde", "name", "plz", "ort", "land", "ust_id", "stundensatz"]
OPTIONAL_SPALTEN = ["strasse", "leitweg_id"]


class KundenError(Exception):
    pass


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


def lade_kunden(datei: str | Path, blatt: str = "Kunden") -> dict[str, Kunde]:
    """Liefert {kunde_kuerzel: Kunde}.

    Wirft KundenError bei fehlenden Pflichtspalten, Duplikaten oder
    nicht-numerischem Stundensatz.
    """
    saetze = blatt_als_dicts(datei, blatt)
    if not saetze:
        raise KundenError(f"Kundenblatt '{blatt}' in {datei} ist leer")
    # Spalten prüfen (Header aus erster gelesener Zeile)
    vorhandene = set(saetze[0].keys())
    fehlt = [s for s in PFLICHTSPALTEN if s not in vorhandene]
    if fehlt:
        raise KundenError(
            f"Pflichtspalten fehlen in {datei}!{blatt}: {', '.join(fehlt)}"
        )
    kunden: dict[str, Kunde] = {}
    for satz in saetze:
        kuerzel = str(satz.get("kunde", "")).strip()
        if not kuerzel:
            continue
        if kuerzel in kunden:
            raise KundenError(f"Kunde '{kuerzel}' mehrfach in {datei}!{blatt}")
        stundensatz = dezimal(satz.get("stundensatz"))
        if stundensatz <= 0:
            raise KundenError(
                f"Stundensatz für Kunde '{kuerzel}' fehlt oder ist <= 0: {satz.get('stundensatz')!r}"
            )
        ust_id = str(satz.get("ust_id", "")).strip()
        if not ust_id:
            raise KundenError(f"USt-IdNr. fehlt für Kunde '{kuerzel}' (B2B Pflicht)")
        kunden[kuerzel] = Kunde(
            kunde=kuerzel,
            name=str(satz.get("name", "")).strip(),
            strasse=str(satz.get("strasse", "")).strip(),
            plz=str(satz.get("plz", "")).strip(),
            ort=str(satz.get("ort", "")).strip(),
            land=str(satz.get("land", "DE")).strip().upper(),
            ust_id=ust_id,
            leitweg_id=str(satz.get("leitweg_id", "")).strip(),
            stundensatz=stundensatz,
        )
    if not kunden:
        raise KundenError(f"Keine Kunden in {datei}!{blatt} gefunden")
    return kunden
