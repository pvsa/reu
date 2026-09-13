"""Einlesen von ODS-Dateien über pyexcel-ods3.

Liefert die Daten als Liste von Dicts je Zeile (Header-Zeile im Blatt
wird als Schlüssel verwendet). Leere Zeilen werden übersprungen.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from pyexcel_ods3 import get_data


class OdsError(Exception):
    """Fehler beim Einlesen einer ODS-Datei."""


def lade_blatt(datei: str | Path, blatt: str) -> list[list[Any]]:
    """Liefert die Roh-Zeilen des Blatts als Liste von Listen.

    Wirft OdsError, wenn Datei fehlt oder Blatt nicht existiert.
    """
    pfad = Path(datei)
    if not pfad.is_file():
        raise OdsError(f"ODS-Datei nicht gefunden: {pfad}")
    try:
        workbook = get_data(str(pfad))
    except Exception as exc:  # noqa: BLE001 – Fehler weiterreichen mit Kontext
        raise OdsError(f"ODS-Datei konnte nicht gelesen werden {pfad}: {exc}") from exc
    if blatt not in workbook:
        verfuegbar = ", ".join(sorted(workbook.keys())) or "(keine)"
        raise OdsError(f"Blatt '{blatt}' fehlt in {pfad}. Vorhanden: {verfuegbar}")
    return workbook[blatt]


def blatt_als_dicts(datei: str | Path, blatt: str) -> list[dict[str, Any]]:
    """Liest ein Blatt und gibt eine Liste von Dicts (Header -> Wert) zurück.

    Die erste nicht-leere Zeile definiert die Spaltennamen. Leere Zeilen
    (komplett leer) werden ignoriert. Zellwerte werden per str() normalisiert
    ist None → "".
    """
    zeilen = lade_blatt(datei, blatt)
    if not zeilen:
        return []
    header: list[str] | None = None
    satzliste: list[dict[str, Any]] = []
    for zeile in zeilen:
        werte = [(z if z is not None else "") for z in zeile]
        if all(w == "" for w in werte):
            continue
        if header is None:
            header = [str(w).strip() for w in werte]
            continue
        # Zeile auf Header-Länge bringen
        while len(werte) < len(header):
            werte.append("")
        satz = {header[i]: werte[i] for i in range(len(header))}
        satzliste.append(satz)
    return satzliste


def zelle(datei: str | Path, blatt: str, zelle_ref: str) -> Any:
    """Liest eine einzelne Zelle, Adresse in Calc-Notation 'Meta!A1' oder 'A1'.

    Nur Spalten A..ZZ und Zeilen ab 1 werden unterstützt.
    """
    if "!" in zelle_ref:
        blatt_name, ref = zelle_ref.split("!", 1)
        blatt = blatt_name
    else:
        ref = zelle_ref
    spalte_buchstaben = "".join(ch for ch in ref if ch.isalpha()).upper()
    zeilen_nr = "".join(ch for ch in ref if ch.isdigit())
    if not spalte_buchstaben or not zeilen_nr:
        raise OdsError(f"Ungültige Zelladresse '{zelle_ref}'")
    spalte = 0
    for ch in spalte_buchstaben:
        spalte = spalte * 26 + (ord(ch) - ord("A") + 1)
    spalte -= 1
    zeile_idx = int(zeilen_nr) - 1
    matrix = lade_blatt(datei, blatt)
    if zeile_idx >= len(matrix):
        return None
    zeile = matrix[zeile_idx]
    if spalte >= len(zeile):
        return None
    return zeile[spalte]
