"""Einlesen von ODS-Dateien ohne externe Abhängigkeiten.

ODS ist ein ZIP-Archiv mit XML-Inhalt; gelesen wird nur content.xml
mit den Python-Standardmodulen zipfile und xml.etree.ElementTree.
Damit ist kein pip-Paket und kein zusätzliches apt-Paket nötig.

Liefert die Daten wie bisher pyexcel-ods3: {blattname: [zeile, ...]},
wobei jede Zeile eine Liste von Zellwerten ist
(str / float / bool / datetime.date / datetime.datetime / None).
"""
from __future__ import annotations

import datetime as _dt
import os
import xml.etree.ElementTree as Et
import zipfile
from pathlib import Path
from typing import Any, Callable

_TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
_OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
_TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"

_TABLE = f"{{{_TABLE_NS}}}table"
_ROW = f"{{{_TABLE_NS}}}table-row"
_COVERED_CELL = f"{{{_TABLE_NS}}}covered-table-cell"
_CELL = f"{{{_TABLE_NS}}}table-cell"
_COLS_REPEATED = f"{{{_TABLE_NS}}}number-columns-repeated"
_ROWS_REPEATED = f"{{{_TABLE_NS}}}number-rows-repeated"
_NAME = f"{{{_TABLE_NS}}}name"
_VALUE_TYPE = f"{{{_OFFICE_NS}}}value-type"
_VALUE = f"{{{_OFFICE_NS}}}value"
_DATE_VALUE = f"{{{_OFFICE_NS}}}date-value"
_BOOL_VALUE = f"{{{_OFFICE_NS}}}boolean-value"
_TEXT_P = f"{{{_TEXT_NS}}}p"
_TEXT_S = f"{{{_TEXT_NS}}}s"
_TEXT_TAB = f"{{{_TEXT_NS}}}tab"
_TEXT_LB = f"{{{_TEXT_NS}}}line-break"

# LibreOffice schreibt leere Füllzeilen/-spalten mit riesigen
# Wiederholungszählungen (bis zu 2**20); echte Lücken im Datenbereich
# haben kleine Zähler. Große Wiederholungen leerer Zeilen/Zellen werden
# daher ignoriert, um Speicher und Laufzeit zu schonen.
_MAX_LEER_WIEDERHOLUNG = 1024


class OdsError(Exception):
    """Fehler beim Einlesen einer ODS-Datei."""


def _zellen_text(zelle: Et.Element) -> str:
    """Extrahiert den Textinhalt einer Zelle (inkl. text:s, Tab, Umbruch)."""
    teile: list[str] = []
    for p in zelle.findall(_TEXT_P):
        for knoten in p.iter():
            if knoten.tag == _TEXT_S:
                anzahl = knoten.get(f"{{{_TEXT_NS}}}c", "1")
                teile.append(" " * int(anzahl or 1))
            elif knoten.tag == _TEXT_TAB:
                teile.append("\t")
            elif knoten.tag == _TEXT_LB:
                teile.append("\n")
            elif knoten.text:
                teile.append(knoten.text)
            if knoten.tail:
                teile.append(knoten.tail)
    # Doppelte Knoten vermeiden: p.iter() liefert p selbst (mit .text/None)
    return "".join(teile).strip()


def _zellen_wert(zelle: Et.Element) -> Any:
    """Wandelt eine Zelle in einen Python-Wert um."""
    if zelle.tag == _COVERED_CELL:
        return None
    typ = zelle.get(_VALUE_TYPE)
    if typ in ("float", "percentage", "currency"):
        roh = zelle.get(_VALUE)
        if roh is None:
            return None
        try:
            return float(roh)
        except ValueError:
            return roh
    if typ == "date":
        roh = zelle.get(_DATE_VALUE) or ""
        if not roh:
            return None
        try:
            if "T" in roh:
                return _dt.datetime.fromisoformat(roh)
            return _dt.date.fromisoformat(roh[:10])
        except ValueError:
            return roh
    if typ == "boolean":
        return zelle.get(_BOOL_VALUE) == "true"
    if typ == "string" or typ is None:
        text = _zellen_text(zelle)
        return text if text else None
    # Unbekannter Typ: Textfallback
    return _zellen_text(zelle) or None


def _zeile_als_liste(zeile: Et.Element) -> list[Any]:
    """Eine table:table-row in eine Werteliste umwandeln."""
    werte: list[Any] = []
    for zelle in zeile:
        if zelle.tag not in (_CELL, _COVERED_CELL):
            continue
        wert = _zellen_wert(zelle)
        wiederholt = zelle.get(_COLS_REPEATED)
        anzahl = int(wiederholt) if wiederholt else 1
        if wert is None and anzahl > _MAX_LEER_WIEDERHOLUNG:
            # leere Füllspalten am Zeilenende
            continue
        werte.extend([wert] * anzahl)
    return werte


def _lese_ods(pfad: Path) -> dict[str, list[list[Any]]]:
    """Liest alle Blätter einer ODS-Datei: {blattname: [zeile, ...]}."""
    try:
        with zipfile.ZipFile(pfad) as archiv:
            content = archiv.read("content.xml")
    except (zipfile.BadZipFile, KeyError) as exc:
        raise OdsError(f"Keine gültige ODS-Datei: {pfad} ({exc})") from exc
    except OSError as exc:
        raise OdsError(f"ODS-Datei konnte nicht gelesen werden {pfad}: {exc}") from exc

    try:
        wurzel = Et.fromstring(content)
    except Et.ParseError as exc:
        raise OdsError(f"content.xml unlesbar in {pfad}: {exc}") from exc

    blaetter: dict[str, list[list[Any]]] = {}
    for tabelle in wurzel.iter(_TABLE):
        name = tabelle.get(_NAME) or ""
        zeilen: list[list[Any]] = []
        for zeile in tabelle.iter(_ROW):
            werte = _zeile_als_liste(zeile)
            wiederholt = zeile.get(_ROWS_REPEATED)
            anzahl = int(wiederholt) if wiederholt else 1
            if not any(w is not None for w in werte) and anzahl > _MAX_LEER_WIEDERHOLUNG:
                # leere Füllzeilen am Blattende
                continue
            zeilen.extend([werte] * anzahl)
        blaetter[name] = zeilen
    return blaetter


def lade_blatt(datei: str | Path, blatt: str) -> list[list[Any]]:
    """Liefert die Roh-Zeilen des Blatts als Liste von Listen.

    Wirft OdsError, wenn Datei fehlt oder Blatt nicht existiert.
    """
    pfad = Path(datei)
    if not pfad.is_file():
        raise OdsError(f"ODS-Datei nicht gefunden: {pfad}")
    workbook = _lese_ods(pfad)
    if blatt not in workbook:
        # Groß-/Kleinschreibung des Blattnamens tolerieren ('services' == 'Services')
        treffer = [n for n in workbook if n.lower() == blatt.lower()]
        if treffer:
            return workbook[treffer[0]]
        verfuegbar = ", ".join(sorted(workbook.keys())) or "(keine)"
        raise OdsError(f"Blatt '{blatt}' fehlt in {pfad}. Vorhanden: {verfuegbar}")
    return workbook[blatt]


def blatt_als_dicts(datei: str | Path, blatt: str) -> list[dict[str, Any]]:
    """Liest ein Blatt und gibt eine Liste von Dicts (Header -> Wert) zurück.

    Die erste nicht-leere Zeile definiert die Spaltennamen. Leere Zeilen
    (komplett leer) werden ignoriert. Zellwerte werden per str() normalisiert,
    ist None -> "".
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


# ---------------------------------------------------------------- Schreiben

# Namensräume, die beim Zurückschreiben mit bekannten Präfixen erhalten
# bleiben sollen (unbekannte erhält ns0:-Präfixe – ebenfalls gültiges XML).
_XMLNS = (
    ("office", _OFFICE_NS),
    ("table", _TABLE_NS),
    ("text", _TEXT_NS),
    ("xlink", "http://www.w3.org/1999/xlink"),
    ("dc", "http://purl.org/dc/elements/1.1/"),
    ("meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"),
    ("fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"),
    ("svg", "http://www.w3.org/2000/svg"),
    ("of", "urn:oasis:names:tc:opendocument:xmlns:of:1.2"),
    ("calcext", "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"),
    ("loext", "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"),
    ("grddl", "http://www.w3.org/2003/g/data-view#"),
    ("css3t", "http://www.w3.org/TR/css3-text/"),
)


def _zeile_mit_zellen(zeile: Et.Element) -> tuple[list[Any], list[Et.Element]]:
    """Wie _zeile_als_liste, zusätzlich mit dem Zell-Element je Position."""
    werte: list[Any] = []
    zellen: list[Et.Element] = []
    for zelle in zeile:
        if zelle.tag not in (_CELL, _COVERED_CELL):
            continue
        wert = _zellen_wert(zelle)
        wiederholt = zelle.get(_COLS_REPEATED)
        anzahl = int(wiederholt) if wiederholt else 1
        if wert is None and anzahl > _MAX_LEER_WIEDERHOLUNG:
            continue
        werte.extend([wert] * anzahl)
        zellen.extend([zelle] * anzahl)
    return werte, zellen


def _setze_textzelle(zelle: Et.Element, text: str) -> None:
    """Macht aus der Zelle eine String-Zelle mit dem gegebenen Text.

    Struktur-Attribute (number-columns-repeated, Stil) bleiben erhalten;
    Typ-/Wert-Attribute und Zellinhalte werden ersetzt.
    """
    for attr in list(zelle.attrib):
        if attr.startswith(f"{{{_OFFICE_NS}}}") and attr != _VALUE_TYPE:
            del zelle.attrib[attr]
    zelle.set(_VALUE_TYPE, "string")
    for kind in list(zelle):
        zelle.remove(kind)
    absatz = Et.SubElement(zelle, _TEXT_P)
    absatz.text = text


def aendere_spalte(
    datei: str | Path,
    blatt: str,
    spalte: str,
    wert_funktion: Callable[[dict[str, Any]], str | None],
) -> int:
    """Setzt Zellenwerte einer Spalte über alle Datenzeilen des Blatts.

    spalte: Headername der Zielspalte (Groß-/Kleinschreibung egal).
    wert_funktion(satz) erhält je Datenzeile ein Dict der Werte mit
    normalisierten (kleingeschriebenen) Headern als Schlüssel und liefert
    den neuen Zelltext – oder None, wenn die Zeile unverändert bleibt.
    Die erste nicht-leere Zeile gilt wie bei blatt_als_dicts als Kopfzeile
    und wird nie geändert.

    Gibt die Anzahl geänderter Zellen zurück (0 = nichts gespeichert).
    Die Datei wird atomar gespeichert (temporäre Datei + replace), damit
    nie ein halbes Archiv entstehen kann. Wirft OdsError bei Problemen.
    """
    pfad = Path(datei)
    if not pfad.is_file():
        raise OdsError(f"ODS-Datei nicht gefunden: {pfad}")
    try:
        with zipfile.ZipFile(pfad) as archiv:
            content = archiv.read("content.xml")
    except (zipfile.BadZipFile, KeyError) as exc:
        raise OdsError(f"Keine gültige ODS-Datei: {pfad} ({exc})") from exc

    try:
        wurzel = Et.fromstring(content)
    except Et.ParseError as exc:
        raise OdsError(f"content.xml unlesbar in {pfad}: {exc}") from exc

    tabelle = None
    for t in wurzel.iter(_TABLE):
        if (t.get(_NAME) or "").lower() == blatt.lower():
            tabelle = t
            break
    if tabelle is None:
        raise OdsError(f"Blatt '{blatt}' fehlt in {pfad}.")

    ziel: int | None = None
    header: list[str] = []
    geaendert = 0
    for zeile in tabelle.iter(_ROW):
        werte, zellen = _zeile_mit_zellen(zeile)
        if not any(w is not None for w in werte):
            continue
        if ziel is None:
            # erste nicht-leere Zeile = Kopfzeile: Zielspalte bestimmen
            header = [str(w).strip().lower() if w is not None else "" for w in werte]
            ziel = next((i for i, h in enumerate(header) if h == spalte.strip().lower()), None)
            if ziel is None:
                vorh = ", ".join(h for h in header if h) or "(keine)"
                raise OdsError(
                    f"Spalte '{spalte}' fehlt in {pfad}!{blatt}. Vorhanden: {vorh}"
                )
            continue
        if ziel >= len(werte) or ziel >= len(zellen):
            continue  # Zeile endet vor der Zielspalte
        satz = {header[i]: werte[i] for i in range(len(header)) if header[i]}
        neu = wert_funktion(satz)
        if neu is None:
            continue
        _setze_textzelle(zellen[ziel], str(neu))
        geaendert += 1

    if geaendert == 0:
        return 0

    for praefix, uri in _XMLNS:
        Et.register_namespace(praefix, uri)
    neuer_content = Et.tostring(wurzel, encoding="UTF-8", xml_declaration=True)
    tmp = pfad.with_name(pfad.name + ".tmp")
    try:
        with zipfile.ZipFile(pfad) as alt, zipfile.ZipFile(tmp, "w") as neu:
            for info in alt.infolist():
                if info.filename == "content.xml":
                    continue
                neu.writestr(info, alt.read(info.filename))
            neu.writestr("content.xml", neuer_content)
        os.replace(tmp, pfad)
    except OSError as exc:
        tmp.unlink(missing_ok=True)
        raise OdsError(f"ODS-Datei konnte nicht gespeichert werden: {pfad} ({exc})") from exc
    return geaendert
