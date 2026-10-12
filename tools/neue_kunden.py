#!/usr/bin/env python3
"""Erzeugt die Kunden-ODS: conf/<user>-kunden.ods.

Blatt 'Kunden' mit dem Header
    kunde | name | strasse | plz | ort | land | ust_id | leitweg_id | stundensatz | zahlungsziel_tage
und Blatt 'Services' mit dem Header
    kunde | service | Kosten pro Stück [€] | 1 | 2 | … | 12

wie von den [kunden]/[services]-Erwartungen von REU gebraucht
(Spaltentitel sind groß-/kleinschreibungsagnostisch).

Die Spalte 'zahlungsziel_tage' ist OPTIONAL: Anzahl Tage ab
Rechnungsdatum; LEER = kein Zahlungsziel (die Rechnung spezifiziert
keine Fälligkeit – kein 'Zahlbar bis' im PDF-Fuß, kein BT-9 im
ZUGFeRD-XML). Fehlt die Spalte komplett, gilt [erechnung]
zahlungsziel_tage aus der conf (dort ebenfalls leer = kein Zahlungsziel).

Nutzung:
    python3 tools/neue_kunden.py <user>             # leere Datei (nur Header)
    python3 tools/neue_kunden.py <user> --beispiel # Demo-Zeilen
    python3 tools/neue_kunden.py <user> --force     # vorhandene Datei überschreiben

Danach in conf/<user>.conf eintragen:
    [kunden]
    datei = conf/<user>-kunden.ods
    blatt = Kunden
"""
import argparse
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

KOPF = [
    "kunde", "name", "strasse", "plz", "ort", "land",
    "ust_id", "leitweg_id", "stundensatz", "zahlungsziel_tage",
]

BEISPIEL = [
    ["ABC", "Alpha Business Consulting GmbH", "Bahnhofstraße 1", "12345",
     "Musterstadt", "DE", "DE123456789", "", "95,00", ""],
    ["DEF", "Delta Energy e.G.", "Weg 2", "54321", "Beispielheim", "DE",
     "", "", "110,00", "21"],
]

SERVICES_KOPF = ["kunde", "service", "Kosten pro Stück [€]"] + [str(m) for m in range(1, 13)]
SERVICES_BEISPIEL = [
    ["ABC", "Hosting & Pflege", "25,00"] + [""] * 6 + ["2", "1", "x"] + [""] * 3,
    ["DEF", "Monitoring", "12,50"] + [""] * 6 + ["x", "", "3"] + [""] * 3,
]

MIMETYPE = "application/vnd.oasis.opendocument.spreadsheet"

MANIFEST = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">'
    '<manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
    '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>'
    '<manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>'
    '<manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>'
    '</manifest:manifest>'
)

STYLES = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2">'
    '<office:styles/></office:document-styles>'
)

META = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2">'
    '<office:meta/></office:document-meta>'
)


def _zelle(wert) -> str:
    return (
        '<table:table-cell office:value-type="string">'
        f'<text:p>{escape(str(wert))}</text:p></table:table-cell>'
    )


def _zeile(werte: list) -> str:
    return '<table:table-row>' + ''.join(_zelle(w) for w in werte) + '</table:table-row>'


def _blatt(name: str, kopf: list, zeilen: list) -> str:
    return (
        f'<table:table table:name="{escape(name)}">'
        + '<table:table-column/>'
        + _zeile(kopf)
        + ''.join(_zeile(z) for z in zeilen)
        + '</table:table>'
    )


def content_xml(zeilen: list, services_zeilen: list) -> str:
    tabellen = (
        _blatt("Kunden", KOPF, zeilen)
        + _blatt("Services", SERVICES_KOPF, services_zeilen)
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document-content'
        ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
        ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"'
        ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"'
        ' office:version="1.2">'
        '<office:body><office:spreadsheet>'
        + tabellen
        + '</office:spreadsheet></office:body>'
        '</office:document-content>'
    )


def schreibe_ods(pfad: Path, zeilen: list, services_zeilen: list) -> None:
    pfad.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pfad, "w") as z:
        info = zipfile.ZipInfo("mimetype")
        z.writestr(info, MIMETYPE, compress_type=zipfile.ZIP_STORED)
        z.writestr("META-INF/manifest.xml", MANIFEST)
        z.writestr("meta.xml", META)
        z.writestr("styles.xml", STYLES)
        z.writestr("content.xml", content_xml(zeilen, services_zeilen))


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Erzeugt conf/<user>-kunden.ods (Blätter 'Kunden' und 'Services')."
    )
    ap.add_argument("user", help="Benutzername, z. B. alice oder philipp")
    ap.add_argument(
        "--beispiel", action="store_true",
        help="Demo-Kunden und Service-Zeilen einfügen (u. a. Zahlungsziel 21 Tage vs. leer = keins)",
    )
    ap.add_argument("--datei", default=None, help="Zielpfad (Default: conf/<user>-kunden.ods)")
    ap.add_argument("--force", action="store_true", help="vorhandene Datei überschreiben")
    args = ap.parse_args()

    pfad = Path(args.datei) if args.datei else Path("conf") / f"{args.user}-kunden.ods"
    if pfad.exists() and not args.force:
        print(f"FEHLER: {pfad} existiert bereits (--force zum Überschreiben).")
        return 1

    zeilen = BEISPIEL if args.beispiel else []
    services_zeilen = SERVICES_BEISPIEL if args.beispiel else []
    schreibe_ods(pfad, zeilen, services_zeilen)
    print(
        f"OK: {pfad} erzeugt (Blatt 'Kunden', {len(KOPF)} Spalten, "
        f"{len(zeilen)} Beispielzeile(n); Blatt 'Services', "
        f"{len(SERVICES_KOPF)} Spalten, {len(services_zeilen)} Beispielzeile(n))."
    )
    print("In conf/" + args.user + ".conf eintragen:")
    print("    [kunden]")
    print(f"    datei = {pfad.as_posix()}")
    print("    blatt = Kunden")
    return 0


if __name__ == "__main__":
    sys.exit(main())
