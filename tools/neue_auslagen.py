#!/usr/bin/env python3
"""Erzeugt eine zentrale Auslagen-ODS: conf/<user>-auslagen.ods.

Blatt 'Auslagen' mit dem Header
    datum | kunde | art | bezeichnung | betrag_netto | belegnr | freigabe
wie von der [auslagen]-Sektion der conf erwartet (pflichtspalten +
freigabe_spalte, Spaltentitel sind groß-/kleinschreibungsagnostisch).

Nutzung:
    python3 tools/neue_auslagen.py <user>             # leere Datei (nur Header)
    python3 tools/neue_auslagen.py <user> --beispiel  # Alice-Demozeilen
    python3 tools/neue_auslagen.py <user> --force     # vorhandene Datei überschreiben

Danach in conf/<user>.conf eintragen:
    [auslagen]
    quelle = ods
    datei = conf/<user>-auslagen.ods
    blatt = Auslagen
"""
import argparse
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

KOPF = ["datum", "kunde", "art", "bezeichnung", "betrag_netto", "belegnr", "freigabe"]

# Alice-Demozeilen: alle Freigabe-Zustände (ja/yes, leer, nein, Re-Nr = erledigt)
BEISPIEL = [
    ["2025-11-03", "ABC", "Fahrt", "Bahnfahrkarte", "42,00", "B-0042", "ja"],
    ["2025-11-10", "ABC", "Material", "USB-Stick", "19,90", "B-0043", ""],
    ["2025-11-15", "DEF", "Fahrt", "Tanken", "65,50", "B-0044", "yes"],
    ["2025-11-20", "DEF", "Post", "Briefporto", "3,30", "B-0045", "nein"],
    ["2025-11-25", "ABC", "Fahrt", "Taxi", "18,00", "B-0046", "2025-041"],
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


def _zelle(wert: str) -> str:
    return (
        '<table:table-cell office:value-type="string">'
        f'<text:p>{escape(wert)}</text:p></table:table-cell>'
    )


def _zeile(werte: list) -> str:
    return '<table:table-row>' + ''.join(_zelle(str(w)) for w in werte) + '</table:table-row>'


def content_xml(zeilen: list) -> str:
    tab = (
        '<table:table table:name="Auslagen">'
        + '<table:table-column/>' * len(KOPF)
        + _zeile(KOPF)
        + ''.join(_zeile(z) for z in zeilen)
        + '</table:table>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<office:document-content'
        ' xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
        ' xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"'
        ' xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"'
        ' office:version="1.2">'
        '<office:body><office:spreadsheet>' + tab + '</office:spreadsheet></office:body>'
        '</office:document-content>'
    )


def schreibe_ods(pfad: Path, zeilen: list) -> None:
    pfad.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pfad, "w") as z:
        info = zipfile.ZipInfo("mimetype")
        z.writestr(info, MIMETYPE, compress_type=zipfile.ZIP_STORED)
        z.writestr("META-INF/manifest.xml", MANIFEST)
        z.writestr("meta.xml", META)
        z.writestr("styles.xml", STYLES)
        z.writestr("content.xml", content_xml(zeilen))


def main() -> int:
    ap = argparse.ArgumentParser(description="Erzeugt conf/<user>-auslagen.ods (Blatt 'Auslagen').")
    ap.add_argument("user", help="Benutzername, z. B. alice oder philipp")
    ap.add_argument("--beispiel", action="store_true", help="Alice-Demozeilen mit allen Freigabe-Zuständen einfügen")
    ap.add_argument("--datei", default=None, help="Zielpfad (Default: conf/<user>-auslagen.ods)")
    ap.add_argument("--force", action="store_true", help="vorhandene Datei überschreiben")
    args = ap.parse_args()

    pfad = Path(args.datei) if args.datei else Path("conf") / f"{args.user}-auslagen.ods"
    if pfad.exists() and not args.force:
        print(f"FEHLER: {pfad} existiert bereits (--force zum Überschreiben).")
        return 1

    zeilen = BEISPIEL if args.beispiel else []
    schreibe_ods(pfad, zeilen)
    print(f"OK: {pfad} erzeugt (Blatt 'Auslagen', {len(KOPF)} Spalten, {len(zeilen)} Beispielzeile(n)).")
    print("In conf/" + args.user + ".conf eintragen:")
    print("    [auslagen]")
    print("    quelle = ods")
    print(f"    datei = {pfad.as_posix()}")
    print("    blatt = Auslagen")
    return 0


if __name__ == "__main__":
    sys.exit(main())
