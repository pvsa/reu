"""PDF/A-3-Konvertierung der Rechnungs-PDF mit Ghostscript (pdfwrite).

Die ZUGFeRD-/Factur-X-Spezifikation verlangt für hybride Rechnungen
einen PDF/A-3B-Container (ISO 19005-3). reportlab (open source) kann
kein PDF/A erzeugen – Base-14-Fonts wie Helvetica werden nicht
eingebettet, Farben sind nicht ICC-charakterisiert. Ghostscript
schließt genau diese Lücke:

    gs --permit-file-read=<icc> -sDEVICE=pdfwrite -dPDFA=3
       -sColorConversionStrategy=RGB -dPDFACompatibilityPolicy=1
       -o ausgabe.pdf <PDFA_def.ps> eingabe.pdf

- **Fonts einbetten, Farben charakterisieren:** pdfwrite bettet die
  Schriften ein und normalisiert alle Farben auf das sRGB-ICC-Profil
  (OutputIntent + pdfaid-XMP).
- **PDFA_def.ps:** Die mit Ghostscript gelieferte Definitionsdatei
  (lib-Verzeichnis) setzt den OutputIntent. Sie referenziert
  standardmäßig 'srgb.icc' im Arbeitsverzeichnis – REU ersetzt den
  Pfad in einer temporären Kopie durch das gefundene sRGB-Profil
  (/ICCProfile-Zeile, markiert '% Customise').
- **-dPDFACompatibilityPolicy=1:** Enthält das PDF etwas, das nicht
  PDF/A-fähig ist, bricht gs mit FEHLER ab statt stillschweigend ein
  unkonformes PDF mit PDF/A-Metadaten zu schreiben (Default 0 würde
  weiter schreiben).
- **SAFER:** gs läuft im sicheren Modus (kein -dNOSAFER); die
  ICC-Datei wird per --permit-file-read expliziet lesbar gemacht.

Das Einbetten des XML (AFRelationship /Alternative, XMP mit pdfaid +
Factur-X-Extension-Schema, Übernahme des OutputIntents) macht danach
wie bisher drafthorse.pdf.attach_xml – die Reihenfolge ist: reportlab
→ PDF/A-3 (dieses Modul) → drafthorse-Embedding.

Konfiguration ([erechnung]): format = pdfa3 schaltet den Schritt an
(Standard). gs_bin, icc_profil und pdfa_def überschreiben die
automatische Suche (nützlich für Exoten-Installationen und Tests).
"""
from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from pathlib import Path

# Verzeichnisse, in denen Ghostscript-Versionen liegen (Debian/Ubuntu:
# /usr/share/ghostscript/<version>/lib bzw. .../iccprofiles).
_SUCHE_BASIS = (
    "/usr/share/ghostscript",
    "/usr/local/share/ghostscript",
    "/usr/lib/ghostscript",
)

# ICC-Profile relativ zur gs-Versionshierarchie und Fallbacks außerhalb
_ICC_MUSTER = ("*/iccprofiles/srgb.icc", "*/iccprofiles/default_rgb.icc")
_ICC_FALLBACK = ("/usr/share/color/icc/colord/sRGB.icc",)

# /ICCProfile-Zeile in PDFA_def.ps ('% Customise' bei Ghostscript)
_RE_ICCPROFIL = re.compile(r"^/ICCProfile\s*\((.*?)\)", re.MULTILINE)

_GS_TIMEOUT = 120  # Sekunden für eine Rechnungs-PDF


class PdfAError(Exception):
    pass


def _finde_datei(muster: tuple[str, ...], fallback: tuple[str, ...] = ()) -> Path | None:
    """Sucht die neueste Version einer gs-Ressource (höchster Versionspfad)."""
    for basis in _SUCHE_BASIS:
        treffer = sorted(Path(basis).glob(muster[0]))
        for m in muster[1:]:
            treffer += [p for p in Path(basis).glob(m) if p not in treffer]
        if treffer:
            return treffer[-1]
    for pfad in fallback:
        if Path(pfad).is_file():
            return Path(pfad)
    return None


def _finde_pdfa_def() -> Path | None:
    """PDFA_def.ps (evtl. auch klein geschrieben pdfa_def.ps) im gs-lib-Verzeichnis."""
    for basis in _SUCHE_BASIS:
        for name in ("PDFA_def.ps", "pdfa_def.ps"):
            treffer = sorted(Path(basis).glob(f"*/lib/{name}"))
            if treffer:
                return treffer[-1]
    return None


def _patche_pdfa_def(inhalt: str, icc_pfad: Path) -> str:
    """Ersetzt die /ICCProfile-Zeile der Definitionsdatei durch den Profil-Pfad."""
    if not _RE_ICCPROFIL.search(inhalt):
        raise PdfAError(
            "PDFA_def.ps enthält keine /ICCProfile-Zeile – unbekanntes Format "
            "der Ghostscript-Definitionsdatei."
        )
    return _RE_ICCPROFIL.sub(f"/ICCProfile ({icc_pfad})", inhalt, count=1)


def pdfa3_konvertieren(
    pdf_bytes: bytes,
    *,
    gs_bin: str = "gs",
    icc_profil: str | Path | None = None,
    pdfa_def: str | Path | None = None,
) -> bytes:
    """Konvertiert eine PDF in PDF/A-3B (Bytes rein, Bytes raus).

    gs_bin:   ausführbare Ghostscript-Datei (Default 'gs').
    icc_profil / pdfa_def: explizite Pfade; ohne Angabe werden srgb.icc
    und PDFA_def.ps der installierten gs-Version automatisch gesucht.

    Wirft PdfAError mit klarer Meldung, wenn gs fehlt, die Ressourcen
    nicht gefunden werden oder gs die Konvertierung abbricht
    (-dPDFACompatibilityPolicy=1: unkonforme Inhalte sind ein Fehler).
    """
    if not pdf_bytes:
        raise PdfAError("Keine PDF-Daten zum Konvertieren übergeben.")

    if shutil.which(gs_bin) is None:
        raise PdfAError(
            f"Ghostscript nicht gefunden ('{gs_bin}'). PDF/A-3 wird für "
            "ZUGFeRD-Rechnungen gefordert – bitte installieren: "
            "apt install ghostscript"
        )

    def_pfad = Path(pdfa_def) if pdfa_def else _finde_pdfa_def()
    if def_pfad is None or not def_pfad.is_file():
        raise PdfAError(
            "PDFA_def.ps (Ghostscript-Definitionsdatei für PDF/A) nicht "
            "gefunden. Bitte Ghostscript vollständig installieren "
            "(apt install ghostscript) oder in der Config unter "
            "[erechnung] pdfa_def = <pfad> setzen."
        )
    icc_pfad = Path(icc_profil) if icc_profil else _finde_datei(_ICC_MUSTER, _ICC_FALLBACK)
    if icc_pfad is None or not icc_pfad.is_file():
        raise PdfAError(
            "Kein sRGB-ICC-Profil gefunden (gs-iccprofiles oder "
            "/usr/share/color/icc/colord/sRGB.icc). Bitte [erechnung] "
            "icc_profil = <pfad> in der Config setzen."
        )

    try:
        def_inhalt = def_pfad.read_text(encoding="utf-8", errors="replace")
    except OSError as exc:
        raise PdfAError(f"PDFA_def.ps nicht lesbar {def_pfad}: {exc}") from exc
    def_gepatcht = _patche_pdfa_def(def_inhalt, icc_pfad)

    with tempfile.TemporaryDirectory(prefix="reu-pdfa3-") as tmp:
        tmp_pfad = Path(tmp)
        eingabe = tmp_pfad / "eingabe.pdf"
        ausgabe = tmp_pfad / "ausgabe.pdf"
        def_datei = tmp_pfad / "PDFA_def.ps"
        eingabe.write_bytes(pdf_bytes)
        def_datei.write_text(def_gepatcht, encoding="utf-8")

        argumente = [
            gs_bin,
            "-q",
            "--permit-file-read=" + str(icc_pfad),
            "-sDEVICE=pdfwrite",
            "-dPDFA=3",
            "-dPDFACompatibilityPolicy=1",
            "-sColorConversionStrategy=RGB",
            "-o" + str(ausgabe),
            str(def_datei),
            str(eingabe),
        ]
        try:
            proc = subprocess.run(
                argumente,
                capture_output=True,
                text=True,
                timeout=_GS_TIMEOUT,
            )
        except subprocess.TimeoutExpired as exc:
            raise PdfAError(
                f"Ghostscript (PDF/A-3) nach {_GS_TIMEOUT} s abgebrochen."
            ) from exc
        except OSError as exc:
            raise PdfAError(f"Ghostscript konnte nicht gestartet werden: {exc}") from exc

        if proc.returncode != 0 or not ausgabe.is_file() or ausgabe.stat().st_size == 0:
            fehler = (proc.stderr or proc.stdout or "").strip()
            if len(fehler) > 600:
                fehler = fehler[-600:]
            raise PdfAError(
                "PDF/A-3-Konvertierung fehlgeschlagen "
                f"(Ghostscript, Exit {proc.returncode}): {fehler}"
            )
        return ausgabe.read_bytes()
