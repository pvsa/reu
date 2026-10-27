"""Prüfgate --valide: erzeugte ZUGFeRD-PDFs extern validieren.

Die PDF/A-3b-Konvertierung (reu/pdfa.py) macht das Rechnungs-PDF
konform, und das eingebettete Faktur-X-XML wird beim Erzeugen bereits
gegen das EN-16931-XSD geprüft (drafthorse + lxml). Als Release-Gate
vor dem Versand gehören aber UNABHÄNGIGE Prüfungen mit externen
Referenz-Werkzeugen dazu:

- **veraPDF** – PDF/A-3b-Konformität des PDFs (ISO 19005-3): eingebettete
  Fonts, Farben/ICC-OutputIntent, XMP-Metadaten, Objektregeln.
- **Mustang** – Faktur-X gesamt: EN-16931-XML + ZUGFeRD-Einbettung
  (AF-Struktur, Metadaten).

CLI: ``run-reu.py <user> --valide [PDF ...]`` – ohne Pfad-Argumente
werden alle ``*.pdf`` aus ``out/final`` geprüft; Argumente können
einzelne Dateien oder Verzeichnisse sein (z.B. ``out/dry-run``).

Tool-Suche je Prüfung (lokal vor Docker):
- veraPDF: ``verapdf`` im PATH, sonst Docker-Image ``verapdf/cli``.
- Mustang: JAR via Umgebungsvariable ``REU_MUSTANG_JAR`` oder
  ``Mustang-CLI*.jar`` im Arbeitsverzeichnis (mit ``java``), sonst
  Docker-Image ``mustangproject/cli``.

Fehlt ein Werkzeug, wird die Prüfung mit Hinweis übersprungen (und
nicht als bestanden gewertet); fehlen ALLE, endet das Gate mit Exit 1.

Exit-Codes: **0** = alle durchgeführten Prüfungen bestanden,
**1** = mindestens eine Prüfung fehlgeschlagen, nichts zu prüfen oder
kein Prüfwerkzeug verfügbar.
"""
from __future__ import annotations

import os
import shutil
import subprocess
from pathlib import Path

from .config import Config

TIMEOUT_SEKUNDEN = 180
VERAPDF_DOCKER_IMAGE = "verapdf/cli"
MUSTANG_DOCKER_IMAGE = "mustangproject/cli"


class ValidierError(Exception):
    pass


def _welches(programm: str) -> str:
    return shutil.which(programm) or ""


def _mustang_jar() -> Path | None:
    """Sucht die Mustang-CLI-JAR: env REU_MUSTANG_JAR, dann Arbeitsverzeichnis."""
    env = os.environ.get("REU_MUSTANG_JAR", "").strip()
    if env:
        pfad = Path(env)
        if pfad.is_file():
            return pfad
        raise ValidierError(f"REU_MUSTANG_JAR zeigt auf keine Datei: {pfad}")
    for kandidat in sorted(Path.cwd().glob("Mustang-CLI*.jar")):
        return kandidat
    return None


def _fuehre_aus(cmd: list[str]) -> tuple[int, str]:
    """Startet ein Prüfwerkzeug, liefert (Exit-Code, kombinierte Ausgabe)."""
    try:
        erg = subprocess.run(
            cmd, capture_output=True, text=True, timeout=TIMEOUT_SEKUNDEN
        )
    except subprocess.TimeoutExpired as exc:
        raise ValidierError(
            f"Prüfwerkzeug nach {TIMEOUT_SEKUNDEN} s abgebrochen: {' '.join(cmd[:4])} …"
        ) from exc
    except FileNotFoundError as exc:
        raise ValidierError(f"Prüfwerkzeug nicht gefunden: {exc}") from exc
    ausgabe = "\n".join(
        teil.strip() for teil in (erg.stdout, erg.stderr) if teil and teil.strip()
    )
    return erg.returncode, ausgabe


def _docker_cmd(image: str, pdf: Path, *werkzeug_args: str) -> list[str]:
    """docker run mit eingehängtem Elternverzeichnis der PDF."""
    return [
        "docker", "run", "--rm",
        "-v", f"{pdf.resolve().parent}:/data",
        image, *werkzeug_args, f"/data/{pdf.name}",
    ]


def pruefe_verapdf(pdf: Path, *, verapdf_bin: str, docker_bin: str) -> tuple[bool, str, str] | None:
    """veraPDF-Prüfung (PDF/A-3b). None, wenn kein veraPDF verfügbar ist."""
    if verapdf_bin:
        rc, ausgabe = _fuehre_aus([verapdf_bin, "-f", "3b", str(pdf.resolve())])
        return rc == 0, ausgabe, "veraPDF (PDF/A-3b, lokal)"
    if docker_bin:
        rc, ausgabe = _fuehre_aus(_docker_cmd(VERAPDF_DOCKER_IMAGE, pdf, "-f", "3b"))
        return rc == 0, ausgabe, f"veraPDF (PDF/A-3b, Docker {VERAPDF_DOCKER_IMAGE})"
    return None


def pruefe_mustang(
    pdf: Path, *, jar: Path | None, java_bin: str, docker_bin: str
) -> tuple[bool, str, str] | None:
    """Mustang-Prüfung (Faktur-X). None, wenn kein Mustang verfügbar ist."""
    if jar is not None:
        if not java_bin:
            return None
        rc, ausgabe = _fuehre_aus([
            java_bin, "-jar", str(jar.resolve()),
            "--action", "validate", "--source", str(pdf.resolve()),
        ])
        return rc == 0, ausgabe, "Mustang (Faktur-X, lokal)"
    if docker_bin:
        rc, ausgabe = _fuehre_aus(_docker_cmd(MUSTANG_DOCKER_IMAGE, pdf, "validate"))
        return rc == 0, ausgabe, f"Mustang (Faktur-X, Docker {MUSTANG_DOCKER_IMAGE})"
    return None


def pdf_pfade(argumente: list[str], standard_verz: Path) -> list[Path]:
    """Löst Pfad-Argumente zu PDF-Dateien auf (Dateien oder Verzeichnisse).

    Ohne Argumente: alle *.pdf aus standard_verz (sortiert).
    """
    if not argumente:
        pfade = sorted(standard_verz.glob("*.pdf")) if standard_verz.is_dir() else []
        if not pfade:
            raise ValidierError(
                f"Keine PDFs in {standard_verz} gefunden – Pfade als Argument "
                "angeben oder erst Rechnungen erzeugen."
            )
        return pfade
    pfade: list[Path] = []
    for arg in argumente:
        pfad = Path(arg).expanduser()
        if pfad.is_dir():
            pfade.extend(sorted(pfad.glob("*.pdf")))
        elif pfad.is_file():
            if pfad.suffix.lower() != ".pdf":
                raise ValidierError(f"Keine PDF-Datei: {pfad}")
            pfade.append(pfad)
        else:
            raise ValidierError(f"Nicht gefunden: {pfad}")
    if not pfade:
        raise ValidierError("Keine PDFs in den angegebenen Pfaden.")
    return pfade


def _kuerze(text: str, grenze: int = 600) -> str:
    text = text.strip()
    if len(text) <= grenze:
        return text
    return text[:grenze] + f" … (+{len(text) - grenze} Zeichen)"


def cmd_valide(cfg: Config, argumente: list[str]) -> int:
    """Prüft PDFs mit allen verfügbaren Werkzeugen, gibt Exit-Code zurück."""
    standard_verz = cfg.ausgabe_verz / "final"
    pfade = pdf_pfade(argumente, standard_verz)

    verapdf_bin = _welches("verapdf")
    docker_bin = _welches("docker")
    java_bin = _welches("java")
    jar = _mustang_jar()

    verapdf_verfuegbar = bool(verapdf_bin or docker_bin)
    mustang_verfuegbar = bool((jar is not None and java_bin) or docker_bin)

    hinweise: list[str] = []
    if not verapdf_verfuegbar:
        hinweise.append(
            "Hinweis: veraPDF nicht gefunden (übersprungen) – Installation: "
            "README-Abschnitt 'Validierung der ZUGFeRD-PDF'."
        )
    if jar is None and not docker_bin:
        hinweise.append(
            "Hinweis: Mustang nicht gefunden (übersprungen) – REU_MUSTANG_JAR "
            "setzen, Mustang-CLI.jar ins Projektverzeichnis legen oder Docker "
            f"({MUSTANG_DOCKER_IMAGE}) nutzen."
        )
    if jar is not None and not java_bin and not docker_bin:
        hinweise.append(
            "Hinweis: Mustang-JAR gefunden, aber 'java' fehlt (übersprungen) – "
            "apt install default-jre."
        )
    for h in hinweise:
        print(h)
    if not verapdf_verfuegbar and not mustang_verfuegbar:
        print("FEHLER: kein Prüfwerkzeug verfügbar – siehe Hinweise oben.")
        return 1

    print(f"Validiere {len(pfade)} PDF(s):")
    gesamt = 0
    fehlgeschlagen = 0
    for pdf in pfade:
        print(f"\n{pdf.name}")
        for pruefung in (
            pruefe_verapdf(pdf, verapdf_bin=verapdf_bin, docker_bin=docker_bin),
            pruefe_mustang(pdf, jar=jar, java_bin=java_bin, docker_bin=docker_bin),
        ):
            if pruefung is None:
                continue
            ok, ausgabe, titel = pruefung
            gesamt += 1
            if ok:
                print(f"  OK   {titel}")
            else:
                fehlgeschlagen += 1
                print(f"  FAIL {titel}")
                if ausgabe:
                    print(f"       {_kuerze(ausgabe)}")

    if gesamt == 0:
        print("FEHLER: keine Prüfung durchgeführt – siehe Hinweise oben.")
        return 1
    print(f"\nErgebnis: {gesamt - fehlgeschlagen}/{gesamt} Prüfung(en) bestanden.")
    if fehlgeschlagen:
        print("FEHLER: mindestens eine Prüfung fehlgeschlagen – nicht versenden.")
        return 1
    return 0


__all__ = [
    "TIMEOUT_SEKUNDEN",
    "ValidierError",
    "cmd_valide",
    "pdf_pfade",
    "pruefe_mustang",
    "pruefe_verapdf",
]
