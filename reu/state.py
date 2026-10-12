"""Persistenter State: Rechnungsliste im Blatt 'Rechnungen' der kunden.ods.

Statt der früheren JSON-Datei conf/<user>.state.json werden die
Rechnungen (Nummern UND Status) im Blatt 'Rechnungen' der
conf/<user>-kunden.ods gepflegt – dieselbe Datei, die auch Kunden und
Services enthält. Damit sind Nummernkreis und Journal direkt sichtbar
und pflegbar: der Status einer Zeile kann dort geändert werden
(z.B. 'offen' -> 'bezahlt'), und das --journal liest die Liste.

Schema des Blatts 'Rechnungen' (eine Zeile je Rechnung):
    re_nr | datum | kunde | leistungszeitraum | netto | ust | brutto | status

Spaltentitel sind groß-/kleinschreibungsagnostisch; zusätzliche Spalten
bleiben unberührt (neue Zeilen bekommen dort leere Zellen). Fehlt das
Blatt, legt REU es beim ersten echten Lauf mit der Kopfzeile an.

Die laufende Nummer wird aus der Liste ABGELEITET: nächste Nummer =
höchste vorhandene Nummer des Jahres + 1 (der Zähler kann dadurch nie
zurück- oder doppelt vergeben werden). Eine noch vorhandene Legacy-
Datei conf/<user>.state.json wird beim ersten echten Lauf einmalig
importiert und in <user>.state.json.importiert umbenannt.

Bei dry_run wird weder geschrieben noch eine echte Re-Nr vergeben.
"""
from __future__ import annotations

import datetime as _dt
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .ods import haenge_zeilen_an, lade_blatt_opt

BLATT_RECHNUNGEN = "Rechnungen"
KOPF = ["re_nr", "datum", "kunde", "leistungszeitraum", "netto", "ust", "brutto", "status"]
STATUS_OFFEN = "offen"
STATUS_ENTWURF = "entwurf"

# Re-Nr 'YYYY-NNN' (NNN auch länger, z.B. 2026-1234 ab 1000 Rechnungen)
_RENR = re.compile(r"^(\d{4})-(\d+)$")


class StateError(Exception):
    pass


@dataclass
class Buchung:
    renr: str
    kunde: str
    datum: str
    leistungszeitraum: str
    netto: str
    ust: str
    brutto: str
    status: str = STATUS_OFFEN

    @property
    def entwurf(self) -> bool:
        return self.status.strip().lower() == STATUS_ENTWURF


@dataclass
class State:
    buchungen: list[Buchung] = field(default_factory=list)
    # Höchste vergebene laufende Nummer je Jahr – aus den Buchungen
    # abgeleitet (bei der Migration auch aus dem Legacy-Zähler).
    max_nummern: dict[int, int] = field(default_factory=dict)
    letzte_rechnungsnummer: str = ""
    # Normalisierte (kleingeschriebene) Spaltenreihenfolge des Blatts –
    # neue Zeilen werden passend dazu geschrieben.
    spalten: list[str] = field(default_factory=lambda: list(KOPF))
    # Legacy-JSON-Buchungen, die noch nicht in die ODS geschrieben sind
    migration_offen: list[Buchung] = field(default_factory=list)


def _str(satz: dict[str, Any], schluessel: str, default: str = "") -> str:
    wert = satz.get(schluessel, default)
    if wert is None:
        return ""
    return str(wert).strip()


def _notiere_renr(state: State, renr: str) -> None:
    """Trägt eine Re-Nr in den je-Jahr-Merker ein (nur wenn höher)."""
    m = _RENR.match(renr.strip())
    if not m:
        return
    jahr, nummer = int(m.group(1)), int(m.group(2))
    if nummer > state.max_nummern.get(jahr, 0):
        state.max_nummern[jahr] = nummer
        state.letzte_rechnungsnummer = renr.strip()


def _buchung_als_zeile(b: Buchung, spalten: list[str]) -> list[str]:
    """Buchung als Zellwerte passend zur Spaltenreihenfolge des Blatts."""
    werte = {
        "re_nr": b.renr,
        "datum": b.datum,
        "kunde": b.kunde,
        "leistungszeitraum": b.leistungszeitraum,
        "netto": b.netto,
        "ust": b.ust,
        "brutto": b.brutto,
        "status": b.status,
    }
    return [werte.get(spalte, "") for spalte in spalten]


def lade_state(
    kunden_datei: str | Path,
    blatt: str = BLATT_RECHNUNGEN,
    *,
    legacy_json: str | Path | None = None,
) -> State:
    """Lädt die Rechnungsliste aus dem Blatt der kunden.ods.

    Fehlt das Blatt, ist der State leer (das Blatt wird beim ersten
    echten Lauf angelegt). legacy_json: frühere conf/<user>.state.json –
    wird einmalig importiert, solange das Blatt noch keine Buchungen
    enthält (geschrieben + umbenannt wird erst beim ersten echten Lauf
    über buche(); dry_run lässt die Datei unangetastet).
    """
    state = State()
    pfad = Path(kunden_datei)
    try:
        zeilen = lade_blatt_opt(pfad, blatt)
    except Exception as exc:  # OdsError, lesbar gemeldet
        raise StateError(f"Rechnungen-Blatt nicht lesbar {pfad}: {exc}") from exc

    if zeilen is not None:
        kopf: list[str] | None = None
        for zeile in zeilen:
            werte = ["" if z is None else z for z in zeile]
            if all(w == "" for w in werte):
                continue
            if kopf is None:
                kopf = [str(w).strip().lower() for w in werte]
                continue
            satz = {kopf[i]: werte[i] for i in range(min(len(kopf), len(werte)))}
            renr = _str(satz, "re_nr")
            if not renr:
                continue  # Zeile ohne Re-Nr (z.B. Leer-/Notizzeile)
            state.buchungen.append(
                Buchung(
                    renr=renr,
                    kunde=_str(satz, "kunde"),
                    datum=_str(satz, "datum"),
                    leistungszeitraum=_str(satz, "leistungszeitraum"),
                    netto=_str(satz, "netto"),
                    ust=_str(satz, "ust"),
                    brutto=_str(satz, "brutto"),
                    status=_str(satz, "status", STATUS_OFFEN) or STATUS_OFFEN,
                )
            )
        if kopf is None:
            state.spalten = list(KOPF)  # Blatt ohne Zeilen -> Standardschema
        else:
            if "re_nr" not in kopf:
                raise StateError(
                    f"Blatt '{blatt}' in {pfad} braucht eine Spalte 're_nr' "
                    f"(vorhanden: {', '.join(k for k in kopf if k) or 'keine'})."
                )
            state.spalten = kopf
        for b in state.buchungen:
            _notiere_renr(state, b.renr)

    # Legacy-JSON einmalig übernehmen, solange das Blatt leer ist
    if legacy_json is not None and not state.buchungen:
        legacy_pfad = Path(legacy_json)
        if legacy_pfad.is_file():
            try:
                roh = json.loads(legacy_pfad.read_text(encoding="utf-8"))
            except json.JSONDecodeError as exc:
                raise StateError(f"Legacy-State-Datei ungültig {legacy_pfad}: {exc}") from exc
            for eintrag in roh.get("buchungen", []):
                b = Buchung(
                    renr=str(eintrag.get("renr", "")),
                    kunde=str(eintrag.get("kunde", "")),
                    datum=str(eintrag.get("datum", "")),
                    leistungszeitraum=str(eintrag.get("leistungszeitraum", "")),
                    netto=str(eintrag.get("netto", "")),
                    ust=str(eintrag.get("ust", "")),
                    brutto=str(eintrag.get("brutto", "")),
                    status=str(eintrag.get("status", "")) or (
                        STATUS_ENTWURF if eintrag.get("entwurf") else STATUS_OFFEN
                    ),
                )
                if not b.renr:
                    continue
                state.migration_offen.append(b)
                state.buchungen.append(b)
                _notiere_renr(state, b.renr)
            # Zähler übernehmen, falls er über den Buchungen liegt
            try:
                jahr = int(roh.get("jahr") or 0)
                nummer = int(roh.get("laufende_nummer") or 0)
            except (TypeError, ValueError):
                jahr = nummer = 0
            if jahr and nummer and nummer > state.max_nummern.get(jahr, 0):
                state.max_nummern[jahr] = nummer
    return state


def naechste_renr(state: State, jahr: int, *, dry_run: bool, start: int = 1) -> str:
    """Erzeugt nächste Re-Nr 'YYYY-NNN' oder ENTWURF bei dry_run.

    dry_run: verwendet ENTWURF-YYYY-MM-DD-HHMMSS, schreibt nichts.
    Die laufende Nummer folgt aus der Rechnungsliste (höchste Nummer
    des Jahres + 1); start ist der konfigurierbare Startwert je Jahr
    ([erechnung] re_nr_start_<jahr>, z.B. 47 -> '2026-047'). Die
    Liste gewinnt, wenn sie schon weiter ist – die Nummer darf niemals
    zurück- oder doppelt vergeben werden.
    """
    if dry_run:
        jetzt = _dt.datetime.now().strftime("%Y-%m-%d-%H%M%S")
        return f"ENTWURF-{jetzt}"
    nummer = state.max_nummern.get(jahr, 0)
    if start > 1:
        nummer = max(nummer, start - 1)
    nummer += 1
    renr = f"{jahr}-{nummer:03d}"
    state.max_nummern[jahr] = nummer
    state.letzte_rechnungsnummer = renr
    return renr


def buche(
    state: State,
    *,
    renr: str,
    kunde: str,
    datum: _dt.date,
    leistungszeitraum: str,
    netto: Any,
    ust: Any,
    brutto: Any,
    entwurf: bool,
    dry_run: bool,
    kunden_datei: str | Path,
    blatt: str = BLATT_RECHNUNGEN,
    legacy_json: str | Path | None = None,
) -> None:
    """Fügt eine Buchung hinzu und persistiert im ODS-Blatt (außer dry_run).

    Bei anstehender Legacy-Migration werden die importierten Buchungen
    mitgeschrieben und die JSON-Datei in *.importiert umbenannt.
    """
    buchung = Buchung(
        renr=renr,
        kunde=kunde,
        datum=datum.isoformat(),
        leistungszeitraum=leistungszeitraum,
        netto=str(netto),
        ust=str(ust),
        brutto=str(brutto),
        status=STATUS_ENTWURF if entwurf else STATUS_OFFEN,
    )
    state.buchungen.append(buchung)
    if dry_run:
        return
    neue_zeilen = [_buchung_als_zeile(buchung, state.spalten)]
    if state.migration_offen:
        neue_zeilen = [
            _buchung_als_zeile(b, state.spalten) for b in state.migration_offen
        ] + neue_zeilen
    try:
        haenge_zeilen_an(kunden_datei, blatt, neue_zeilen, kopf=state.spalten)
    except Exception as exc:  # OdsError, lesbar gemeldet
        raise StateError(f"Rechnungen-Blatt nicht schreibbar {kunden_datei}: {exc}") from exc
    if state.migration_offen:
        state.migration_offen = []
        if legacy_json is not None:
            legacy_pfad = Path(legacy_json)
            if legacy_pfad.is_file():
                legacy_pfad.rename(legacy_pfad.with_name(legacy_pfad.name + ".importiert"))


def journal_zeilen(state: State) -> list[dict[str, str]]:
    """Liste von Dicts für die Rechnungsübersicht (Konsole/CSV)."""
    return [
        {
            "renr": b.renr,
            "kunde": b.kunde,
            "datum": b.datum,
            "leistungszeitraum": b.leistungszeitraum,
            "netto": b.netto,
            "ust": b.ust,
            "brutto": b.brutto,
            "status": b.status,
        }
        for b in state.buchungen
    ]
