"""Persistenter State: Rechnungsnummern-Kreis und Journalbuchungen.

Format JSON: conf/<user>.state.json. Bei dry_run wird die Datei
NICHT verändert und es wird keine echte Re-Nr vergeben.
"""
from __future__ import annotations

import datetime as _dt
import json
from dataclasses import dataclass, field, asdict
from decimal import Decimal
from pathlib import Path
from typing import Any


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
    status: str = "offen"
    entwurf: bool = False


@dataclass
class State:
    letzte_rechnungsnummer: str = ""
    jahr: int = 0
    laufende_nummer: int = 0
    buchungen: list[Buchung] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "letzte_rechnungsnummer": self.letzte_rechnungsnummer,
            "jahr": self.jahr,
            "laufende_nummer": self.laufende_nummer,
            "buchungen": [asdict(b) for b in self.buchungen],
        }


def _dec_str(x: Any) -> str:
    if isinstance(x, Decimal):
        return str(x)
    return str(x)


def lade_state(pfad: Path) -> State:
    if not pfad.is_file():
        return State()
    try:
        roh = json.loads(pfad.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise StateError(f"State-Datei ungültig {pfad}: {exc}") from exc
    buchungen = [Buchung(**b) for b in roh.get("buchungen", [])]
    return State(
        letzte_rechnungsnummer=roh.get("letzte_rechnungsnummer", ""),
        jahr=roh.get("jahr", 0),
        laufende_nummer=roh.get("laufende_nummer", 0),
        buchungen=buchungen,
    )


def _speichere_state(state: State, pfad: Path) -> None:
    pfad.parent.mkdir(parents=True, exist_ok=True)
    pfad.write_text(
        json.dumps(state.to_dict(), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )


def naechste_renr(state: State, jahr: int, *, dry_run: bool) -> str:
    """Erzeugt nächste Re-Nr 'YYYY-NNN' oder ENTWURF bei dry_run.

    dry_run: verwendet ENTWURF-YYYY-MM-DD-HHMMSS, schreibt nichts.
    """
    if dry_run:
        jetzt = _dt.datetime.now().strftime("%Y-%m-%d-%H%M%S")
        return f"ENTWURF-{jetzt}"
    if state.jahr != jahr:
        state.jahr = jahr
        state.laufende_nummer = 0
    state.laufende_nummer += 1
    renr = f"{jahr}-{state.laufende_nummer:03d}"
    state.letzte_rechnungsnummer = renr
    return renr


def buche(
    state: State,
    *,
    renr: str,
    kunde: str,
    datum: _dt.date,
    leistungszeitraum: str,
    netto: Decimal,
    ust: Decimal,
    brutto: Decimal,
    entwurf: bool,
    dry_run: bool,
    pfad: Path,
) -> None:
    """Fügt eine Buchung hinzu und persistiert (außer bei dry_run)."""
    state.buchungen.append(
        Buchung(
            renr=renr,
            kunde=kunde,
            datum=datum.isoformat(),
            leistungszeitraum=leistungszeitraum,
            netto=_dec_str(netto),
            ust=_dec_str(ust),
            brutto=_dec_str(brutto),
            status="entwurf" if entwurf else "offen",
            entwurf=entwurf,
        )
    )
    if dry_run:
        return
    _speichere_state(state, pfad)


def journal_zeilen(state: State) -> list[dict[str, str]]:
    """Liste von Dicts für die Rechnungsübersicht (Konsole/CSV)."""
    return [asdict(b) for b in state.buchungen]
