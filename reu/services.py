"""Service-Positionen aus dem (optionalen) Blatt 'Services' der Kunden-ODS.

Format des Blatts (eine Zeile je Kunde und Service):

    kunde | service | Kosten pro Stück [€] | 1 | 2 | 3 | … | 12

Die Spalten '1'–'12' stehen für die Monate des Jahres; der Zellwert ist
die Anzahl der Serviceeinheiten in diesem Monat (leer = 0).

Pro Service und Monat mit Menge > 0 entsteht eine eigene
Rechnungsposition (siehe positionen_fuer_zeitraum).
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Any

from .ods import OdsError, blatt_als_dicts
from .util import Zeitraum, dezimal

BLATT_DEFAULT = "Services"

_FLOAT_HEADER = re.compile(r"^\d+\.0$")  # numerischer Header, z.B. '1.0'


class ServicesError(Exception):
    """Fehler beim Einlesen der Service-Definitionen."""


@dataclass
class Service:
    kunde: str
    name: str
    kosten_pro_stueck: Decimal
    menge: dict[int, Decimal]  # Monat 1-12 -> Anzahl Serviceeinheiten


def _norm_satz(satz: dict[str, Any]) -> dict[str, Any]:
    """Normalisiert Header-Schlüssel ('1.0'/'01' -> '1')."""
    aus: dict[str, Any] = {}
    for schluessel, wert in satz.items():
        k = schluessel.strip()
        if _FLOAT_HEADER.match(k):
            k = k[:-2]
        elif len(k) == 2 and k.isdigit() and k.startswith("0"):
            k = str(int(k))
        aus[k] = wert
    return aus


def _kosten_spalte(satz: dict[str, Any]) -> str | None:
    """Findet die Spalte 'Kosten pro Stück [€]' ( Groß-/Kleinschreibung egal)."""
    for schluessel in satz:
        if schluessel.strip().lower().startswith("kosten"):
            return schluessel
    return None


def lade_services(
    datei: str | Path,
    blatt: str = BLATT_DEFAULT,
    *,
    bekannte_kunden: set[str] | None = None,
) -> dict[str, list[Service]]:
    """Liest das Blatt 'Services' und liefert {kunde: [Service, ...]}.

    Das Blatt ist OPTIONAL: Fehlt es in der ODS, gibt es keine Services
    ({}) – das ist kein Fehler. Fehlt die ODS selbst oder ist eine
    Service-Zeile unvollständig, wird ServicesError geworfen.
    """
    pfad = Path(datei)
    if not pfad.is_file():
        raise ServicesError(f"Kunden-ODS nicht gefunden: {pfad}")
    try:
        saetze = blatt_als_dicts(pfad, blatt)
    except OdsError as exc:
        if f"Blatt '{blatt}' fehlt" in str(exc):
            return {}  # Services-Blatt ist optional
        raise ServicesError(str(exc)) from exc

    services: dict[str, list[Service]] = {}
    for rohsatz in saetze:
        satz = _norm_satz(rohsatz)
        kunde = str(satz.get("kunde", "")).strip()
        name = str(satz.get("service", "")).strip()
        if not kunde and not name:
            continue  # komplett leere Zeile
        if not kunde or not name:
            raise ServicesError(
                f"Services-Zeile unvollständig (Kunde/Service fehlt): {rohsatz!r}"
            )
        if bekannte_kunden is not None and kunde not in bekannte_kunden:
            raise ServicesError(
                f"Service für unbekannten Kunden '{kunde}' – "
                f"in kunden.ods nicht definiert"
            )
        kosten_spalte = _kosten_spalte(satz)
        if not kosten_spalte:
            raise ServicesError(
                f"Spalte 'Kosten pro Stück [€]' fehlt für Service '{name}' (Kunde {kunde})"
            )
        kosten = dezimal(satz.get(kosten_spalte))
        if kosten <= 0:
            raise ServicesError(
                f"Kosten pro Stück für Service '{name}' (Kunde {kunde}) "
                f"fehlt oder ist <= 0: {satz.get(kosten_spalte)!r}"
            )
        menge: dict[int, Decimal] = {}
        for monat in range(1, 13):
            menge[monat] = dezimal(satz.get(str(monat)))
        services.setdefault(kunde, []).append(
            Service(kunde=kunde, name=name, kosten_pro_stueck=kosten, menge=menge)
        )
    return services


def positionen_fuer_zeitraum(
    services: list[Service], zeitraum: Zeitraum
) -> list[dict[str, Any]]:
    """Service-Positionen für den Zeitraum, nur Monate mit Menge > 0.

    Rückgabe: Liste von {service, monat, jahr, menge, einzelpreis}.
    """
    positionen: list[dict[str, Any]] = []
    for s in services:
        for monat in zeitraum.monate:
            anzahl = s.menge.get(monat, Decimal("0"))
            if anzahl > 0:
                positionen.append(
                    {
                        "service": s.name,
                        "monat": monat,
                        "jahr": zeitraum.jahr,
                        "menge": anzahl,
                        "einzelpreis": s.kosten_pro_stueck,
                    }
                )
    return positionen
