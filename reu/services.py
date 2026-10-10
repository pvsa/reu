"""Service-Positionen aus dem (optionalen) Blatt 'Services' der Kunden-ODS.

Format des Blatts (eine Zeile je Kunde und Service):

    kunde | service | Kosten pro Stück [€] | 1 | 2 | 3 | … | 12

Die Spalten '1'–'12' stehen für die Monate des Jahres. Zellwert je Monat:
- Zahl (z.B. '2' oder 2) = Anzahl der Serviceeinheiten in diesem Monat
- Markierung ('x', 'ja', 'yes', 'j', 'y', 'ok', 'wahr', 'true', '✓' oder
  Wahrheitszelle) = 1 Einheit (Service in diesem Monat erbracht/markiert)
- leer = 0
Spaltentitel sind groß-/kleinschreibungsagnostisch ('Kunde' wie 'kunde').
Andere, nicht leere Zellwerte sind ein FEHLER (kein stillschweigendes
Übergehen – sonst fehlt das Gewerk unbemerkt auf der Rechnung).

Pro Service und Monat mit Menge > 0 entsteht eine eigene
Rechnungsposition (siehe positionen_fuer_zeitraum).
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
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
    """Normalisiert Header-Schlüssel: 'Kunde'->'kunde', '1.0'/'01' -> '1'.

    Groß-/Kleinschreibung der Spaltentitel ist egal.
    """
    aus: dict[str, Any] = {}
    for schluessel, wert in satz.items():
        k = schluessel.strip().lower()
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


# Zellwerte, die als Markierung gelten (je 1 Serviceeinheit im Monat);
# Groß-/Kleinschreibung egal.
_MARKIERUNGEN = {"x", "ja", "yes", "j", "y", "ok", "wahr", "true", "✓", "✔"}


def _zahl_aus_text(text: str) -> Decimal | None:
    """Parst eine Zahl in DE- oder EN-Schreibweise; None wenn keine Zahl.

    (dezimal() liefert bei unparsebarem Text Decimal('0') und kann
    'keine Zahl' daher nicht von einer echten 0 unterscheiden.)
    """
    t = text
    if "," in t and "." in t:
        if t.rfind(",") > t.rfind("."):
            t = t.replace(".", "").replace(",", ".")
        else:
            t = t.replace(",", "")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return Decimal(t)
    except InvalidOperation:
        return None


def _menge_aus_zelle(wert: object) -> Decimal:
    """Menge einer Monatszelle: Zahl = Anzahl, Markierung = 1, leer = 0.

    Nicht leere, aber weder als Zahl noch als Markierung lesbare Werte
    werfen ServicesError – ein Gewerk darf nie stillschweigend als
    Menge 0 von der Rechnung verschwinden.
    """
    if wert is None:
        return Decimal("0")
    if isinstance(wert, bool):
        return Decimal("1") if wert else Decimal("0")
    if isinstance(wert, (int, float, Decimal)):
        zahl = dezimal(wert)
        if zahl < 0:
            raise ServicesError(
                f"Negative Anzahl in einer Monats-Spalte des Services-Blatts: {wert!r}"
            )
        return zahl
    text = str(wert).strip()
    if not text:
        return Decimal("0")
    if text.lower() in _MARKIERUNGEN:
        return Decimal("1")
    zahl = _zahl_aus_text(text)
    if zahl is not None and zahl >= 0:
        return zahl
    raise ServicesError(
        f"Ungültiger Wert {wert!r} in einer Monats-Spalte des Services-Blatts "
        "(erlaubt: Anzahl als Zahl, Markierung wie 'x'/'ja'/'✓' = 1 Einheit, "
        "leer = 0)"
    )


def lade_services(
    datei: str | Path,
    blatt: str = BLATT_DEFAULT,
    *,
    bekannte_kunden: set[str] | None = None,
) -> dict[str, list[Service]]:
    """Liest das Blatt 'Services' und liefert {kunde: [Service, ...]}.

    Das Blatt ist OPTIONAL: Fehlt es in der ODS, gibt es keine Services
    ({}) – das ist kein Fehler. Fehlt die ODS selbst, ist eine
    Service-Zeile unvollständig oder werden die Spalten 'kunde'/'service'
    nicht erkannt, wird ServicesError geworfen. Spaltentitel sind
    groß-/kleinschreibungsagnostisch ('Kunde' wie 'kunde').
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
            if any(
                v is not None and str(v).strip() for v in rohsatz.values()
            ):
                # Zeile hat Werte, aber kunde/service sind leer oder die
                # Spalten werden nicht erkannt – nicht stillschweigend skippen!
                raise ServicesError(
                    "Services-Zeile enthält Werte, aber 'kunde'/'service' sind "
                    f"leer oder die Spaltennamen werden nicht erkannt: {rohsatz!r} "
                    "(erwartet: kunde, service, Kosten pro Stück [€], 1–12)"
                )
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
        if not any(str(m) in satz for m in range(1, 13)):
            raise ServicesError(
                "Monats-Spalten '1'-'12' fehlen im Services-Blatt für "
                f"Service '{name}' (Kunde {kunde}) – Zeile: {rohsatz!r}"
            )
        menge: dict[int, Decimal] = {}
        for monat in range(1, 13):
            menge[monat] = _menge_aus_zelle(satz.get(str(monat)))
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
