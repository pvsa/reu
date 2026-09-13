"""Gemeinsame Hilfsfunktionen und Konstanten für REU."""
from __future__ import annotations

import datetime as _dt
from decimal import Decimal, InvalidOperation

UST_SATZ = Decimal("0.19")
UST_PROZENT = Decimal("19.00")
WAehrUNG = "EUR"
UNIT_STUNDE = "HUR"
UNIT_STUECK = "C62"


def dezimal(wert: object, *, default: Decimal | None = None) -> Decimal:
    """Parst einen Wert zu Decimal; akzeptiert DE- und EN-Schreibweise.

    "42,00" und "42.00" werden beide zu Decimal('42.00').
    Leere Werte liefern ``default`` (oder Decimal('0') wenn None).
    """
    if wert is None:
        return default if default is not None else Decimal("0")
    if isinstance(wert, Decimal):
        return wert
    if isinstance(wert, (int, float)):
        return Decimal(str(wert))
    text = str(wert).strip()
    if not text:
        return default if default is not None else Decimal("0")
    text = text.replace(".", "").replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return default if default is not None else Decimal("0")


def euro(wert: Decimal) -> str:
    """Formatiert Decimal als DE-Währung '1.234,56'."""
    s = f"{wert:.2f}"
    sign = "-" if s.startswith("-") else ""
    if sign:
        s = s[1:]
    ganzzahl, _, nachkomma = s.partition(".")
    ganzzahl = f"{int(ganzzahl):,}".replace(",", ".")
    return f"{sign}{ganzzahl},{nachkomma}"


def menge(wert: Decimal, nachkomma: int = 2) -> str:
    """Formatiert eine Menge DE: '42,50'."""
    s = f"{wert:.{nachkomma}f}"
    ganzzahl, _, nk = s.partition(".")
    return f"{int(ganzzahl):,}".replace(",", ".") + (f",{nk}" if nk else ",00")


def jahr_monat_text(year: int, month: int) -> str:
    """'11/2025'."""
    return f"{month:02d}/{year}"


def leistungszeitraum_iso(year: int, month: int) -> str:
    """'2025-11'."""
    return f"{year:04d}-{month:02d}"


def letzter_des_monats(year: int, month: int) -> _dt.date:
    naechster = _dt.date(year + (month // 12), (month % 12) + 1, 1)
    return naechster - _dt.timedelta(days=1)


def date_to_iso(d: object) -> str | None:
    if d is None:
        return None
    if isinstance(d, _dt.datetime):
        return d.date().isoformat()
    if isinstance(d, _dt.date):
        return d.isoformat()
    return str(d)
