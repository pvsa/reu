"""Gemeinsame Hilfsfunktionen und Konstanten für REU."""
from __future__ import annotations

import datetime as _dt
import re as _re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation

UST_SATZ = Decimal("0.19")
UST_PROZENT = Decimal("19.00")
WAehrUNG = "EUR"
UNIT_STUNDE = "HUR"
UNIT_STUECK = "C62"


class ZeitraumError(ValueError):
    """Ungültiger Zeitraum-Parameter (Monat/Bereich/Quartal)."""


@dataclass(frozen=True)
class Zeitraum:
    """Abrechnungszeitraum: ein oder mehrere Monate eines Jahres.

    Wird aus dem CLI-Argument gebildet: Einzelmonat (11), Monatsbereich
    (10-12) oder Quartal (Q4). Alle Monate liegen im selben Jahr.
    """
    jahr: int
    monate: tuple[int, ...]

    @property
    def erster_monat(self) -> int:
        return self.monate[0]

    @property
    def letzter_monat(self) -> int:
        return self.monate[-1]

    @property
    def einzelmonat(self) -> bool:
        return len(self.monate) == 1

    def _ist_quartal(self) -> bool:
        if len(self.monate) != 3:
            return False
        erste = self.erster_monat
        return (erste - 1) % 3 == 0 and self.monate == (erste, erste + 1, erste + 2)

    @property
    def iso(self) -> str:
        """Kompakt für State/Journal/Mail: '2025-11', '2025-Q4', '2025-10-2025-12'."""
        if self.einzelmonat:
            return f"{self.jahr:04d}-{self.erster_monat:02d}"
        if self._ist_quartal():
            q = (self.erster_monat - 1) // 3 + 1
            return f"{self.jahr:04d}-Q{q}"
        return (
            f"{self.jahr:04d}-{self.erster_monat:02d}"
            f"-{self.jahr:04d}-{self.letzter_monat:02d}"
        )

    @property
    def text(self) -> str:
        """Lesbar für Rechnung/PDF: '11/2025', 'Q4/2025', '10/2025-12/2025'."""
        if self.einzelmonat:
            return f"{self.erster_monat:02d}/{self.jahr}"
        if self._ist_quartal():
            q = (self.erster_monat - 1) // 3 + 1
            return f"Q{q}/{self.jahr}"
        return (
            f"{self.erster_monat:02d}/{self.jahr}-{self.letzter_monat:02d}/{self.jahr}"
        )

    @property
    def start_datum(self) -> _dt.date:
        return _dt.date(self.jahr, self.erster_monat, 1)

    @property
    def ende_datum(self) -> _dt.date:
        return letzter_des_monats(self.jahr, self.letzter_monat)


def _pruefe_monat(monat: int, arg: str) -> None:
    if not 1 <= monat <= 12:
        raise ZeitraumError(f"Ungültiger Monat in {arg!r}: {monat} (erlaubt 1-12).")


def parse_zeitraum(arg: str, jahr: int) -> Zeitraum:
    """Parst ein CLI-Zeitraum-Argument: '11', '10-12' oder 'Q4'/'q4'."""
    arg = (arg or "").strip()

    m = _re.fullmatch(r"[Qq]([1-4])", arg)
    if m:
        q = int(m.group(1))
        erste = 3 * q - 2
        return Zeitraum(jahr=jahr, monate=(erste, erste + 1, erste + 2))

    m = _re.fullmatch(r"(\d{1,2})-(\d{1,2})", arg)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        _pruefe_monat(a, arg)
        _pruefe_monat(b, arg)
        if a > b:
            raise ZeitraumError(
                f"Ungültiger Zeitraum {arg!r}: erster Monat liegt nach dem letzten."
            )
        return Zeitraum(jahr=jahr, monate=tuple(range(a, b + 1)))

    m = _re.fullmatch(r"(\d{1,2})", arg)
    if m:
        a = int(m.group(1))
        _pruefe_monat(a, arg)
        return Zeitraum(jahr=jahr, monate=(a,))

    raise ZeitraumError(
        f"Ungültiger Zeitraum {arg!r} – erlaubt: Monat (11), Bereich (10-12) oder Quartal (Q4)."
    )


def dezimal(wert: object, *, default: Decimal | None = None) -> Decimal:
    """Parst einen Wert zu Decimal; akzeptiert DE- und EN-Schreibweise.

    "42,00" und "42.00" werden beide zu Decimal('42.00').
    Leere Werte liefern 'default' (oder Decimal('0') wenn None).
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
    if "," in text and "." in text:
        # Beide Trenner: 1.234,56 (DE) bzw. 1,234.56 (EN) – der letzte gewinnt
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    # nur '.': EN-Dezimalpunkt, bleibt unverändert
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
