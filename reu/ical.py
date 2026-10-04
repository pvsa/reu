"""Stufe 1: Stunden aus iCal-Feed.

Lädt den iCal-Download, parst VEVENTs im angefragten Zeitraum (Einzelmonat,
Monatsbereich oder Quartal) und gruppiert nach Kundenkürzel (aus dem
SUMMARY-Präfix 'ABC: …' oder der CATEGORIES-Eigenschaft). Gibt ein
dict {kunde: [Einzeltermine]} zurück.

Format eines Eintrags:
    {"datum": date, "dauer_h": Decimal, "beschreibung": str}
"""
from __future__ import annotations

import datetime as _dt
import re
import sys
from decimal import Decimal
from pathlib import Path
from typing import Any

import requests
from icalendar import Calendar
import pytz

from .config import ICalConf
from .util import Zeitraum

_KUNDE_PREFIX = re.compile(r"^\s*([A-Z0-9_-]{1,10})\s*[:\-]\s*(.*)$")


class ICalError(Exception):
    pass


def _lade_ical(conf: ICalConf) -> bytes:
    url = conf.url
    if url.startswith("file://"):
        pfad = url[len("file://"):]
        try:
            return Path(pfad).read_bytes()
        except OSError as exc:
            raise ICalError(f"Lokale iCal-Datei nicht lesbar {pfad}: {exc}") from exc
    auth = None
    if conf.username:
        auth = (conf.username, conf.password or "")
    headers = {"Accept": "text/calendar"}
    letzter_fehler: Exception | None = None
    for versuch in range(1, conf.max_retries + 1):
        try:
            resp = requests.get(
                conf.url,
                auth=auth,
                headers=headers,
                timeout=conf.timeout,
            )
            resp.raise_for_status()
            return resp.content
        except requests.RequestException as exc:  # noqa: PERF203
            letzter_fehler = exc
            if versuch < conf.max_retries:
                import time

                time.sleep(conf.retry_delay)
    raise ICalError(
        f"iCal-Download nach {conf.max_retries} Versuchen fehlgeschlagen: {letzter_fehler}"
    ) from letzter_fehler


def _kunde_aus_event(event: Any) -> tuple[str, str] | None:
    """Kürzel aus SUMMARY-Präfix oder CATEGORIES.

    Liefert (kuerzel, beschreibung) oder None, wenn der Termin kein
    Kürzel trägt – solche Termine sind nicht abrechenbar (interne/
    technische Einträge wie Backups) und werden übersprungen.
    """
    summary = str(event.get("summary") or "")
    m = _KUNDE_PREFIX.match(summary)
    if m:
        return m.group(1), (m.group(2).strip() or summary)
    cat = event.get("categories")
    if cat:
        cats = [str(c) for c in cat.cats] if hasattr(cat, "cats") else [str(cat)]
        if cats:
            return cats[0].strip(), summary
    return None


def _dauer_stunden(start: _dt.datetime, end: _dt.datetime) -> Decimal:
    delta = end - start
    sekunden = delta.total_seconds()
    stunden = Decimal(str(sekunden)) / Decimal("3600")
    return stunden.quantize(Decimal("0.01"))


def _zeitraum_bereich(zeitraum: Zeitraum) -> tuple[_dt.datetime, _dt.datetime]:
    """UTC-Bereich: erster Tag des ersten Monats bis exklusiv erster Tag
    des Folgemonats des letzten Monats."""
    tz = pytz.UTC
    start = tz.localize(_dt.datetime(zeitraum.jahr, zeitraum.erster_monat, 1))
    if zeitraum.letzter_monat == 12:
        naechstes_jahr, naechster_monat = zeitraum.jahr + 1, 1
    else:
        naechstes_jahr, naechster_monat = zeitraum.jahr, zeitraum.letzter_monat + 1
    end = tz.localize(_dt.datetime(naechstes_jahr, naechster_monat, 1))
    return start, end


def lade_stunden(conf: ICalConf, zeitraum: Zeitraum) -> dict[str, list[dict[str, Any]]]:
    """Lädt iCal und liefert {kunde: [ {datum, dauer_h, beschreibung} ]}.

    Termine außerhalb des Zeitraums werden ignoriert. Termine ohne Dauer
    (z.B. Ganztages) werden als 0,00 h erfasst und entsprechend ausgewiesen.
    """
    roh = _lade_ical(conf)
    try:
        cal = Calendar.from_ical(roh)
    except Exception as exc:  # noqa: BLE001
        raise ICalError(f"iCal konnte nicht geparsed werden: {exc}") from exc

    start, end = _zeitraum_bereich(zeitraum)
    ergebnis: dict[str, list[dict[str, Any]]] = {}
    ohne_kuerzel: list[str] = []

    for event in cal.walk("VEVENT"):
        dtstart = event.get("dtstart")
        dtend = event.get("dtend")
        if dtstart is None or dtend is None:
            continue
        vstart = dtstart.dt
        vend = dtend.dt
        # datetime/date vereinheitlichen
        if isinstance(vstart, _dt.datetime):
            ev_start = vstart if vstart.tzinfo else pytz.UTC.localize(vstart)
        else:
            ev_start = _dt.datetime.combine(vstart, _dt.time.min, tzinfo=pytz.UTC)
        if isinstance(vend, _dt.datetime):
            ev_end = vend if vend.tzinfo else pytz.UTC.localize(vend)
        else:
            ev_end = _dt.datetime.combine(vend, _dt.time.max, tzinfo=pytz.UTC)
        if ev_end <= ev_start:
            continue
        if not (ev_start < end and ev_end > start):
            continue
        kunde_aus_event = _kunde_aus_event(event)
        if kunde_aus_event is None:
            summary = str(event.get("summary") or "").strip()
            ohne_kuerzel.append(
                f"{ev_start.astimezone(pytz.UTC).date()}: {summary or '(ohne SUMMARY)'}"
            )
            continue
        kunde, beschreibung = kunde_aus_event
        dauer = _dauer_stunden(ev_start, ev_end)
        ergebnis.setdefault(kunde, []).append(
            {
                "datum": ev_start.astimezone(pytz.UTC).date(),
                "dauer_h": dauer,
                "beschreibung": beschreibung or str(event.get("summary") or ""),
            }
        )

    # Hinweis auf übersprungene, nicht abrechenbare Termine
    if ohne_kuerzel:
        print(
            f"Hinweis: {len(ohne_kuerzel)} Termin(e) ohne Kundenkürzel übersprungen "
            "(nicht abrechenbar, Präfix 'ABC: …' fehlt):",
            file=sys.stderr,
        )
        for eintrag in ohne_kuerzel[:10]:
            print(f"  {eintrag}", file=sys.stderr)
        if len(ohne_kuerzel) > 10:
            print(f"  … und {len(ohne_kuerzel) - 10} weitere", file=sys.stderr)

    # sortieren nach Datum
    for kunde in ergebnis:
        ergebnis[kunde].sort(key=lambda e: e["datum"])
    return ergebnis


def stunden_summe(kundentermine: list[dict[str, Any]]) -> Decimal:
    return sum((e["dauer_h"] for e in kundentermine), Decimal("0"))
