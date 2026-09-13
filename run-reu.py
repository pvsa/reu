#!/usr/bin/env python3
"""REU – CLI-Dispatch.

Aufruf:
    ./run-reu.py <user> <month> <year> --hours
    ./run-reu.py <user> <month> <year> --expenses
    ./run-reu.py <user> <month> <year> --invoice
    ./run-reu.py <user> <month> <year> --full [--dry-run]
    ./run-reu.py <user> --journal <year>

Ohne Action-Flag wird nur die Hilfe angezeigt (keine Ausführung).
"""
from __future__ import annotations

import argparse
import datetime as _dt
import sys
from decimal import Decimal
from pathlib import Path

from reu import (
    config as config_mod,
    expenses as expenses_mod,
    ical as ical_mod,
    invoice as invoice_mod,
    kunden as kunden_mod,
    smtp as smtp_mod,
    state as state_mod,
    zugferd as zugferd_mod,
)
from reu.config import Config
from reu.kunden import Kunde
from reu.util import euro, leistungszeitraum_iso, letzter_des_monats


def _ausgabe_verzeichnis(cfg: Config, dry_run: bool) -> Path:
    verz = cfg.ausgabe_verz / ("dry-run" if dry_run else "final")
    verz.mkdir(parents=True, exist_ok=True)
    return verz


def _lade_kunden(cfg: Config) -> dict[str, Kunde]:
    return kunden_mod.lade_kunden(cfg.kunden.datei, cfg.kunden.blatt)


def _drucke_tabelle(zeilen: list[list[str]], kopf: list[str]) -> None:
    breiten = [len(kopf)]
    matrix = [kopf] + zeilen
    spalten = len(kopf)
    breiten = [max(len(str(r[i])) for r in matrix if i < len(r)) for i in range(spalten)]
    def fmt(r):
        return "  ".join(str(c).ljust(breiten[i]) for i, c in enumerate(r[:spalten]))
    print(fmt(kopf))
    print("-" * (sum(breiten) + 2 * (spalten - 1)))
    for r in zeilen:
        print(fmt(r))


def cmd_hours(cfg: Config, month: int, year: int) -> int:
    kunden = _lade_kunden(cfg)
    stunden = ical_mod.lade_stunden(cfg.ical, year, month)
    if not stunden:
        print(f"Keine Termine im Leistungszeitraum {leistungszeitraum_iso(year, month)} gefunden.")
        return 0
    print(f"Stunden {leistungszeitraum_iso(year, month)}:")
    for k in sorted(stunden):
        kunde = kunden.get(k)
        name = kunde.name if kunde else "(unbekannt)"
        total = ical_mod.stunden_summe(stunden[k])
        print(f"\n=== {k} – {name} ({euro(kunde.stundensatz)}/h) ===")
        for e in stunden[k]:
            print(f"  {e['datum'].isoformat()}  {e['dauer_h']:>6} h  {e['beschreibung']}")
        print(f"  Summe: {total} h  → Netto {euro(total * kunde.stundensatz) if kunde else '?'}")
    return 0


def cmd_expenses(cfg: Config, month: int, year: int) -> int:
    kunden = _lade_kunden(cfg)
    erg = expenses_mod.lade_auslagen(cfg.auslagen, year, month, set(kunden))
    print(f"Auslagen {leistungszeitraum_iso(year, month)} – Freigabe: {'JA' if erg['freigegeben'] else 'NEIN (Entwurf)'}")
    for k in sorted(erg["auslagen"]):
        kunde = kunden.get(k, None)
        name = kunde.name if kunde else "(unbekannt)"
        zeilen = erg["auslagen"][k]
        print(f"\n=== {k} – {name} ===")
        for z in zeilen:
            print(f"  {z['datum']}  {z['belegnr']:>10}  {z['art']:<12}  {z['bezeichnung']:<25}  {euro(z['betrag_netto'])}")
        print(f"  Summe Netto: {euro(expenses_mod.auslagen_summe(zeilen))}")
    return 0


def _erzeuge_rechnung(
    cfg: Config,
    kunde: Kunde,
    stunden: list[dict],
    auslagen: list[dict],
    month: int,
    year: int,
    state: state_mod.State,
    dry_run: bool,
    freigegeben: bool,
) -> dict:
    re_datum = letzter_des_monats(year, month)
    faelligkeit = re_datum + _dt.timedelta(days=cfg.erechnung.zahlungsziel_tage)
    # entwurf = keine finale Rechnung (kein XML im PDF-Hinweis, keine echte Re-Nr).
    # Im Dry-Run wird trotzdem das ZUGFeRD-XML eingebettet, wenn Meta=yes,
    # damit die PDF inkl. XML geprüft werden kann – aber mit ENTWURF-Re-Nr.
    entwurf = not freigegeben
    renr = state_mod.naechste_renr(state, year, dry_run=dry_run)

    out_verz = _ausgabe_verzeichnis(cfg, dry_run)
    basis_name = f"{kunde.kunde}_Rechnung_{renr}.pdf"
    pdf_pfad = out_verz / basis_name

    summen = invoice_mod.erzeuge_rechnung_pdf(
        cfg=cfg,
        kunde=kunde,
        stunden=stunden,
        auslagen=auslagen,
        year=year,
        month=month,
        renr=renr,
        re_datum=re_datum,
        faelligkeit=faelligkeit,
        entwurf=entwurf,
        freigegeben=freigegeben,
        ausgabe_pfad=pdf_pfad,
    )

    final_pdf = pdf_pfad
    if freigegeben:
        # ZUGFeRD-XML einbetten (auch im Dry-Run, damit die PDF prüfbar ist)
        final_pdf = out_verz / f"{kunde.kunde}_Rechnung_{renr}_ZUGFeRD.pdf"
        zugferd_mod.erzeuge_zugferd_pdf(
            pdf_pfad=pdf_pfad,
            ausgabe_pfad=final_pdf,
            cfg=cfg,
            kunde=kunde,
            stunden=stunden,
            auslagen=auslagen,
            summen=summen,
            renr=renr,
            re_datum=re_datum,
            faelligkeit=faelligkeit,
            year=year,
            month=month,
        )
        pdf_pfad.unlink(missing_ok=True)
    else:
        # Entwurf (Meta!A1 != yes): keine ZUGFeRD-Einbettung
        pass

    state_mod.buche(
        state,
        renr=renr,
        kunde=kunde.kunde,
        datum=re_datum,
        leistungszeitraum=leistungszeitraum_iso(year, month),
        netto=summen["netto"],
        ust=summen["ust"],
        brutto=summen["brutto"],
        entwurf=entwurf,
        dry_run=dry_run,
        pfad=cfg.state_pfad,
    )
    return {"kunde": kunde.kunde, "renr": renr, "pdf": str(final_pdf), "summen": summen, "entwurf": entwurf}


def cmd_invoice(cfg: Config, month: int, year: int, *, dry_run: bool, nur_kunde: str | None) -> int:
    kunden = _lade_kunden(cfg)
    stunden_alle = ical_mod.lade_stunden(cfg.ical, year, month)
    auslagen_erg = expenses_mod.lade_auslagen(cfg.auslagen, year, month, set(kunden))
    # freigegeben = Meta!A1 == 'yes'. Im Dry-Run wird die ZUGFeRD-PDF trotzdem
    # erzeugt (zum Prüfen); Re-Nr/State/Versand steuert dry_run separat.
    freigegeben = auslagen_erg["freigegeben"]
    state = state_mod.lade_state(cfg.state_pfad)

    kunden_keys = sorted(set(stunden_alle) | set(auslagen_erg["auslagen"]))
    if nur_kunde:
        kunden_keys = [k for k in kunden_keys if k == nur_kunde]
    if not kunden_keys:
        print("Keine abrechenbaren Kunden für diesen Zeitraum (weder Stunden noch Auslagen).")
        return 0

    ergebnisse = []
    for k in kunden_keys:
        kunde = kunden.get(k)
        if kunde is None:
            print(f"Warnung: Kunde '{k}' nicht in kunden.ods – übersprungen.")
            continue
        stunden = stunden_alle.get(k, [])
        auslagen = auslagen_erg["auslagen"].get(k, [])
        if not stunden and not auslagen:
            continue
        erg = _erzeuge_rechnung(cfg, kunde, stunden, auslagen, month, year, state, dry_run, freigegeben)
        ergebnisse.append(erg)
        print(f"{'ENTWURF ' if erg['entwurf'] else 'FINALE  '}{k}: {erg['renr']}  Netto {euro(erg['summen']['netto'])}  Brutto {euro(erg['summen']['brutto'])}  → {erg['pdf']}")

    if dry_run:
        print("\nDRY RUN — keine E-Mails versendet, keine Re-Nr vergeben, kein State geschrieben.")
    return 0


def cmd_full(cfg: Config, month: int, year: int, *, dry_run: bool, nur_kunde: str | None) -> int:
    rc = cmd_invoice(cfg, month, year, dry_run=dry_run, nur_kunde=nur_kunde)
    if dry_run:
        return rc
    # Versand nur im echten Lauf
    kunden = _lade_kunden(cfg)
    state = state_mod.lade_state(cfg.state_pfad)
    out_verz = _ausgabe_verzeichnis(cfg, dry_run=False)
    # Neueste Buchungen dieses Laufs versenden: anhand Re-Nr in State
    for b in state.buchungen:
        if b.entwurf:
            continue
        if not b.leistungszeitraum == leistungszeitraum_iso(year, month):
            continue
        pdf_name = f"{b.kunde}_Rechnung_{b.renr}_ZUGFeRD.pdf"
        pdf_pfad = out_verz / pdf_name
        if not pdf_pfad.is_file():
            print(f"Warnung: finale PDF für Versand nicht gefunden: {pdf_pfad}")
            continue
        kunde = kunden.get(b.kunde)
        betreff = f"Rechnung {b.renr} – {kunde.name if kunde else b.kunde}"
        text = (
            f"Sehr geehrte Damen und Herren,\n\n"
            f"anbei erhalten Sie unsere Rechnung {b.renr} für den Leistungszeitraum "
            f"{b.leistungszeitraum} über einen Gesamtbetrag von {euro(Decimal(b.brutto))}.\n\n"
            f"Diese Rechnung enthält eine elektronische Rechnung gem. EN 16931 (ZUGFeRD).\n\n"
            f"Mit freundlichen Grüßen\n{cfg.leistender.name}\n"
        )
        try:
            smtp_mod.sende_rechnung(
                conf=cfg.smtp,
                empfaenger=None,
                pdf_pfad=pdf_pfad,
                betreff=betreff,
                text=text,
                dry_run=False,
            )
            print(f"Versendet: {b.renr} an {cfg.smtp.recipient_email}")
        except smtp_mod.SmtpError as exc:
            print(f"FEHLER beim Versand {b.renr}: {exc}")
    return rc


def cmd_journal(cfg: Config, year: int) -> int:
    state = state_mod.lade_state(cfg.state_pfad)
    zeilen = state_mod.journal_zeilen(state)
    if not zeilen:
        print(f"Keine Buchungen im State ({cfg.state_pfad}).")
        return 0
    if year:
        zeilen = [z for z in zeilen if z["datum"].startswith(str(year))]
    kopf = ["ReNr", "Datum", "Kunde", "LZ", "Netto", "USt", "Brutto", "Status"]
    tab = [[z["renr"], z["datum"], z["kunde"], z["leistungszeitraum"],
            z["netto"], z["ust"], z["brutto"], z["status"]] for z in zeilen]
    _drucke_tabelle(tab, kopf)
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="run-reu.py",
        description="REU – Rechnungserstellung (Stunden + Auslagen + ZUGFeRD).",
    )
    p.add_argument("user", help="Konfigurationsname, z.B. alice")
    p.add_argument("month", nargs="?", type=int, help="Monat (1-12)")
    p.add_argument("year", nargs="?", type=int, help="Jahr (z.B. 2025)")
    g = p.add_mutually_exclusive_group()
    g.add_argument("--hours", action="store_true", help="Nur Stunden aus iCal anzeigen")
    g.add_argument("--expenses", action="store_true", help="Nur Auslagen aus ODS prüfen")
    g.add_argument("--invoice", action="store_true", help="Rechnung(en) erzeugen (kein Versand)")
    g.add_argument("--full", action="store_true", help="Rechnung + ZUGFeRD + Versand (finale)")
    g.add_argument("--journal", action="store_true", help="Rechnungsübersicht aus State")
    p.add_argument("--dry-run", action="store_true",
                   help="PDF/XML erzeugen, aber kein Versand, keine Re-Nr, kein State")
    p.add_argument("--kunde", help="Nur diesen Kunden verarbeiten (Kürzel aus kunden.ods)")
    p.add_argument("--journal-year", type=int, dest="jyear",
                   help="Jahr-Filter für --journal")
    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    hat_action = any([args.hours, args.expenses, args.invoice, args.full, args.journal])
    if not hat_action:
        parser.print_help()
        return 0

    if args.journal:
        try:
            cfg = config_mod.lade_config(args.user)
        except config_mod.ConfigError as exc:
            print(f"Config-Fehler: {exc}")
            return 2
        return cmd_journal(cfg, args.jyear or 0)

    if not args.month or not args.year:
        print("Für --hours/--expenses/--invoice/--full sind Monat und Jahr Pflicht.")
        parser.print_help()
        return 2

    try:
        cfg = config_mod.lade_config(args.user)
    except config_mod.ConfigError as exc:
        print(f"Config-Fehler: {exc}")
        return 2

    try:
        if args.hours:
            return cmd_hours(cfg, args.month, args.year)
        if args.expenses:
            return cmd_expenses(cfg, args.month, args.year)
        if args.invoice:
            return cmd_invoice(cfg, args.month, args.year, dry_run=args.dry_run, nur_kunde=args.kunde)
        if args.full:
            return cmd_full(cfg, args.month, args.year, dry_run=args.dry_run, nur_kunde=args.kunde)
    except Exception as exc:  # noqa: BLE001
        print(f"FEHLER: {exc}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
