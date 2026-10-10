"""Stufe 3: Rechnungs-PDF mit reportlab erzeugen.

Pro Kunde und Zeitraum eine Rechnung mit:
  1. Sammelposition Stunden (Leistungszeitraum, Menge HUR, Stundensatz,
     Verweis auf Anlage Arbeitsstunden)
  2. Service-Positionen (je Service und Monat, Menge Stück)
  3. Sammelposition Auslagen (Verweis auf Anlage Auslagen)
  4. Summenblock (Netto, USt 19%, Brutto)
  5. Fuß mit USt-IdNr. (soweit vorhanden), IBAN/BIC, Zahlungsziel,
     Rechnungsnummer, Hinweis auf eingebettete E-Rechnung (nur bei Freigabe).

Außerdem werden zwei Anlage-Seiten beigelegt:
  - Anlage 'Arbeitsstunden' (Spalten: Datum, Startzeit, Beschreibung, Dauer)
  - Anlage 'Auslagen' (Spalten: Datum, Art, Bezeichnung, Beleg-Nr., Betrag)

Erste Seite im Briefkopf-Stil: Logo oben links, Absender + Rechnungsdaten
oben rechts auf Logo-Höhe, Empfänger-Adressfeld darunter.
"""
from __future__ import annotations

import datetime as _dt
from decimal import Decimal
from pathlib import Path
from typing import Any
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
)

from .config import Config, Leistender
from .kunden import Kunde
from .util import (
    euro,
    menge,
    Zeitraum,
    UST_PROZENT,
    UST_SATZ,
)


class InvoiceError(Exception):
    pass


def _styles() -> dict[str, ParagraphStyle]:
    styles = getSampleStyleSheet()
    return {
        "titel": ParagraphStyle("titel", parent=styles["Title"], fontSize=18, spaceAfter=4),
        "normal": ParagraphStyle("normal", parent=styles["Normal"], fontSize=9, leading=12),
        "klein": ParagraphStyle("klein", parent=styles["Normal"], fontSize=8, leading=10),
        "fuss": ParagraphStyle("fuss", parent=styles["Normal"], fontSize=8, leading=11),
    }


def _adresse_kunde(k: Kunde) -> str:
    zeilen = [k.name]
    if k.strasse:
        zeilen.append(k.strasse)
    zeilen.append(f"{k.plz} {k.ort}")
    zeilen.append(k.land)
    return "<br/>".join(zeilen)


def _adresse_leistender_zeilen(l: Leistender) -> list[str]:
    """Absenderadresse als Einzelzeilen (für den Briefkopf oben rechts)."""
    zeilen = [l.name]
    if l.strasse:
        zeilen.append(l.strasse)
    zeilen.append(f"{l.plz} {l.ort}")
    zeilen.append(l.land)
    return zeilen


def _positionen(
    kunde: Kunde,
    stunden: list[dict[str, Any]],
    auslagen: list[dict[str, Any]],
    services: list[dict[str, Any]],
    zeitraum: Zeitraum,
) -> tuple[list[list[Any]], dict[str, Decimal]]:
    """Baut die Positionstabelle und Summen.

    Services: Liste von {service, monat, jahr, menge, einzelpreis} –
    je Service und Monat mit Menge > 0 eine eigene Position.

    Returns (zeilen_fuer_tabelle, summen) mit summen:
      netto_stunden, netto_services, netto_auslagen, netto, ust, brutto,
      stunden_total
    """
    stunden_total = sum((e["dauer_h"] for e in stunden), Decimal("0"))
    netto_stunden = (stunden_total * kunde.stundensatz).quantize(Decimal("0.01"))
    netto_services = sum(
        (
            (s["menge"] * s["einzelpreis"]).quantize(Decimal("0.01"))
            for s in services
        ),
        Decimal("0"),
    )
    netto_auslagen = sum((a["betrag_netto"] for a in auslagen), Decimal("0"))
    netto = netto_stunden + netto_services + netto_auslagen
    ust = (netto * UST_SATZ).quantize(Decimal("0.01"))
    brutto = netto + ust

    pos_zeilen: list[list[Any]] = [
        ["Pos.", "Bezeichnung", "Menge", "Einzelpreis", "Netto"],
    ]
    pos_nr = 1
    if stunden_total > 0:
        pos_zeilen.append(
            [
                str(pos_nr),
                f"Beratungsleistung Leistungszeitraum {zeitraum.text} "
                f"(siehe Anlage Arbeitsstunden)",
                menge(stunden_total),
                euro(kunde.stundensatz),
                euro(netto_stunden),
            ]
        )
        pos_nr += 1
    for s in services:
        pos_zeilen.append(
            [
                str(pos_nr),
                f"{s['service']} – {s['monat']:02d}/{s['jahr']}",
                menge(s["menge"]),
                euro(s["einzelpreis"]),
                euro((s["menge"] * s["einzelpreis"]).quantize(Decimal("0.01"))),
            ]
        )
        pos_nr += 1
    if auslagen:
        # Sammelposition – Details stehen in der Anlage 'Auslagen'
        pos_zeilen.append(
            [
                str(pos_nr),
                "Auslagen (siehe Anlage Auslagen)",
                "1,00",
                euro(netto_auslagen),
                euro(netto_auslagen),
            ]
        )

    summen = {
        "stunden_total": stunden_total,
        "netto_stunden": netto_stunden,
        "netto_services": netto_services,
        "netto_auslagen": netto_auslagen,
        "netto": netto,
        "ust": ust,
        "brutto": brutto,
    }
    return pos_zeilen, summen


def _summenblock(summen: dict[str, Decimal]) -> Table:
    data = [
        ["Netto gesamt", euro(summen["netto"])],
        [f"USt {UST_PROZENT:.0f}%", euro(summen["ust"])],
        ["Brutto gesamt", euro(summen["brutto"])],
    ]
    t = Table(data, colWidths=[60 * mm, 30 * mm])
    t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("ALIGN", (1, 0), (1, -1), "RIGHT"),
                ("LINEABOVE", (0, 1), (-1, 1), 0.5, colors.grey),
                ("LINEABOVE", (0, 2), (-1, 2), 0.5, colors.grey),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("FONTNAME", (0, 2), (-1, 2), "Helvetica-Bold"),
            ]
        )
    )
    return t


def _stunden_anlage(stunden: list[dict[str, Any]], kunde: Kunde, zeitraum: Zeitraum) -> list:
    """Stundentabelle als Anlage-Seite."""
    styles = _styles()
    story: list = []
    story.append(Paragraph(f"Anlage: Arbeitsstunden {zeitraum.text}", styles["titel"]))
    story.append(Paragraph(f"Kunde: {kunde.name} ({kunde.kunde})", styles["normal"]))
    story.append(Spacer(1, 6 * mm))
    data = [["Datum", "Start", "Beschreibung", "Dauer (h)"]]
    total = Decimal("0")
    for e in stunden:
        data.append(
            [
                e["datum"].isoformat(),
                str(e.get("start", "")),
                e["beschreibung"],
                menge(e["dauer_h"]),
            ]
        )
        total += e["dauer_h"]
    data.append(["Summe", "", "", menge(total)])
    t = Table(data, colWidths=[20 * mm, 15 * mm, 113 * mm, 22 * mm])
    t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -2), 0.25, colors.grey),
                ("LINEABOVE", (0, -1), (-1, -1), 0.5, colors.black),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("ALIGN", (1, 1), (1, -1), "CENTER"),
                ("ALIGN", (3, 1), (3, -1), "RIGHT"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(t)
    return story


def _auslagen_anlage(auslagen: list[dict[str, Any]], kunde: Kunde, zeitraum: Zeitraum) -> list:
    """Auslagentabelle als Anlage-Seite."""
    styles = _styles()
    story: list = []
    story.append(Paragraph(f"Anlage: Auslagen {zeitraum.text}", styles["titel"]))
    story.append(Paragraph(f"Kunde: {kunde.name} ({kunde.kunde})", styles["normal"]))
    story.append(Spacer(1, 6 * mm))
    data = [["Datum", "Art", "Bezeichnung", "Beleg-Nr.", "Betrag"]]
    total = Decimal("0")
    for a in auslagen:
        data.append(
            [
                str(a["datum"]),
                a["art"],
                a["bezeichnung"],
                a["belegnr"],
                euro(a["betrag_netto"]),
            ]
        )
        total += a["betrag_netto"]
    data.append(["Summe", "", "", "", euro(total)])
    t = Table(data, colWidths=[24 * mm, 22 * mm, 85 * mm, 20 * mm, 24 * mm])
    t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -2), 0.25, colors.grey),
                ("LINEABOVE", (0, -1), (-1, -1), 0.5, colors.black),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("ALIGN", (4, 1), (4, -1), "RIGHT"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(t)
    return story


def erzeuge_rechnung_pdf(
    *,
    cfg: Config,
    kunde: Kunde,
    stunden: list[dict[str, Any]],
    auslagen: list[dict[str, Any]],
    services: list[dict[str, Any]],
    zeitraum: Zeitraum,
    renr: str,
    re_datum: _dt.date,
    faelligkeit: _dt.date,
    entwurf: bool,
    freigegeben: bool,
    ausgabe_pfad: Path,
) -> dict[str, Decimal]:
    """Erzeugt die Rechnungs-PDF und liefert die Summen zurück.

    entwurf: True → Wasserzeichen ENTWURF, Hinweis 'keine E-Rechnung eingebettet'.
    freigegeben: True → Hinweis auf eingebettete ZUGFeRD (nur bei finaler Rechnung).
    """
    styles = _styles()
    pos_zeilen, summen = _positionen(kunde, stunden, auslagen, services, zeitraum)

    story: list = []
    # Briefkopf: Logo (oben links) und der rechte Block mit Absender +
    # Rechnungsdaten werden auf der ersten Seite direkt auf dem Canvas
    # gezeichnet (onFirstPage-Callback) – beide mit Oberkante an der
    # topMargin-Kante, also auf gleicher Höhe, unabhängig von
    # Flowable-Alignment-Eigenheiten der reportlab-Version.
    # Im Story bleibt nur der Platzhalter-Abstand.
    logo_w, logo_h = 31.5 * mm, 14 * mm
    hat_logo = cfg.logo_path.is_file()

    def _briefkopf(canvas, doc):  # noqa: ANN001 – reportlab-Signatur
        # Logo oben links am Satzspiegelrand
        if hat_logo:
            try:
                canvas.drawImage(
                    str(cfg.logo_path),
                    doc.leftMargin,
                    doc.pagesize[1] - doc.topMargin - logo_h,
                    width=logo_w,
                    height=logo_h,
                    mask="auto",
                )
            except Exception:  # noqa: BLE001 – Logo darf den Druck nie blockieren
                pass
        # Rechter Block: Absenderadresse + Rechnungsdaten, oben rechts,
        # Oberkante auf gleicher Höhe wie das Logo (topMargin-Kante).
        canvas.saveState()
        canvas.setFont("Helvetica", 9)
        x_rechts = doc.pagesize[0] - doc.rightMargin
        y = doc.pagesize[1] - doc.topMargin - 6.5
        for zeile in _adresse_leistender_zeilen(cfg.leistender):
            canvas.drawRightString(x_rechts, y, zeile)
            y -= 12
        y -= 4  # kleine Lücke zwischen Absender und Rechnungsdaten
        for zeile in (
            f"Rechnungsdatum: {re_datum.isoformat()}",
            f"Rechnungsnummer: {renr}",
            f"Leistungszeitraum: {zeitraum.text}",
        ):
            canvas.drawRightString(x_rechts, y, zeile)
            y -= 12
        canvas.restoreState()

    story.append(Spacer(1, (logo_h + 12 * mm) if hat_logo else 16 * mm))

    # Empfängeradresse links (Absender + Re-Daten stehen oben rechts,
    # per Canvas auf Logo-Höhe – siehe _briefkopf)
    story.append(Spacer(1, 4 * mm))
    story.append(Paragraph(_adresse_kunde(kunde), styles["normal"]))
    # Ausgleich: Der Re-Daten-Block (3 Zeilen à 12 pt) steht jetzt oben
    # rechts im Canvas statt unter der Absenderadresse – dieser Abstand
    # hält Titel und Positionstabelle auf der gewohnten Höhe.
    story.append(Spacer(1, 6 * mm + 36))

    story.append(Paragraph("Rechnung", styles["titel"]))
    if entwurf:
        story.append(
            Paragraph(
                '<font color="red"><b>ENTWURF – nicht zur Zahlung freigegeben</b></font>',
                styles["normal"],
            )
        )
    story.append(Spacer(1, 4 * mm))

    # Positionstabelle – Bezeichnung als Paragraph, damit lange Texte
    # automatisch in der Spalte umbrechen (mehrzeilig).
    tab_daten = [pos_zeilen[0]] + [
        [z[0]]
        + [Paragraph(escape(str(z[1])).replace("\n", "<br/>"), styles["normal"])]
        + list(z[2:])
        for z in pos_zeilen[1:]
    ]
    t = Table(tab_daten, colWidths=[12 * mm, 92 * mm, 18 * mm, 24 * mm, 26 * mm])
    t.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
                ("ALIGN", (2, 0), (-1, -1), "RIGHT"),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(t)
    story.append(Spacer(1, 4 * mm))
    story.append(_summenblock(summen))
    story.append(Spacer(1, 8 * mm))

    # Fuß
    fuss_text = (
        f"Zahlungsziel: bis {faelligkeit.isoformat()}.<br/>"
        f"IBAN: {cfg.erechnung.iban}<br/>"
        f"BIC: {cfg.erechnung.bic}<br/>"
        f"Bank: {cfg.erechnung.bank}<br/>"
        f"USt-IdNr. (Leistender): {cfg.leistender.ust_id}<br/>"
    )
    if kunde.ust_id:
        fuss_text += f"USt-IdNr. (Kunde): {kunde.ust_id}<br/>"
    if freigegeben and not entwurf:
        fuss_text += (
            "<br/>Diese Rechnung enthält eine elektronische Rechnung gem. "
            "EN 16931 (ZUGFeRD / Faktur-X) als eingebettete Datei."
        )
    else:
        fuss_text += (
            "<br/>Dies ist ein Rechnungsentwurf. Es ist keine elektronische "
            "Rechnung gem. EN 16931 eingebettet."
        )
    story.append(Paragraph(fuss_text, styles["fuss"]))

    # Anlage: Stunden
    if stunden:
        story.append(PageBreak())
        story.extend(_stunden_anlage(stunden, kunde, zeitraum))

    # Anlage: Auslagen
    if auslagen:
        story.append(PageBreak())
        story.extend(_auslagen_anlage(auslagen, kunde, zeitraum))

    ausgabe_pfad.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(ausgabe_pfad),
        pagesize=A4,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title=f"Rechnung {renr}",
        author=cfg.leistender.name,
    )
    doc.build(story, onFirstPage=_briefkopf)
    return summen
