"""Stufe 4: ZUGFeRD – Faktur-X XML (EN 16931) + Einbettung in PDF/A-3.

Verwendet drafthorse für das XML (CrossIndustryInvoice, Profil EN 16931)
und drafthorse.pdf.attach_xml für das PDF/A-3 + AF-Embedding.

Mapping (B2B, immer 19% USt, ein Stunden-Sammelposten + Auslagenposten).
TaxRegistration (USt-IdNr.) wird nur gesetzt, wenn eine USt-IdNr. vorliegt
(Verkäufer: Pflicht in der Config; Käufer: optional).
"""
from __future__ import annotations

import datetime as _dt
from decimal import Decimal
from pathlib import Path
from typing import Any

from drafthorse.models.accounting import ApplicableTradeTax
from drafthorse.models.document import Document
from drafthorse.models.note import IncludedNote
from drafthorse.models.party import TaxRegistration
from drafthorse.models.payment import PaymentMeans, PaymentTerms
from drafthorse.models.tradelines import LineItem
from drafthorse.pdf import attach_xml

from .config import Config
from .kunden import Kunde
from .util import (
    UST_PROZENT,
    UNIT_STUNDE,
    UNIT_STUECK,
    Zeitraum,
)

GUIDELINE_EN16931 = "urn:cen.eu:en16931:2017"
WAehrUNG = "EUR"


class ZugferdError(Exception):
    pass


def _setz_party(party: Any, name: str, strasse: str, plz: str, ort: str, land: str,
                 ust_id: str, leitweg_id: str = "") -> None:
    party.name = name
    if leitweg_id:
        party.id = leitweg_id
    addr = party.address
    if strasse:
        addr.line_one = strasse
    addr.postcode = plz
    addr.city_name = ort
    addr.country_id = land
    if ust_id:
        party.tax_registrations.add(TaxRegistration(id=("VA", ust_id)))


def _add_lineitem(
    document: Document,
    *,
    line_id: str,
    name: str,
    description: str,
    menge: Decimal,
    unit: str,
    einzelpreis: Decimal,
) -> None:
    item = LineItem()
    item.document.line_id = line_id
    product = item.product
    product.name = name
    if description:
        product.description = description
    item.agreement.gross.amount = einzelpreis
    item.agreement.net.amount = einzelpreis
    item.delivery.billed_quantity = (menge, unit)
    tax = item.settlement.trade_tax
    tax.type_code = "VAT"
    tax.category_code = "S"
    tax.rate_applicable_percent = UST_PROZENT
    item.settlement.monetary_summation.total_amount = (menge * einzelpreis).quantize(Decimal("0.01"))
    document.trade.items.add(item)


def erzeuge_xml(
    *,
    cfg: Config,
    kunde: Kunde,
    stunden: list[dict[str, Any]],
    auslagen: list[dict[str, Any]],
    summen: dict[str, Decimal],
    renr: str,
    re_datum: _dt.date,
    faelligkeit: _dt.date,
    zeitraum: Zeitraum,
) -> bytes:
    """Erzeugt validiertes Faktur-X XML (EN 16931)."""
    document = Document()
    document.context.guideline_parameter.id = GUIDELINE_EN16931
    document.header.id = renr
    document.header.type_code = "380"
    document.header.issue_date_time = re_datum

    note = IncludedNote()
    note.content = "E-Rechnung gem. EN 16931 (ZUGFeRD / Faktur-X)"
    note.subject_code = "REG"
    document.header.notes.add(note)

    _setz_party(
        document.trade.agreement.seller,
        name=cfg.leistender.name,
        strasse=cfg.leistender.strasse,
        plz=cfg.leistender.plz,
        ort=cfg.leistender.ort,
        land=cfg.leistender.land,
        ust_id=cfg.leistender.ust_id,
        leitweg_id=cfg.leistender.leitweg_id,
    )
    _setz_party(
        document.trade.agreement.buyer,
        name=kunde.name,
        strasse=kunde.strasse,
        plz=kunde.plz,
        ort=kunde.ort,
        land=kunde.land,
        ust_id=kunde.ust_id,
        leitweg_id=kunde.leitweg_id,
    )

    document.trade.delivery.event.occurrence = zeitraum.ende_datum

    settlement = document.trade.settlement
    settlement.currency_code = WAehrUNG

    tax = ApplicableTradeTax()
    tax.calculated_amount = summen["ust"]
    tax.type_code = "VAT"
    tax.basis_amount = summen["netto"]
    tax.category_code = "S"
    tax.rate_applicable_percent = UST_PROZENT
    settlement.trade_tax.add(tax)

    terms = PaymentTerms()
    terms.description = f"Zahlbar bis {faelligkeit.isoformat()}"
    terms.due = faelligkeit
    settlement.terms.add(terms)

    pm = PaymentMeans()
    pm.type_code = "58"
    pm.payee_account.iban = (cfg.erechnung.iban or "").replace(" ", "")
    if cfg.erechnung.bic:
        pm.payee_institution.bic = cfg.erechnung.bic
    settlement.payment_means.add(pm)

    # Position 1: Stunden (Sammelposition)
    stunden_total = summen["stunden_total"]
    _add_lineitem(
        document,
        line_id="1",
        name=f"Beratungsleistung Leistungszeitraum {zeitraum.text}",
        description=f"{stunden_total} Stunden zu je {kunde.stundensatz} EUR",
        menge=stunden_total,
        unit=UNIT_STUNDE,
        einzelpreis=kunde.stundensatz,
    )
    for i, a in enumerate(auslagen, start=2):
        bez = f"{a['art']} {a['bezeichnung']}".strip()
        if a["belegnr"]:
            bez += f" (Beleg {a['belegnr']})"
        _add_lineitem(
            document,
            line_id=str(i),
            name=bez,
            description=f"Auslage, Beleg {a['belegnr']}" if a["belegnr"] else "Auslage",
            menge=Decimal("1"),
            unit=UNIT_STUECK,
            einzelpreis=a["betrag_netto"],
        )

    mon = settlement.monetary_summation
    mon.line_total = summen["netto"]
    mon.charge_total = Decimal("0")
    mon.allowance_total = Decimal("0")
    mon.tax_basis_total = summen["netto"]
    mon.tax_total = summen["ust"]
    mon.grand_total = summen["brutto"]
    mon.due_amount = summen["brutto"]

    try:
        xml = document.serialize(schema="FACTUR-X_EN16931")
    except Exception as exc:  # noqa: BLE001
        raise ZugferdError(f"XML-Erzeugung/-Validierung fehlgeschlagen: {exc}") from exc
    return xml


def bette_xml_in_pdf(pdf_bytes: bytes, xml_bytes: bytes) -> bytes:
    """Erzeugt ZUGFeRD-PDF (PDF/A-3 + AF-EmbeddedFile)."""
    try:
        return attach_xml(pdf_bytes, xml_bytes, level="EN 16931")
    except Exception as exc:  # noqa: BLE001
        raise ZugferdError(f"PDF/A-3-Einbettung fehlgeschlagen: {exc}") from exc


def erzeuge_zugferd_pdf(
    *,
    pdf_pfad: Path,
    ausgabe_pfad: Path,
    cfg: Config,
    kunde: Kunde,
    stunden: list[dict[str, Any]],
    auslagen: list[dict[str, Any]],
    summen: dict[str, Decimal],
    renr: str,
    re_datum: _dt.date,
    faelligkeit: _dt.date,
    zeitraum: Zeitraum,
) -> Path:
    """Erzeugt XML + PDF/A-3 und schreibt die finale ZUGFeRD-PDF."""
    xml = erzeuge_xml(
        cfg=cfg,
        kunde=kunde,
        stunden=stunden,
        auslagen=auslagen,
        summen=summen,
        renr=renr,
        re_datum=re_datum,
        faelligkeit=faelligkeit,
        zeitraum=zeitraum,
    )
    pdf_bytes = pdf_pfad.read_bytes()
    zugferd_pdf = bette_xml_in_pdf(pdf_bytes, xml)
    ausgabe_pfad.parent.mkdir(parents=True, exist_ok=True)
    ausgabe_pfad.write_bytes(zugferd_pdf)
    return ausgabe_pfad
