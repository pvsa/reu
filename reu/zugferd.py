"""Stufe 4: ZUGFeRD – Faktur-X XML (EN 16931) + Einbettung in PDF/A-3.

Verwendet drafthorse für das XML (CrossIndustryInvoice, Profil EN 16931)
und drafthorse.pdf.attach_xml für das PDF/A-3 + AF-Embedding.

Mapping (B2B, immer 19% USt, ein Stunden-Sammelposten + Serviceposten
je Service und Monat + Auslagenposten).
TaxRegistration (USt-IdNr.) wird nur gesetzt, wenn eine USt-IdNr. vorliegt
(Verkäufer: Pflicht in der Config; Käufer: optional).
"""
from __future__ import annotations

import datetime as _dt
from decimal import Decimal
from pathlib import Path
from typing import Any

import drafthorse
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


# ------------------------------------------------- drafthorse-Kompatibilitaet
# Die Modell-API von drafthorse hat sich ueber die Versionen geaendert.
# Diese Helfer halten erzeuge_xml kompatibel mit drafthorse 2.2 bis 2025.x:
#
# 1. Betraege: CurrencyField (z.B. MonetarySummation.tax_basis_total/
#    grand_total bis drafthorse 2025.1) akzeptiert den Betrag direkt und
#    setzt intern die Waehrung auf None – kein currencyID-Attribut.
#    Das 2-Tupel (wert, waehrung) schreibt currencyID: zulässig und von
#    BR-CO-15 GEFORDERT, aber nur auf ram:TaxTotalAmount (BT-110) – auf
#    allen anderen Amount-Elementen ist es CII-DT-031 (fatal) verboten.
#    Daher: _betrag (ohne Attribut) und _betrag_waehrung (nur tax_total).
# 2. IncludedNote.content: erst ab 2025.1 ein StringField (vorher
#    MultiStringField -> Wert per .add() statt direkter Zuweisung; die
#    wuerde AttributeError 'Read-only!' aus Field.__set__ ausloesen).
# 3. TradeSettlement.payment_means: erst ab 2025.1 ein MultiField (vorher
#    einzelnes Field -> Werte in-place am Element setzen statt .add()).
#
# drafthorse 1.x nutzt die inkompatible ZUGFeRD-1.0-API (serialize() ohne
# schema-Parameter) und wird mit einer klaren Meldung abgewiesen.


def _pruefe_drafthorse_version() -> None:
    """Wirft ZugferdError bei drafthorse 1.x (ZUGFeRD-1.0-API, kein EN 16931)."""
    version = str(getattr(drafthorse, "version", ""))
    try:
        hauptversion = int(version.split(".")[0])
    except ValueError:
        return  # unbekanntes Versionsformat: nicht blocken
    if hauptversion < 2:
        raise ZugferdError(
            f"installierte drafthorse-Version {version!r} ist zu alt "
            "(ZUGFeRD-1.0-API, kein EN 16931 / FACTUR-X). "
            "Bitte aktualisieren: pip install -U 'drafthorse>=2.2'"
        )


def _betrag(objekt: Any, name: str, wert: Decimal) -> None:
    """Setzt ein Betragsfeld versionstolerant – OHNE currencyID-Attribut.

    CurrencyField akzeptiert den Betrag direkt (drafthorse >= 2.2 setzt
    intern die Waehrung auf None). Nur falls eine sehr alte drafthorse-
    Version CurrencyField noch ein 2-Tupel verlangt, wird (wert, None)
    gesetzt – ebenfalls ohne currencyID: Auf allen Amount-Elementen
    AUSSER TaxTotalAmount verbietet EN 16931 (CII-DT-031, fatal) das
    Attribut; die Waehrung steht allein im InvoiceCurrencyCode.
    """
    try:
        setattr(objekt, name, wert)
    except TypeError:
        setattr(objekt, name, (wert, None))


def _betrag_waehrung(objekt: Any, name: str, wert: Decimal, waehrung: str) -> None:
    """Setzt ein Betragsfeld MIT currencyID-Attribut – nur TaxTotalAmount!

    BR-CO-15 (EN 16931) erwartet ram:TaxTotalAmount mit currencyID
    passend zum InvoiceCurrencyCode (ohne das Attribut greift die Regel
    auf den Ausnahmefall 'keine USt' zurueck und schlaegt fehl).
    CurrencyField schreibt das Attribut beim 2-Tupel (wert, waehrung).
    """
    try:
        setattr(objekt, name, (wert, waehrung))
    except TypeError:
        setattr(objekt, name, wert)


def _notizinhalt(note: IncludedNote, text: str) -> None:
    """IncludedNote.content versionstolerant setzen."""
    inhalt = note.content
    if hasattr(inhalt, "add"):  # MultiStringField (drafthorse <= 2.4)
        inhalt.add(text)
    else:  # StringField (drafthorse 2025.x)
        note.content = text


def _konfiguriere_payment_means(pm: PaymentMeans, cfg: Config) -> None:
    pm.type_code = "58"
    pm.payee_account.iban = (cfg.erechnung.iban or "").replace(" ", "")
    if cfg.erechnung.bic:
        pm.payee_institution.bic = cfg.erechnung.bic


def _payment_means_hinzufuegen(settlement: Any, cfg: Config) -> None:
    """PaymentMeans versionstolerant am TradeSettlement setzen."""
    ziel = settlement.payment_means
    if hasattr(ziel, "add"):  # MultiField (drafthorse 2025.x)
        pm = PaymentMeans()
        _konfiguriere_payment_means(pm, cfg)
        ziel.add(pm)
    else:  # einzelnes Field (drafthorse <= 2.4): in-place konfigurieren
        _konfiguriere_payment_means(ziel, cfg)


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
    _betrag(item.agreement.gross, "amount", einzelpreis)
    _betrag(item.agreement.net, "amount", einzelpreis)
    item.delivery.billed_quantity = (menge, unit)
    tax = item.settlement.trade_tax
    tax.type_code = "VAT"
    tax.category_code = "S"
    tax.rate_applicable_percent = UST_PROZENT
    _betrag(
        item.settlement.monetary_summation,
        "total_amount",
        (menge * einzelpreis).quantize(Decimal("0.01")),
    )
    document.trade.items.add(item)


def erzeuge_xml(
    *,
    cfg: Config,
    kunde: Kunde,
    stunden: list[dict[str, Any]],
    auslagen: list[dict[str, Any]],
    services: list[dict[str, Any]],
    summen: dict[str, Decimal],
    renr: str,
    re_datum: _dt.date,
    faelligkeit: _dt.date,
    zeitraum: Zeitraum,
) -> bytes:
    """Erzeugt validiertes Faktur-X XML (EN 16931)."""
    _pruefe_drafthorse_version()
    document = Document()
    document.context.guideline_parameter.id = GUIDELINE_EN16931
    document.header.id = renr
    document.header.type_code = "380"
    document.header.issue_date_time = re_datum

    note = IncludedNote()
    _notizinhalt(note, "E-Rechnung gem. EN 16931 (ZUGFeRD / Faktur-X)")
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
    _betrag(tax, "calculated_amount", summen["ust"])
    tax.type_code = "VAT"
    _betrag(tax, "basis_amount", summen["netto"])
    tax.category_code = "S"
    tax.rate_applicable_percent = UST_PROZENT
    settlement.trade_tax.add(tax)

    terms = PaymentTerms()
    terms.description = f"Zahlbar bis {faelligkeit.isoformat()}"
    terms.due = faelligkeit
    settlement.terms.add(terms)

    _payment_means_hinzufuegen(settlement, cfg)

    # Positionen: Stunden (falls > 0), dann Services (je Service+Monat),
    # dann EINE Sammelposition Auslagen – konsistent zur PDF (Details in Anlage).
    pos_nr = 1
    stunden_total = summen["stunden_total"]
    if stunden_total > 0:
        _add_lineitem(
            document,
            line_id=str(pos_nr),
            name=(
                f"Beratungsleistung Leistungszeitraum {zeitraum.text} "
                f"(siehe Anlage Arbeitsstunden)"
            ),
            description=f"{stunden_total} Stunden zu je {kunde.stundensatz} EUR",
            menge=stunden_total,
            unit=UNIT_STUNDE,
            einzelpreis=kunde.stundensatz,
        )
        pos_nr += 1
    for s in services:
        _add_lineitem(
            document,
            line_id=str(pos_nr),
            name=f"{s['service']} – {s['monat']:02d}/{s['jahr']}",
            description=f"Serviceleistung {s['monat']:02d}/{s['jahr']}",
            menge=s["menge"],
            unit=UNIT_STUECK,
            einzelpreis=s["einzelpreis"],
        )
        pos_nr += 1
    if auslagen:
        _add_lineitem(
            document,
            line_id=str(pos_nr),
            name="Auslagen (siehe Anlage Auslagen)",
            description="Auslagen gemäß Anlage Auslagen",
            menge=Decimal("1"),
            unit=UNIT_STUECK,
            einzelpreis=summen["netto_auslagen"],
        )

    mon = settlement.monetary_summation
    _betrag(mon, "line_total", summen["netto"])
    _betrag(mon, "charge_total", Decimal("0"))
    _betrag(mon, "allowance_total", Decimal("0"))
    _betrag(mon, "tax_basis_total", summen["netto"])
    # TaxTotalAmount (BT-110) als EINZIGES Amount-Element MIT currencyID
    # (BR-CO-15); alle anderen Amounts bleiben attributlos (CII-DT-031).
    _betrag_waehrung(mon, "tax_total", summen["ust"], WAehrUNG)
    _betrag(mon, "grand_total", summen["brutto"])
    _betrag(mon, "due_amount", summen["brutto"])

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
    services: list[dict[str, Any]],
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
        services=services,
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
