"""Config-Loader: lädt conf/<user>.conf (INI) und legt Pfade fest.

Die Kundendaten werden NICHT hier geladen – das macht kunden.py
aus der pro-User-ODS-Datei. Hier nur die Konfigurationsdatei (INI)
mit [leistender], [kunden], [auslagen], [ical], [smtp], [pdf], [erechnung].

Hinweise:
- smtp_username/smtp_password sind OPTIONAL – nur setzen, wenn der
  SMTP-Server eine Authentifizierung verlangt (z.B. offener Relay ohne).
- Eine Steuernummer des Leistenden wird nicht mehr geführt (nur ust_id).
"""
from __future__ import annotations

import configparser
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


class ConfigError(Exception):
    pass


def _basis_verzeichnis() -> Path:
    return Path(__file__).resolve().parent.parent


@dataclass
class Leistender:
    name: str
    strasse: str
    plz: str
    ort: str
    land: str
    ust_id: str
    leitweg_id: str = ""


@dataclass
class ERechnung:
    profil: str = "ZUGFeRD"
    format: str = "pdfa3"
    zahlungsziel_tage: int = 14
    iban: str = ""
    bic: str = ""
    bank: str = ""


@dataclass
class AuslagenConf:
    quelle: str = "ods"
    datei: str = ""
    blatt: str = "Auslagen"
    freigabe_spalte: str = "freigabe"
    pflichtspalten: list[str] = field(
        default_factory=lambda: ["datum", "kunde", "art", "bezeichnung", "betrag_netto", "belegnr"]
    )


@dataclass
class ICalConf:
    url: str = ""
    username: str = ""
    password: str = ""
    auth_method: str = "basic"
    timeout: int = 60
    max_retries: int = 3
    retry_delay: int = 5


@dataclass
class SmtpConf:
    server: str = ""
    port: int = 587
    sender_email: str = ""
    recipient_email: str = ""
    use_starttls: bool = True
    smtp_username: str = ""
    smtp_password: str = ""


@dataclass
class KundenConf:
    datei: str
    blatt: str = "Kunden"


@dataclass
class Config:
    user: str
    conf_pfad: Path
    basis: Path
    leistender: Leistender
    erechnung: ERechnung
    auslagen: AuslagenConf
    ical: ICalConf
    smtp: SmtpConf
    kunden: KundenConf
    logo_path: Path
    ausgabe_verz: Path
    state_pfad: Path


def _get(c: configparser.ConfigParser, section: str, key: str, *, fallback: Any = None, typ: type = str) -> Any:
    if not c.has_option(section, key):
        if fallback is not None:
            return fallback
        raise ConfigError(f"Config-Eintrag fehlt: [{section}] {key}")
    roh = c.get(section, key)
    if typ is bool:
        return roh.strip().lower() in {"true", "yes", "1", "ja"}
    if typ is int:
        return int(roh.strip())
    return roh.strip()


def lade_config(user: str, *, basis: Path | None = None) -> Config:
    basis = basis or _basis_verzeichnis()
    conf_pfad = basis / "conf" / f"{user}.conf"
    if not conf_pfad.is_file():
        raise ConfigError(f"Config-Datei nicht gefunden: {conf_pfad}")
    parser = configparser.ConfigParser(interpolation=None, strict=False)
    parser.optionxform = str
    parser.read(conf_pfad, encoding="utf-8")

    for sektion in ("leistender", "kunden", "auslagen", "ical", "smtp", "pdf", "erechnung"):
        if not parser.has_section(sektion):
            raise ConfigError(f"Config-Sektion fehlt: [{sektion}] in {conf_pfad}")

    leistender = Leistender(
        name=_get(parser, "leistender", "name"),
        strasse=_get(parser, "leistender", "strasse", fallback=""),
        plz=_get(parser, "leistender", "plz"),
        ort=_get(parser, "leistender", "ort"),
        land=_get(parser, "leistender", "land", fallback="DE"),
        ust_id=_get(parser, "leistender", "ust_id"),
        leitweg_id=_get(parser, "leistender", "leitweg_id", fallback=""),
    )

    erechnung = ERechnung(
        profil=_get(parser, "erechnung", "profil", fallback="ZUGFeRD"),
        format=_get(parser, "erechnung", "format", fallback="pdfa3"),
        zahlungsziel_tage=int(_get(parser, "erechnung", "zahlungsziel_tage", fallback="14")),
        iban=_get(parser, "erechnung", "iban", fallback=""),
        bic=_get(parser, "erechnung", "bic", fallback=""),
        bank=_get(parser, "erechnung", "bank", fallback=""),
    )

    pflichtspalten_roh = _get(parser, "auslagen", "pflichtspalten", fallback="")
    pflichtspalten = [s.strip() for s in pflichtspalten_roh.split(",") if s.strip()] or [
        "datum", "kunde", "art", "bezeichnung", "betrag_netto", "belegnr"
    ]
    auslagen_datei_roh = _get(parser, "auslagen", "datei")
    auslagen = AuslagenConf(
        quelle=_get(parser, "auslagen", "quelle", fallback="ods"),
        datei=str(
            (basis / auslagen_datei_roh)
            if not os.path.isabs(auslagen_datei_roh)
            else auslagen_datei_roh
        ),
        blatt=_get(parser, "auslagen", "blatt", fallback="Auslagen"),
        freigabe_spalte=_get(parser, "auslagen", "freigabe_spalte", fallback="freigabe"),
        pflichtspalten=pflichtspalten,
    )

    ical = ICalConf(
        url=_get(parser, "ical", "url"),
        username=_get(parser, "ical", "username", fallback=""),
        password=_get(parser, "ical", "password", fallback=""),
        auth_method=_get(parser, "ical", "auth_method", fallback="basic"),
        timeout=int(_get(parser, "ical", "timeout", fallback="60")),
        max_retries=int(_get(parser, "ical", "max_retries", fallback="3")),
        retry_delay=int(_get(parser, "ical", "retry_delay", fallback="5")),
    )

    smtp = SmtpConf(
        server=_get(parser, "smtp", "server", fallback=""),
        port=int(_get(parser, "smtp", "port", fallback="587")),
        sender_email=_get(parser, "smtp", "sender_email", fallback=""),
        recipient_email=_get(parser, "smtp", "recipient_email", fallback=""),
        use_starttls=_get(parser, "smtp", "use_starttls", fallback=True, typ=bool),
        smtp_username=_get(parser, "smtp", "smtp_username", fallback=""),
        smtp_password=_get(parser, "smtp", "smtp_password", fallback=""),
    )

    kunden_datei = _get(parser, "kunden", "datei")
    kunden = KundenConf(
        datei=str((basis / kunden_datei) if not os.path.isabs(kunden_datei) else kunden_datei),
        blatt=_get(parser, "kunden", "blatt", fallback="Kunden"),
    )

    logo_roh = _get(parser, "pdf", "logo_path", fallback="logo.jpg")
    logo_path = (basis / logo_roh) if not os.path.isabs(logo_roh) else Path(logo_roh)

    ausgabe_verz = basis / "out"
    state_pfad = basis / "conf" / f"{user}.state.json"

    return Config(
        user=user,
        conf_pfad=conf_pfad,
        basis=basis,
        leistender=leistender,
        erechnung=erechnung,
        auslagen=auslagen,
        ical=ical,
        smtp=smtp,
        kunden=kunden,
        logo_path=logo_path,
        ausgabe_verz=ausgabe_verz,
        state_pfad=state_pfad,
    )
