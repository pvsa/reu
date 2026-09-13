"""E-Mail-Versand der Rechnungs-PDF per SMTP.

Bei dry_run: no-op (es wird nichts gesendet, nichts geladen).
"""
from __future__ import annotations

import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path

from .config import SmtpConf


class SmtpError(Exception):
    pass


def sende_rechnung(
    *,
    conf: SmtpConf,
    empfaenger: str | None,
    pdf_pfad: Path,
    betreff: str,
    text: str,
    dry_run: bool,
) -> None:
    """Sendet die PDF als Anhang. Bei dry_run: nur Konsole-Meldung.

    empfaenger: überschreibt recipient_email aus conf, falls gesetzt.
    """
    if dry_run:
        print(f"DRY RUN: E-Mail-Versand übersprungen (wäre an {empfaenger or conf.recipient_email}).")
        return
    if not conf.server:
        raise SmtpError("[smtp] server ist nicht konfiguriert – Versand nicht möglich.")
    an = empfaenger or conf.recipient_email
    if not an:
        raise SmtpError("Kein Empfänger angegeben und [smtp] recipient_email leer.")
    if not pdf_pfad.is_file():
        raise SmtpError(f"PDF nicht gefunden: {pdf_pfad}")

    msg = EmailMessage()
    msg["From"] = conf.sender_email or conf.smtp_username
    msg["To"] = an
    msg["Subject"] = betreff
    msg.set_content(text)
    data = pdf_pfad.read_bytes()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="pdf",
        filename=pdf_pfad.name,
    )

    try:
        context = ssl.create_default_context()
        if conf.use_starttls:
            with smtplib.SMTP(conf.server, conf.port, timeout=60) as server:
                server.starttls(context=context)
                if conf.smtp_username:
                    server.login(conf.smtp_username, conf.smtp_password or "")
                server.send_message(msg)
        else:
            with smtplib.SMTP_SSL(conf.server, conf.port, context=context, timeout=60) as server:
                if conf.smtp_username:
                    server.login(conf.smtp_username, conf.smtp_password or "")
                server.send_message(msg)
    except smtplib.SMTPException as exc:  # noqa: BLE001
        raise SmtpError(f"SMTP-Versand fehlgeschlagen: {exc}") from exc
