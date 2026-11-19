# REU – Offene Punkte

## Später angehen

- [ ] **E-Mail-Versand spezifizieren (wie und an wen)**:
  Derzeit sendet REU NUR mit `--send` (Standard: kein Versand, nur Hinweis).
  Empfänger ist [smtp] recipient_email – eine Adresse für ALLE Rechnungen,
  Text hardcoded in `cmd_full`. Später klären:
  Empfänger je Kunde (eigene Spalte in der kunden.ods?),
  Betreff-/Text-Vorlage, Anrede, CC/Archiv-Adresse.

- [ ] **GitHub-Anbindung**: Repo anlegen, Commit/Push der aktuellen Sandbox-Version.
  Vorbedingung: GitHub-Connector im Vibe-Session aktivieren (`list_unauthenticated_connectors`
  → `ask_enable_connector`), danach `gh`/`git` authentifiziert nutzen.
- [ ] **Serientermine (RRULE)** im iCal-Parser (`reu/ical.py`):
  Derzeit werden nur Einzeltermine (`VEVENT` mit DTSTART/DTEND) gelesen.
  Wiederkehrende Termine (`RRULE`/`RDATE`) werden noch nicht expandiert.
  Lösung: `dateutil.rrule` zur Expansion im angefragten Monat verwenden
  (Dependency dann aufnehmen: `python-dateutil`).

## Mittelbar / Verbesserungen

- [x] **requirements.txt** anlegen (`drafthorse>=2.2` – Versionen 2.2–2025.x
  werden von `reu/zugferd.py` versionstolerant unterstützt, 1.x wird mit
  klarer Meldung abgewiesen; `reportlab`, `requests`, `icalendar`, `pytz`;
  `lxml`/`pypdf` transitiv via drafthorse).
- [x] **Volle PDF/A-3-Konformität**: Quell-PDF wird vor `attach_xml` per
  Ghostscript nach PDF/A-3b konvertiert (`reu/pdfa.py`, `-dPDFA=3
  -dPDFACompatibilityPolicy=1`, gepatchte `PDFA_def.ps` mit sRGB-ICC);
  schaltbar via `[erechnung] format`, gs-Binärdatei/ICC/def per Konfig
  übersteuerbar. Rest-Validierung beim ersten echten Versand: veraPDF
  (PDF/A) + Mustang (Faktur-X), siehe README.
- [ ] **Mustang-Validierung in CI** integrieren (Docker) laut README.
- [ ] **drafthorse-Lizenzhinweis** in README ergänzen (Apache-2.0 Code,
  proprietäre FeRD-Lizenz für mitgelieferte .xsd-Schemata).
