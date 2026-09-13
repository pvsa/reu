# REU – Offene Punkte

## Später angehen

- [ ] **GitHub-Anbindung**: Repo anlegen, Commit/Push der aktuellen Sandbox-Version.
  Vorbedingung: GitHub-Connector im Vibe-Session aktivieren (`list_unauthenticated_connectors`
  → `ask_enable_connector`), danach `gh`/`git` authentifiziert nutzen.
- [ ] **Serientermine (RRULE)** im iCal-Parser (`reu/ical.py`):
  Derzeit werden nur Einzeltermine (`VEVENT` mit DTSTART/DTEND) gelesen.
  Wiederkehrende Termine (`RRULE`/`RDATE`) werden noch nicht expandiert.
  Lösung: `dateutil.rrule` zur Expansion im angefragten Monat verwenden
  (Dependency dann aufnehmen: `python-dateutil`).

## Mittelbar / Verbesserungen

- [ ] **requirements.txt** anlegen
  (`pyexcel-ods3 drafthorse reportlab requests icalendar pytz`,
  `lxml`/`pypdf` transitiv via drafthorse).
- [ ] **Volle PDF/A-3-Konformität**: Quell-PDF vor `attach_xml` PDF/A-konform
  erzeugen (ICC-Profil/OutputIntent hinterlegen, z.B. via Ghostscript
  oder reportlab PDF/A-Option), da drafthorse kein PDF/A aus dem Nichts baut.
- [ ] **Mustang-Validierung in CI** integrieren (Docker) laut README.
- [ ] **drafthorse-Lizenzhinweis** in README ergänzen (Apache-2.0 Code,
  proprietäre FeRD-Lizenz für mitgelieferte .xsd-Schemata).
