# REU – Rechnungserstellung (Stunden + Auslagen + ZUGFeRD)

Lokales CLI-Tool: iCal-Stunden + ODS-Auslagen + ODS-Kundenstammdaten
werden zu B2B-Rechnungen (immer 19 % USt) verarbeitet, mit eingebettetem
Faktur-X-XML (ZUGFeRD, Profil EN 16931). Versand per SMTP im finalen Lauf.

## Fixierte Parameter

| Punkt | Festlegung |
|-------|-----------|
| Empfänger | B2B, regelbesteuert |
| USt | **immer 19 %** (Stunden + Auslagen) |
| Stundenposition | **eine Sammelposition** je Kunde und Zeitraum |
| Format | **ZUGFeRD** (PDF/A-3 mit eingebettetem Faktur-X-XML, EN 16931) |
| Kundendaten | **pro User** in conf/<user>-kunden.ods |
| Auslagen | lokale ODS, Pfad mit {year}/{month} (pro Monat, wird gemerged) |
| Zeitraum | Einzelmonat, Monatsbereich oder Quartal |
| Ausführung | lokal beim Rechnungsersteller |

## Zeiträume

Das CLI-Argument 'month' erlaubt drei Formen:

| Argument | Bedeutung | Journal/State (iso) | Anzeige im PDF (text) |
|----------|-----------|---------------------|------------------------|
| 11       | Einzelmonat November | 2025-11 | 11/2025 |
| 10-12    | Monatsbereich Okt–Dez | 2025-10-2025-12 | 10/2025-12/2025 |
| Q4       | Quartal (Okt–Dez) | 2025-Q4 | Q4/2025 |

Bei Mehrmonats-Zeiträumen gilt:

- **iCal**: alle Termine vom ersten Tag des ersten bis zum letzten Tag des
  letzten Monats (Zeitzonen werden normalisiert).
- **Auslagen**: die Monats-ODSen werden je Monat geladen ({year}/{month}-
  Template) und zusammengefasst. Fehlende Monatsdateien sind erlaubt,
  mindestens eine muss existieren. Freigabe (Meta!A1 = yes) muss in
  **allen** vorhandenen Monatsdateien gesetzt sein.
- **Rechnungsdatum** = letzter Tag des letzten Monats des Zeitraums,
  Zahlungsziel ab diesem Datum.

## Modulstruktur

```
reu/
├── run-reu.py              # CLI-Dispatch
├── reu/
│   ├── __init__.py
│   ├── config.py           # lädt conf/<user>.conf (INI)
│   ├── kunden.py           # Kundenstammdaten aus <user>-kunden.ods
│   ├── ods.py              # ODS-Lesehilfe (Python-stdlib, ohne Zusatzpakete)
│   ├── ical.py             # Stufe 1: Stunden aus iCal → dict
│   ├── expenses.py         # Stufe 2: Auslagen aus ODS + Freigabe-Check
│   ├── invoice.py          # Stufe 3: Rechnungs-PDF (reportlab)
│   ├── zugferd.py          # Stufe 4: drafthorse XML + PDF/A-3-Embedding
│   ├── smtp.py             # Versand (no-op im Dry-Run)
│   ├── state.py            # Re-Nr-Kreis + Journal (JSON)
│   └── util.py             # Hilfsfunktionen, Zeitraum-Parsing + Konstanten
├── conf/
│   ├── alice.conf          # EINZIGE Beispiel-Config
│   ├── alice-kunden.ods    # Beispiel-Kundenstammdaten
│   └── alice.state.json    # (wird beim finalen Lauf angelegt)
├── out/
│   ├── final/              # finale Rechnungen
│   └── dry-run/            # Entwürfe / Dry-Run-Ausgaben
└── README.md
```

## Config conf/<user>.conf (INI)

Sektionen: [ical], [smtp], [pdf], [leistender], [erechnung], [kunden],
[auslagen]. Siehe conf/alice.conf (einzige Beispiel-Config).

- **[smtp] smtp_username / smtp_password sind optional** – nur setzen, wenn
  der SMTP-Server eine Authentifizierung verlangt (ohne = Relay ohne Login).
- **[leistender]** führt keine Steuernummer mehr, nur ust_id.

Die [kunden:*]-Sektionen gibt es nicht mehr – die Kundendaten liegen in der
ODS-Datei, deren Pfad in [kunden] datei steht (pro User:
conf/<user>-kunden.ods).

## Kunden-ODS conf/<user>-kunden.ods

Ein Blatt 'Kunden', eine Zeile je Kunde:

| kunde | name | strasse | plz | ort | land | ust_id | leitweg_id | stundensatz |
|-------|------|---------|-----|-----|------|--------|-----------|-------------|

Pflicht: kunde, name, plz, ort, land, stundensatz.
Optional: strasse, leitweg_id (B2B optional, B2G Pflicht) und **ust_id**.
Eine leere ust_id ist erlaubt: sie wird dann weder im PDF-Fuß noch im
ZUGFeRD-XML (TaxRegistration VA) ausgewiesen.
'stundensatz' als DE-Dezimal '95,00' oder EN '95.00'.

## Auslagen-ODS conf/auslagen_{year}-{month:02d}.ods

Pfad-Template aus [auslagen] datei mit {year}/{month} – **eine Datei je
Monat**. Bei Mehrmonats-Zeiträumen werden alle vorhandenen Monatsdateien
zusammengefasst.

**Blatt 'Auslagen'** (eine Zeile pro Beleg):

| datum | kunde | art | bezeichnung | betrag_netto | belegnr |
|-------|-------|-----|-------------|--------------|---------|

Pflichtspalten konfigurierbar (pflichtspalten); Standard wie oben.
'betrag_netto' als '42,00' oder '42.00'. USt-Satz-Spalte entfällt (immer 19 %).

**Blatt 'Meta'** (Freigabe-Signal):
- A1 = 'no' (Standard) → nur Entwurf, kein ZUGFeRD-XML, keine echte Re-Nr.
- A1 = 'yes' → finale Rechnung, echte Re-Nr, State aktualisiert, ZUGFeRD erzeugt.

## CLI

```bash
# Nur Stunden testen (früh im Monat)
./run-reu.py alice 11 2025 --hours

# Nur Auslagen aus ODS prüfen
./run-reu.py alice 11 2025 --expenses

# Rechnungsentwurf (Meta!A1 ≠ yes → ENTWURF, kein XML, kein Versand, kein State)
./run-reu.py alice 11 2025 --invoice

# QUARTALSABRECHNUNG: Oktober bis Dezember in einer Rechnung je Kunde
./run-reu.py alice Q4 2025 --full

# Beliebiger Monatsbereich (z.B. November + Dezember)
./run-reu.py alice 11-12 2025 --full

# Rechnung zum Prüfen mit eingebettetem ZUGFeRD-XML, aber OHNE Versand
# (Meta!A1 = yes vorausgesetzt). Re-Nr bleibt ENTWURF-…, kein State.
./run-reu.py alice Q1 2025 --full --dry-run

# Finale Rechnung + ZUGFeRD + Versand + Journal (Meta!A1 = yes vorausgesetzt)
./run-reu.py alice 11 2025 --full

# Rechnungsübersicht aus State (fürs Finanzamt)
./run-reu.py alice --journal --journal-year 2025

# Nur einen Kunden verarbeiten
./run-reu.py alice Q4 2025 --invoice --kunde ABC
```

**Ohne Action-Flag** (--hours/--expenses/--invoice/--full/--journal)
wird nur die Hilfe ausgegeben – keine Ausführung.

### Dry-Run (--dry-run)

| Aktion | normal | --dry-run |
|--------|--------|-----------|
| Rechnungs-PDF erzeugen | ✓ | ✓ |
| ZUGFeRD-XML + PDF/A-3-Embedding (nur bei Meta=yes) | ✓ | ✓ — damit PDF+XML prüfbar |
| Re-Nr vergeben + State schreiben | ✓ | ✗ — verwendet ENTWURF-… |
| E-Mail-Versand | ✓ | ✗ — komplett übersprungen |
| Ausgabe in | out/final/ | out/dry-run/ |

Meta!A1 und --dry-run sind orthogonal: Meta!A1 steuert Entwurf-vs-finale-Logik
(XML ja/nein), --dry-run steuert Seiteneffekte (Versand, State, Re-Nr). So kann
auch bei versehentlich gesetztem Meta!A1 = yes nichts nach draußen gehen und
keine Nummer verbraucht werden.

## Rechnungsnummer

Aus conf/<user>.state.json, Format YYYY-NNN (fortlaufend pro Jahr),
Jahreswechsel setzt den Kreis zurück. Bei Entwurf/Dry-Run:
ENTWURF-YYYY-MM-DD-HHMMSS (nicht persistent).

## Abhängigkeiten

```
(ODS lesen)    # entfällt – reu/ods.py nutzt zipfile + xml (Python-stdlib)
drafthorse     # Faktur-X XML (EN 16931) + PDF/A-3-Embedding
pypdf          # (transitiv über drafthorse)
reportlab      # Rechnungs-PDF
requests icalendar pytz  # iCal-Download/Parse
lxml           # XML-Validierung (transitiv über drafthorse)
```

Installieren (Debian/Ubuntu ohne pip):

```bash
apt install python3-drafthorse python3-reportlab python3-icalendar \
            python3-pytz python3-requests python3-lxml python3-pypdf
```

Hinweis: `python3-drafthorse` gibt es ab Debian 13 (trixie) bzw.
Ubuntu 25.10; auf älteren Releases bleibt für drafthorse nur pip:

```bash
python3 -m pip install drafthorse reportlab requests icalendar pytz
```

## iCal-Feed

[ical] url kann 'https://…', 'http://…' oder 'file://…' sein.
Ein Termineintrag braucht ein Kundenkürzel im SUMMARY-Präfix: **exakt
drei Zeichen, direkt gefolgt von einem Doppelpunkt** ('ZTR: …' – nicht
'NOD25:' oder 'HA:'), oder ein Kürzel in der CATEGORIES-Eigenschaft.
Termine außerhalb des angefragten
Zeitraums werden ignoriert. Termine **ohne** Kürzel (z.B. interne/
technische Einträge wie Backups) gelten als nicht abrechenbar und werden
stillschweigend übersprungen – der Lauf bricht nicht ab.

## Validierung der ZUGFeRD-PDF (Pflicht)

Die erzeugte ZUGFeRD-PDF muss vor Versand validiert werden, z.B. mit
Mustang (offiziell):

```bash
docker run --rm -v "$PWD/out/final:/data" mustangproject/cli \
  validate /data/ABC_Rechnung_2025-001_ZUGFeRD.pdf
```

oder lokal:

```bash
java -jar Mustang-CLI.jar --action validate --source out/final/ABC_Rechnung_2025-001_ZUGFeRD.pdf
```

Das eingebettete XML wird beim Erzeugen bereits gegen das EN-16931-XSD
validiert (via drafthorse + lxml); die PDF/A-3-Konformität des sichtbaren
PDFs sollte zusätzlich via Mustang/VeraPDF geprüft werden.

### Hinweis PDF/A-3

drafthorse setzt das PDF/A-3-Flag und die AF-Relation, benötigt dafür aber
ein Quell-PDF mit OutputIntent/ICC-Profil. reportlab erzeugt standardmäßig
kein PDF/A. Für volle PDF/A-3-Konformität das Quell-PDF entsprechend
erzeugen (ICC-Profil hinterlegen) oder das PDF nachträglich konvertieren.
Das eingebettete Faktur-X-XML ist in jedem Fall gültig (EN 16931).
