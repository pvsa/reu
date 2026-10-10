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
| Auslagen | **zentrale ODS** je User: conf/<user>-auslagen.ods (Zuordnung via Freigabe-Spalte) |
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
- **Auslagen**: **eine zentrale Datei** `<user>-auslagen.ods` – die Spalte
  `freigabe` entscheidet, welche Zeilen abgerechnet werden (Datum-unabhängig,
  siehe Abschnitt Auslagen-ODS). Existiert die Datei **nicht**, fragt das
  CLI **zu Beginn** um Bestätigung – **Default ist Fortfahren** (Enter oder
  `j`); nur `n`/`nein` bricht ab. Ansonsten weist das CLI generell auf die
  Auslagenlage hin (Zeilen gesamt, abgerechnet, übersprungen).
  Freigabe zeilenweise über die Spalte `freigabe` (siehe unten).
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
│   ├── services.py         # Service-Positionen aus Blatt 'Services' der kunden.ods
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
│   ├── alice-auslagen.ods  # zentrale Auslagen-ODS (per tools/neue_auslagen.py erzeugbar)
│   └── alice.state.json    # (wird beim finalen Lauf angelegt)
├── tools/
│   └── neue_auslagen.py    # erzeugt conf/<user>-auslagen.ods
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
'stundensatz' als DE-Dezimal '95,00' oder EN '95.00'. **'0,00' ist erlaubt**
(Kunde ohne Honorar, z.B. nur Auslagen/Services; Stunden erscheinen dann mit
0,00 € auf der Rechnung). Negativ ist unzulässig.

**Blatt 'Services' (optional, eine Zeile je Kunde und Service):**

| kunde | service | Kosten pro Stück [€] | 1 | 2 | 3 | … | 12 |
|-------|---------|----------------------|---|---|---|---|----|

Die Spalten '1'–'12' sind die Monate des Jahres; der Zellwert ist die
Anzahl der Serviceeinheiten in diesem Monat (leer = 0). Spaltentitel sind
groß-/kleinschreibungsagnostisch ('Kunde' wie 'kunde'). Pro Service und
Monat mit Menge > 0 entsteht eine eigene Rechnungsposition auf der
Rechnungs-Erstseite (und im ZUGFeRD-XML). Kunden mit nur Service-Einheiten
(ohne Stunden/Auslagen) bekommen ebenfalls eine Rechnung.

## Auslagen-ODS conf/<user>-auslagen.ods

**Anlegen:** `python3 tools/neue_auslagen.py <user>
[--beispiel]` erzeugt die Datei mit korrektem Header (Abbruch, wenn sie
bereits existiert; `--force` überschreibt; `--beispiel` fügt Alice-Demozeilen
mit allen Freigabe-Zuständen ein).

**Eine zentrale Datei je User** ([auslagen] datei, z.B.
`conf/philipp-auslagen.ods`), wie die `<user>-kunden.ods`. **Die Spalte `datum` ist reine Beleginformation** – die Zuordnung läuft
allein über die Spalte `freigabe`: Jede freigegebene Zeile (`ja`/`yes`)
fließt in die nächste Rechnung, unabhängig von Datum und Abrechnungs-
zeitraum. Alle Zeilen bleiben in der Datei liegen und werden fortlaufend
ergänzt. Datum als ODS-Datumzelle, `TT.MM.JJJJ` oder `JJJJ-MM-TT`
(leer/ungültig ist erlaubt).

**Blatt 'Auslagen'** (eine Zeile pro Beleg):

| datum | kunde | art | bezeichnung | betrag_netto | belegnr | freigabe |
|-------|-------|-----|-------------|--------------|---------|----------|

Pflichtspalten konfigurierbar (pflichtspalten); Standard wie oben (ohne
freigabe). 'betrag_netto' als '42,00' oder '42.00'. USt-Satz-Spalte entfällt (immer 19 %).
**Spaltentitel sind groß-/kleinschreibungsagnostisch** – 'Freigabe' wie
'freigabe', 'Datum' wie 'datum' (wie beim Services-Blatt).

**Freigabe zeilenweise** über die Spalte `freigabe` (Name konfigurierbar
über [auslagen] freigabe_spalte, Default `freigabe`):

- `yes` oder `ja` → Zeile wird abgerechnet (Groß-/Kleinschreibung egal).
- leer/`no`/`nein` → Zeile wird **übersprungen** (Ausgabe: WARNUNG je Zeile).
- Die Rechnung ist **finale** (echte Re-Nr, State, ZUGFeRD, Versand), sobald
  **mindestens eine** Zeile der Datei `yes`/`ja` hat – sonst ENTWURF.
- Fehlt die Spalte komplett: keine Zeile freigegeben → ENTWURF + WARNUNG.

**Erledigt-Vermerk (Schutz vor doppelter Abrechnung):** Nach einem finalen
Lauf (`--invoice`/`--full`, ohne `--dry-run`) trägt REU die **Rechnungsnummer
in die Freigabe-Spalte** ALLER abgerechneten Zeilen ein (`ja` → `2026-001`,
datum-unabhängig).
Diese Zeilen gelten bei künftigen Läufen als **bereits abgerechnet**:
Sie werden nicht erneut abgerechnet und nicht als nicht freigegeben
angemeckert (in der Zusammenfassung als „bereits abgerechnet" gezählt,
`--expenses` listet sie mit Re-Nr). FINALE bleibt an eine aktive
`yes`/`ja`-Freigabe gebunden. Das Eintragen erfolgt atomar – hat
LibreOffice die Datei gerade geöffnet, schlägt es mit FEHLER-Hinweis fehl
und der Lauf bleibt unberührt.

## Kunden-Berücksichtigung je Zeitraum

**Jeder Kunde, der eines der Gewerke genutzt hat (Stunden/Services im
Zeitraum, freigegebene Auslagen datum-unabhängig), wird berücksichtigt** –
egal welches:

- **Arbeitsstunden** (iCal-Termine mit Kundenkürzel im Zeitraum)
- **Services** (Blatt 'Services', Menge > 0 in einem Zeitraum-Monat)
- **Auslagen** (freigegebene Zeilen der zentralen Auslagen-ODS – datum-unabhängig)

Auch Kunden mit **nur einem** Gewerk (nur Stunden, nur Services oder nur
Auslagen) bekommen eine eigene Rechnung. Sonderfall Auslagen: Die
Rechnung entsteht erst, sobald mindestens eine Zeile des Kunden
`freigabe` = `yes`/`ja` hat. Hat ein Kunde nur **offene** (nicht
freigegebene) Auslagen und sonst nichts, erhält er
noch **keine** Rechnung – er wird stattdessen ausdrücklich gemeldet:
`Hinweis: <kunde> – N nicht freigegebene Auslagenzeile(n) … noch keine
Rechnung, bis freigabe=yes/ja`. Kein Kunde verschwindet stillschweigend.

## Rechnungsaufbau (PDF)

- **Briefkopf (Erstseite)**: Logo oben links am Satzspiegelrand; rechts
  daneben Absenderadresse + Rechnungsdatum/-nummer/Leistungszeitraum,
  **oben rechts auf gleicher Höhe wie das Logo** (Oberkante an der
  oberen Satzspiegelkante, rechtsbündig). Empfängeradresse darunter links.
- **Erstseite**: Positionstabelle + Summenblock. Spalten:
  `Pos. | Bezeichnung | Menge | Einzelpreis | Netto` (ohne Einheiten-Spalte;
  Einheiten stehen nur im ZUGFeRD-XML). Lange Bezeichnungen brechen
  automatisch mehrzeilig in der Spalte um. Die Positionstexte
  verweisen auf die Anlagen: Stunden-Sammelposition mit
  „(siehe Anlage Arbeitsstunden)", Services je Service und Monat,
  Auslagen als **eine** Sammelposition „Auslagen (siehe Anlage Auslagen)".
- **Anlage: Arbeitsstunden** – Datum | Start | Beschreibung | Dauer (h)
  (nur wenn Stunden vorhanden)
- **Anlage: Auslagen** – Datum | Art | Bezeichnung | Beleg-Nr. | Betrag
  (nur wenn Auslagen freigegeben vorhanden)
- Das ZUGFeRD-XML enthält dieselben Positionen (Stunden, Services je
  Monat, eine Auslagen-Sammelposition) – konsistent zur PDF.

## CLI

```bash
# Nur Stunden testen (früh im Monat)
./run-reu.py alice 11 2025 --hours

# Nur Auslagen aus ODS prüfen
./run-reu.py alice 11 2025 --expenses

# Rechnungsentwurf (keine Zeile mit freigabe=yes → ENTWURF, kein XML, kein Versand, kein State)
./run-reu.py alice 11 2025 --invoice

# QUARTALSABRECHNUNG: Oktober bis Dezember in einer Rechnung je Kunde
./run-reu.py alice Q4 2025 --full

# Beliebiger Monatsbereich (z.B. November + Dezember)
./run-reu.py alice 11-12 2025 --full

# Rechnung zum Prüfen mit eingebettetem ZUGFeRD-XML, aber OHNE Versand
# (freigabe=yes in mindestens einer Zeile vorausgesetzt). Re-Nr bleibt ENTWURF-…, kein State.
./run-reu.py alice Q1 2025 --full --dry-run

# Finale Rechnung + ZUGFeRD + Versand + Journal (freigabe=yes vorausgesetzt)
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
| ZUGFeRD-XML + PDF/A-3-Embedding (nur bei freigabe=yes) | ✓ | ✓ — damit PDF+XML prüfbar |
| Re-Nr vergeben + State schreiben | ✓ | ✗ — verwendet ENTWURF-… |
| E-Mail-Versand | ✓ | ✗ — komplett übersprungen |
| Ausgabe in | out/final/ | out/dry-run/ |

Die Freigabe-Spalte und --dry-run sind orthogonal: freigabe steuert Entwurf-vs-finale-Logik
(XML ja/nein), --dry-run steuert Seiteneffekte (Versand, State, Re-Nr). So kann
auch bei versehentlich freigegebenen Zeilen nichts nach draußen gehen und
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
