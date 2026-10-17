# REU – Rechnungserstellung (Stunden + Auslagen + ZUGFeRD)

Lokales CLI-Tool: iCal-Stunden + ODS-Auslagen + ODS-Kundenstammdaten
werden zu B2B-Rechnungen (immer 19 % USt) verarbeitet, mit eingebettetem
Faktur-X-XML (ZUGFeRD, Profil EN 16931). E-Mail-Versand per SMTP nur mit `--send` (Standard: kein Versand).

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

| Argument | Bedeutung | Journal/Rechnungen (iso) | Anzeige im PDF (text) |
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
  Fälligkeit (Zahlungsziel) je Kunde: Spalte `zahlungsziel_tage` der
  kunden.ods (leer = **kein** Zahlungsziel), sonst
  `[erechnung] zahlungsziel_tage` – nirgends angegeben = offen.

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
│   ├── pdfa.py            # PDF/A-3b-Konvertierung via Ghostscript (pdfwrite)
│   ├── zugferd.py          # Stufe 4: drafthorse XML + PDF/A-3-Embedding
│   ├── smtp.py             # Versand (no-op im Dry-Run)
│   ├── state.py            # Re-Nr-Kreis + Journal (Blatt 'Rechnungen' der kunden.ods)
│   └── util.py             # Hilfsfunktionen, Zeitraum-Parsing + Konstanten
├── conf/
│   ├── alice.conf          # EINZIGE Beispiel-Config
│   ├── alice-kunden.ods    # Beispiel-Kundenstammdaten (per tools/neue_kunden.py erzeugbar)
│   ├── alice-auslagen.ods  # zentrale Auslagen-ODS (per tools/neue_auslagen.py erzeugbar)
│   └── <user>.state.json    # LEGACY: wird einmalig ins Blatt 'Rechnungen' importiert
├── tools/
│   ├── neue_auslagen.py    # erzeugt conf/<user>-auslagen.ods
│   └── neue_kunden.py      # erzeugt conf/<user>-kunden.ods (Kunden + Services + Rechnungen)
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
- **[erechnung] zahlungsziel_tage ist optional** – nur ein Fallback für
  Kunden OHNE Zahlungsziel-Spalte in der kunden.ods; leer/nicht gesetzt =
  kein Zahlungsziel (siehe Kunden-ODS).
- **[erechnung] re_nr_start_<jahr> ist optional** – Startwert der laufenden
  Re-Nummer für ein Jahr, z.B. `re_nr_start_2026 = 47` → erste Re-Nr des
  Jahres ist `2026-047` (für mitten im Jahr übernommene Nummernfolgen).
  Jahre ohne Eintrag starten bei `001`. Die Rechnungsliste gewinnt, wenn
  sie schon weiter ist (Nummern werden nie zurück- oder doppelt vergeben);
  ein Wechsel des Eintrags zieht daher nur nach oben.

Die [kunden:*]-Sektionen gibt es nicht mehr – die Kundendaten liegen in der
ODS-Datei, deren Pfad in [kunden] datei steht (pro User:
conf/<user>-kunden.ods).

## Kunden-ODS conf/<user>-kunden.ods

Ein Blatt 'Kunden', eine Zeile je Kunde:

| kunde | name | strasse | plz | ort | land | ust_id | leitweg_id | stundensatz | zahlungsziel_tage |
|-------|------|---------|-----|-----|------|--------|-----------|-------------|-------------------|

**Anlegen:** `python3 tools/neue_kunden.py <user>
[--beispiel]` erzeugt die Datei mit korrektem Header (Blätter 'Kunden'
und 'Services'; Abbruch, wenn sie bereits existiert; `--force`
überschreibt).

Pflicht: kunde, name, plz, ort, land, stundensatz.
Optional: strasse, leitweg_id (B2B optional, B2G Pflicht) und **ust_id**.
Eine leere ust_id ist erlaubt: sie wird dann weder im PDF-Fuß noch im
ZUGFeRD-XML (TaxRegistration VA) ausgewiesen.
'stundensatz' als DE-Dezimal '95,00' oder EN '95.00'. **'0,00' ist erlaubt**
(Kunde ohne Honorar, z.B. nur Auslagen/Services; Stunden erscheinen dann mit
0,00 € auf der Rechnung). Negativ ist unzulässig.

**Zahlungsziel (optional), Spalte `zahlungsziel_tage`** (Spaltentitel
groß-/kleinschreibungsagnostisch, Varianten wie 'Zahlungsziel' werden
erkannt): Anzahl Tage ab Rechnungsdatum, z.B. `21`.
**Leere Zelle = kein Zahlungsziel** – die Rechnung spezifiziert dann
bewusst KEINE Fälligkeit (kein 'Zahlbar bis' im PDF-Fuß, kein BT-9 im
ZUGFeRD-XML). Fehlt die Spalte komplett, greift
`[erechnung] zahlungsziel_tage` aus der conf (dort leer = ebenfalls
kein Zahlungsziel). Die kunden.ods-Spalte hat also Vorrang.

**Blatt 'Services' (optional, eine Zeile je Kunde und Service):**

| kunde | service | Kosten pro Stück [€] | 1 | 2 | 3 | … | 12 |
|-------|---------|----------------------|---|---|---|---|----|

Die Spalten '1'–'12' sind die Monate des Jahres. Zellwert je Monat:

- **Zahl** (z. B. `2`) = Anzahl der Serviceeinheiten in diesem Monat
- **Markierung** (`x`, `ja`, `yes`, `j`, `y`, `ok`, `wahr`, `true`, `✓`
  oder Wahrheitszelle) = **1 Einheit** – der Service gilt für diesen
  markierten Monat
- **leer** = 0 (Monat nicht abgerechnet)
- Andere Werte sind ein **FEHLER** – das Gewerk darf nicht stillschweigend
  als Menge 0 von der Rechnung verschwinden.

Spaltentitel sind groß-/kleinschreibungsagnostisch ('Kunde' wie 'kunde',
numerische Titel wie '7.0' werden als Monat 7 erkannt). Pro Service und
Monat mit Menge > 0 (oder Markierung) entsteht eine eigene
Rechnungsposition auf der Rechnungs-Erstseite (und im ZUGFeRD-XML) –
**jeder markierte Monat innerhalb des Zeitraums im Aufruf**. Kunden mit
nur Service-Einheiten (ohne Stunden/Auslagen) bekommen ebenfalls eine
Rechnung.

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

| datum | kunde | art | bezeichnung | betrag_netto | belegnr | freigabe | währung |
|-------|-------|-----|-------------|--------------|---------|----------|---------|

**Währung (optional):** Ist in irgendeiner freigegebenen Zeile eine
Währung angegeben (`währung`, `waehrung` oder `currency`; `€` gilt als
EUR), führt die Rechnung diese Währung – für Beträge im PDF, das
ZUGFeRD-XML (BT-5) und die Steuerwährung. Fehlt die Spalte ganz oder
ist die Zelle leer: **EUR** (Anzeige `€`). Widersprüchliche Währungen
in einer Rechnung sind ein Fehler (eine Rechnung darf nur eine Währung
führen).

**Beleg-Nr. als Text formatieren:** Die Spalte `belegnr` ist im
Template `tools/neue_auslagen.py` als **Textzelle (Format `@`)** mit
Beispielwert `1302003929154` hinterlegt. Bitte neue Belegnummern
ebenso als Text erfassen – sonst speichert LibreOffice rein numerische
Nummern als **Zahl**, und es hängt sich ein `.0` an (z.B.
`1302003929154.0`). REU normalisiert beide Formen (Zahl und
verunglückter Text wie `1302003929154.0`) zwar automatisch, mit
Text-Format ist die Datei aber sauber. Bestehende Datei in
LibreOffice: Beleg-Nr.-Spalte markieren → Format → Zellen → Text
(Formatcode `@`).

Pflichtspalten konfigurierbar (pflichtspalten); Standard wie oben (ohne
freigabe). 'betrag_netto' als '42,00' oder '42.00'. USt-Satz-Spalte entfällt (immer 19 %).
**Spaltentitel sind groß-/kleinschreibungsagnostisch** – 'Freigabe' wie
'freigabe', 'Datum' wie 'datum' (wie beim Services-Blatt).

**Freigabe zeilenweise** über die Spalte `freigabe` (Name konfigurierbar
über [auslagen] freigabe_spalte, Default `freigabe`):

- `yes` oder `ja` → Zeile wird abgerechnet (Groß-/Kleinschreibung egal).
- leer/`no`/`nein` → Zeile wird **übersprungen** (Ausgabe: WARNUNG je Zeile).
- Die Rechnung ist **finale** (echte Re-Nr, Eintrag ins Blatt 'Rechnungen', ZUGFeRD; Versand nur mit `--send`), sobald
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

**Zeitraum-Bezug der Gewerke (im Gegensatz zu Auslagen):** Services
werden für jeden markierten Monat **innerhalb des Zeitraums im Aufruf**
abgerechnet, Arbeitsstunden folgen ihrem **Termin-Datum** (Datum liegt im
Zeitraum). Liegen im Aufruf-Zeitraum keine Service-Monate und/oder keine
Termine mit Kundenkürzel vor, meldet das CLI das ausdrücklich als Hinweis
(„keine markierten Service-Monate im Zeitraum …“, „keine Termine mit
Kundenkürzel …“) – ein Gewerk fehlt nie stillschweigend auf der Rechnung.

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

# Rechnungsentwurf (keine Zeile mit freigabe=yes → ENTWURF, kein XML, kein Versand, keine Re-Nr)
./run-reu.py alice 11 2025 --invoice

# QUARTALSABRECHNUNG: Oktober bis Dezember in einer Rechnung je Kunde
./run-reu.py alice Q4 2025 --full

# Beliebiger Monatsbereich (z.B. November + Dezember)
./run-reu.py alice 11-12 2025 --full

# Rechnung zum Prüfen mit eingebettetem ZUGFeRD-XML, aber OHNE Versand
# (freigabe=yes in mindestens einer Zeile vorausgesetzt). Re-Nr bleibt ENTWURF-…, Blatt 'Rechnungen' unverändert.
./run-reu.py alice Q1 2025 --full --dry-run

# Finale Rechnung + ZUGFeRD + Journal, OHNE E-Mail-Versand (Standard!)
./run-reu.py alice 11 2025 --full

# Finale Rechnung + ZUGFeRD + Journal + E-Mail-Versand (freigabe=yes vorausgesetzt)
./run-reu.py alice 11 2025 --full --send

# Rechnungsübersicht aus dem Blatt 'Rechnungen' der kunden.ods (fürs Finanzamt)
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
| Re-Nr vergeben + Rechnungen-Blatt schreiben | ✓ | ✗ — verwendet ENTWURF-… |
| E-Mail-Versand | nur mit `--send` | ✗ — komplett übersprungen |
| Ausgabe in | out/final/ | out/dry-run/ |

Die Freigabe-Spalte und --dry-run sind orthogonal: freigabe steuert Entwurf-vs-finale-Logik
(XML ja/nein), --dry-run steuert Seiteneffekte (Versand, Rechnungen-Blatt, Re-Nr). So kann
auch bei versehentlich freigegebenen Zeilen nichts nach draußen gehen und
keine Nummer verbraucht werden.

## Rechnungsnummer und Rechnungsliste (Blatt 'Rechnungen')

Nummernkreis und Journal werden im Blatt **Rechnungen** der
conf/<user>-kunden.ods gepflegt (ersetzte die frühere
conf/<user>.state.json). Schema (eine Zeile je Rechnung):

```
re_nr | datum | kunde | leistungszeitraum | netto | ust | brutto | status
```

- **Nummernkreis:** Format YYYY-NNN, fortlaufend pro Jahr; die nächste
  Nummer ist die **höchste vorhandene Nummer des Jahres + 1** – der Kreis
  wird aus der Liste abgeleitet und kann nie zurück- oder doppelt
  vergeben werden. Jahreswechsel setzt den Kreis zurück.
- **Pflegen:** Die Liste ist direkt editierbar – z.B. Status `offen` →
  `bezahlt` setzen; `--journal` liest das Blatt (Zusatzspalten bleiben
  unberührt, neue Zeilen bekommen dort leere Zellen). Fehlt das Blatt,
  legt REU es beim ersten echten Lauf an (Kopfzeile inklusive).
- **Migration:** Existiert noch eine conf/<user>.state.json, wird sie
  beim ersten echten Lauf einmalig importiert und in
  <user>.state.json.importiert umbenannt (Dry-Run lässt sie liegen).
- **Entwurf/Dry-Run:** ENTWURF-YYYY-MM-DD-HHMMSS (nicht persistent).

## Abhängigkeiten

```
(ODS lesen)    # entfällt – reu/ods.py nutzt zipfile + xml (Python-stdlib)
drafthorse     # Faktur-X XML (EN 16931) + PDF/A-3-Embedding
pypdf          # (transitiv über drafthorse)
reportlab      # Rechnungs-PDF
ghostscript    # PDF/A-3b-Konvertierung des Rechnungs-PDF (re/pdfa.py)
requests icalendar pytz  # iCal-Download/Parse
lxml           # XML-Validierung (transitiv über drafthorse)
```

Installieren (Debian/Ubuntu ohne pip):

```bash
apt install ghostscript python3-drafthorse python3-reportlab python3-icalendar \
            python3-pytz python3-requests python3-lxml python3-pypdf
```

Hinweis: `python3-drafthorse` gibt es ab Debian 13 (trixie) bzw.
Ubuntu 25.10; auf älteren Releases bleibt für drafthorse nur pip:

```bash
python3 -m pip install 'drafthorse>=2.2' reportlab requests icalendar pytz
```

**drafthorse-Version:** REU unterstützt drafthorse **2.2 bis 2025.x**
(die Modell-API hat sich zwischen den Versionen geändert – `reu/zugferd.py`
gleicht das intern aus: CurrencyField-2-Tupel vs. DecimalField,
MultiStringField vs. StringField für Notizen, Field vs. MultiField für
PaymentMeans). drafthorse **1.x** nutzt die inkompatible ZUGFeRD-1.0-API
und wird mit einer klaren Fehlermeldung abgewiesen:

```bash
python3 -m pip install -U 'drafthorse>=2.2'
```

## iCal-Feed

[ical] url kann 'https://…', 'http://…' oder 'file://…' sein.
Ein Termineintrag braucht ein Kundenkürzel im SUMMARY-Präfix: **exakt
drei Zeichen, direkt gefolgt von einem Doppelpunkt** ('BBB: …' – nicht
'AAAA:' oder 'CC:'), oder ein Kürzel in der CATEGORIES-Eigenschaft.
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
validiert (via drafthorse + lxml). Das sichtbare PDF wird dagegen durch
REU selbst PDF/A-3b-konform erzeugt (siehe unten) und kann zusätzlich
mit veraPDF geprüft werden:

```bash
verapdf out/final/ABC_Rechnung_2025-001_ZUGFeRD.pdf
```

### PDF/A-3-Pipeline

drafthorse setzt nur das PDF/A-3-Flag und die AF-Relation; es macht ein
beliebiges Quell-PDF aber **nicht** konform. REU löst das daher zweistufig:

1. **`reu/pdfa.py`**: Das reportlab-PDF wird vor dem Embedding per
   Ghostscript nach **PDF/A-3b** konvertiert:
   `gs -sDEVICE=pdfwrite -dPDFA=3 -sColorConversionStrategy=RGB
   -dPDFACompatibilityPolicy=1 --permit-file-read=<srgb.icc>
   -o out.pdf <gepatchte PDFA_def.ps> in.pdf`
   Die mitgelieferte `PDFA_def.ps` wird automatisch gesucht (Debian/Ubuntu:
   `/usr/share/ghostscript/<version>/lib/`), die `/ICCProfile`-Zeile wird
   auf den gefundenen sRGB-ICC-Pfad umgebogen. `-dPDFACompatibilityPolicy=1`
   sorgt dafür, dass gs bei nicht konvertierbarem Inhalt **mit Fehler
   abbricht**, statt stillschweigend ein nicht-konformes PDF zu erzeugen.
2. **`reu/zugferd.py`**: drafthorse bettet das Faktur-X-XML in das bereits
   konforme PDF ein (XMP `pdfaid:part=3/B`, Factur-X-Extension,
   AFRelationship `/Alternative`, Dateiname `factur-x.xml`).

Konfiguration in `[erechnung]`:

| Option | Default | Bedeutung |
|---|---|---|
| `format` | `pdfa3` | `pdfa3` = Ghostscript-Konvertierung vor dem Embedding; jeder andere Wert überspringet sie |
| `gs_bin` | `gs` | Pfad zum Ghostscript-Binary |
| `icc_profil` | leer = auto | sRGB-ICC-Profil; leer = automatische Suche (`.../iccprofiles/srgb.icc`, `default_rgb.icc`, `/usr/share/color/icc/colord/sRGB.icc`) |
| `pdfa_def` | leer = auto | Pfad zur `PDFA_def.ps`; leer = automatische Suche im Ghostscript-lib-Verzeichnis |

Tritt bei der Konvertierung ein Fehler auf (gs fehlt, ICC-Profil nicht
gefunden, gs bricht ab), bricht der Lauf mit einer klaren Meldung ab –
es wird **nie** eine ZUGFeRD-PDF ohne PDF/A-3-Konvertierung erzeugt.

Release-Gate: Vor dem ersten echten Versand die erzeugte PDF einmalig mit
**veraPDF** (PDF/A) und **Mustang** (Faktur-X gesamt) validieren (s. o.).
