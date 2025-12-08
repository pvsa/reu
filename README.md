# reu
Rechnungs Erstellungs Unterstützung

Anforderungen:
Das Tool läd von einer URL einen Kalender herunter und extrahiert alle Termine deren Titel die mit dem Pattern "ABC:" (also drei große Buchstaben gefolgt ovn einem Doppelpunkt) beginnen. Diese Pattern stellen unterschiedliche Rechnungsempfänger/kunden dar.
Aus den Dauern der Termine erstellt es Arbeitszeiteinträge - gruppoert nach den unsterschiedlichen Pattern/Kunden. Also Alle "ABC" in eine Abrechnung, alle "DEF" in eine und so weiter.
Begrenztr wird die Suche und Ausgabe auf einen Monat.

Python Biliotheken:
 - argparse
 - configparser
 - os
 - sys
 - re
 - smtplib
 - tempfile
 - pytz


Nutzung:
Um Download der ics Datei und den Versand - nach Erstellung - per Mail zu realisieren, gibt es eine conf-Datei unter conf/:

./run-reu.py [-h] username month year





