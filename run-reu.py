#!/usr/bin/env python3
"""
iCal zu PDF Rechnungsgenerator
Erstellt monatliche Rechnungen aus iCal-Dateien für verschiedene Kunden
"""

import argparse
import configparser
import os
import sys
import re
from datetime import datetime, timedelta
from collections import defaultdict
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText
import tempfile
import pytz
from zoneinfo import ZoneInfo

# Externe Bibliotheken (müssen installiert werden)
try:
    import requests
    from icalendar import Calendar, Event
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.lib.units import cm
    from reportlab.lib.colors import black
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
except ImportError as e:
    print(f"Fehler beim Importieren der Bibliotheken: {e}")
    print("Bitte installieren Sie die erforderlichen Pakete:")
    print("pip install requests icalendar reportlab")
    sys.exit(1)


class InvoiceGenerator:
    def __init__(self, config_file, username, month, year):
        self.username = username
        self.month = month
        self.year = year
        self.config = self.load_config(config_file)
        self.local_tz = ZoneInfo("Europe/Berlin")  # Deutsche Zeitzone
        
    def load_config(self, config_file):
        """Lädt die Konfigurationsdatei für den Benutzer"""
        config = configparser.ConfigParser()
        try:
            config.read(config_file)
            return config
        except Exception as e:
            print(f"Fehler beim Laden der Konfiguration: {e}")
            sys.exit(1)
    
    def download_ical(self):
        """Lädt die iCal-Datei von der konfigurierten URL herunter"""
        try:
            url = self.config['ical']['url']
            username = self.config['ical']['username']
            password = self.config['ical']['password']
            
            # Verschiedene Authentifizierungsmethoden versuchen
            auth_method = self.config.get('ical', 'auth_method', fallback='basic')
            
            # Konfigurierbares Timeout (Standard: 60 Sekunden)
            timeout = int(self.config.get('ical', 'timeout', fallback='60'))
            
            # Retry-Logik
            max_retries = int(self.config.get('ical', 'max_retries', fallback='3'))
            retry_delay = int(self.config.get('ical', 'retry_delay', fallback='5'))
            
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            
            print(f"Verbinde zu: {url}")
            print(f"Authentifizierungsmethode: {auth_method}")
            print(f"Timeout: {timeout} Sekunden")
            
            for attempt in range(max_retries):
                try:
                    print(f"Versuch {attempt + 1}/{max_retries}...")
                    
                    if auth_method.lower() == 'basic':
                        # Standard HTTP Basic Auth
                        response = requests.get(url, auth=(username, password), headers=headers, timeout=timeout)
                    elif auth_method.lower() == 'digest':
                        # HTTP Digest Auth
                        from requests.auth import HTTPDigestAuth
                        response = requests.get(url, auth=HTTPDigestAuth(username, password), headers=headers, timeout=timeout)
                    elif auth_method.lower() == 'token':
                        # Token-basierte Authentifizierung (z.B. für CalDAV)
                        headers['Authorization'] = f'Bearer {password}'
                        response = requests.get(url, headers=headers, timeout=timeout)
                    elif auth_method.lower() == 'url_params':
                        # Parameter in der URL
                        params = {'username': username, 'password': password}
                        response = requests.get(url, params=params, headers=headers, timeout=timeout)
                    else:
                        # Fallback zu Basic Auth
                        response = requests.get(url, auth=(username, password), headers=headers, timeout=timeout)
                    
                    print(f"HTTP Status: {response.status_code}")
                    
                    if response.status_code == 401:
                        print("Authentifizierung fehlgeschlagen!")
                        print("Mögliche Lösungen:")
                        print("1. Überprüfen Sie Benutzername und Passwort")
                        print("2. Versuchen Sie eine andere auth_method in der Konfiguration:")
                        print("   - basic (Standard)")
                        print("   - digest")
                        print("   - token")
                        print("   - url_params")
                        print("3. Prüfen Sie ob die URL korrekt ist")
                        print("4. Möglicherweise ist ein App-spezifisches Passwort erforderlich")
                        sys.exit(1)
                    elif response.status_code == 404:
                        print("Fehler 404: Die angegebene URL wurde nicht gefunden")
                        print("Bitte überprüfen Sie die URL in der Konfigurationsdatei")
                        sys.exit(1)
                    elif response.status_code == 403:
                        print("Fehler 403: Zugriff verweigert")
                        print("Der Benutzer hat keine Berechtigung für diese Ressource")
                        sys.exit(1)
                    
                    response.raise_for_status()
                    
                    # Prüfe ob die Antwort eine gültige iCal-Datei ist
                    content = response.text
                    if not content.strip().startswith('BEGIN:VCALENDAR'):
                        print("Warnung: Die heruntergeladene Datei scheint keine gültige iCal-Datei zu sein")
                        print("Erste 200 Zeichen der Antwort:")
                        print(content[:200])
                        print("\nVersuche trotzdem fortzufahren...")
                    
                    print(f"iCal-Datei erfolgreich heruntergeladen ({len(content)} Zeichen)")
                    return content
                    
                except requests.exceptions.Timeout:
                    print(f"Zeitüberschreitung beim Versuch {attempt + 1}/{max_retries}")
                    if attempt < max_retries - 1:
                        print(f"Warte {retry_delay} Sekunden vor nächstem Versuch...")
                        import time
                        time.sleep(retry_delay)
                    else:
                        print("Fehler: Zeitüberschreitung beim Herunterladen der iCal-Datei")
                        print(f"Alle {max_retries} Versuche mit {timeout}s Timeout fehlgeschlagen")
                        print("Lösungsvorschläge:")
                        print("1. Erhöhen Sie das Timeout in der Konfiguration: timeout = 120")
                        print("2. Überprüfen Sie Ihre Internetverbindung")
                        print("3. Versuchen Sie es zu einem anderen Zeitpunkt")
                        print("4. Kontaktieren Sie den Administrator des Kalenderservers")
                        sys.exit(1)
                except requests.exceptions.ConnectionError:
                    print(f"Verbindungsfehler beim Versuch {attempt + 1}/{max_retries}")
                    if attempt < max_retries - 1:
                        print(f"Warte {retry_delay} Sekunden vor nächstem Versuch...")
                        import time
                        time.sleep(retry_delay)
                    else:
                        print("Fehler: Verbindung zur URL fehlgeschlagen")
                        print("Überprüfen Sie Ihre Internetverbindung und die URL")
                        sys.exit(1)
                        
        except Exception as e:
            print(f"Fehler beim Herunterladen der iCal-Datei: {e}")
            print(f"URL: {url}")
            print(f"Benutzername: {username}")
            sys.exit(1)
    
    def parse_ical(self, ical_content):
        """Parst die iCal-Datei und filtert relevante Einträge"""
        calendar = Calendar.from_ical(ical_content)
        events = []
        
        for component in calendar.walk():
            if component.name == "VEVENT":
                # Prüfe ob der Eintrag im gewünschten Monat liegt
                start_date = component.get('dtstart').dt
                
                # Konvertiere zu lokaler Zeit falls nötig
                if hasattr(start_date, 'tzinfo') and start_date.tzinfo is not None:
                    start_date = start_date.astimezone(self.local_tz)
                elif not hasattr(start_date, 'hour'):  # Falls es nur ein Datum ist
                    start_date = datetime.combine(start_date, datetime.min.time())
                    start_date = start_date.replace(tzinfo=self.local_tz)
                
                if (start_date.month == self.month and 
                    start_date.year == self.year):
                    
                    description = str(component.get('description', ''))
                    summary = str(component.get('summary', ''))
                    
                    # Prüfe auf Kundencode am Anfang der Beschreibung
                    # PvSA, debug summary in month
                    #print(summary)
                    match = re.match(r'^([A-Z]{3}):', summary)
                    if match:
                        customer_code = match.group(1)
                        # Entferne Kundencode aus der Beschreibung
                        clean_description = summary[4:].strip()
                        
                        # Berechne Dauer falls vorhanden
                        duration = None
                        if 'dtend' in component:
                            end_date = component.get('dtend').dt
                            if hasattr(end_date, 'tzinfo') and end_date.tzinfo is not None:
                                end_date = end_date.astimezone(self.local_tz)
                            elif not hasattr(end_date, 'hour'):
                                end_date = datetime.combine(end_date, datetime.min.time())
                                end_date = end_date.replace(tzinfo=self.local_tz)
                            
                            if hasattr(start_date, 'hour') and hasattr(end_date, 'hour'):
                                duration = end_date - start_date
                        
                        events.append({
                            'customer_code': customer_code,
                            'date': start_date,
                            'description': clean_description,
                            'duration': duration
                        })
        
        return events
    
    def group_by_customer(self, events):
        """Gruppiert Events nach Kundencode"""
        customer_events = defaultdict(list)
        for event in events:
            customer_events[event['customer_code']].append(event)
        return dict(customer_events)
    
    def create_pdf(self, customer_code, events):
        """Erstellt ein PDF für einen Kunden"""
        month_names = {
            1: 'Januar', 2: 'Februar', 3: 'März', 4: 'April',
            5: 'Mai', 6: 'Juni', 7: 'Juli', 8: 'August',
            9: 'September', 10: 'Oktober', 11: 'November', 12: 'Dezember'
        }
        
        month_name = month_names[self.month]
        filename = f"{customer_code}_Abrechnung_{month_name}_{self.year}.pdf"
        
        # Erstelle PDF
        doc = SimpleDocTemplate(filename, pagesize=A4)
        story = []
        styles = getSampleStyleSheet()
        
        # Logo einfügen (falls vorhanden)
        logo_path = self.config.get('pdf', 'logo_path', fallback='logo.jpg')
        if os.path.exists(logo_path):
            try:
                logo = Image(logo_path, width=4*cm, height=2*cm)
                story.append(logo)
                story.append(Spacer(1, 12))
            except Exception as e:
                print(f"Warnung: Logo konnte nicht geladen werden: {e}")
        
        # Titel
        title_style = styles['Title']
        title = Paragraph(f"Anlage Arbeitsstunden {customer_code} - {month_name} {self.year}", title_style)
        story.append(title)
        story.append(Spacer(1, 20))
        
        # Kundendaten (falls in Config vorhanden)
        #if self.config.has_section('invoice_header'):
        #header_style = styles['Normal']
        #    for key, value in self.config['invoice_header'].items():
        #story.append(Paragraph(f"<b>{key.title()}:</b> {value}", header_style))
        #story.append(Spacer(1, 20))
        
        # Tabelle mit Terminen
        table_data = [['Datum', 'Uhrzeit', 'Beschreibung', 'Dauer']]
        
        total_duration = timedelta()
        for event in sorted(events, key=lambda x: x['date']):
            date_str = event['date'].strftime('%d.%m.%Y')
            time_str = event['date'].strftime('%H:%M')
            description = event['description'] if event['description'] else event['summary']
            
            duration_str = ''
            if event['duration']:
                hours, remainder = divmod(event['duration'].total_seconds(), 3600)
                minutes, _ = divmod(remainder, 60)
#                duration_str = f"{int(hours):02d}:{int(minutes):02d}"
                decimal_hours = hours + minutes / 60
                duration_str = f"{decimal_hours:.2f}".replace('.', ',')
                
                total_duration += event['duration']
            
            table_data.append([date_str, time_str, description, duration_str])
        
        # Gesamtdauer hinzufügen
        if total_duration.total_seconds() > 0:
            hours, remainder = divmod(total_duration.total_seconds(), 3600)
            minutes, _ = divmod(remainder, 60)
            #table_data.append(['', '', 'Gesamtdauer:', f"{int(hours):02d}:{int(minutes):02d}"])
            decimal_total_hours = hours + minutes / 60
            table_data.append(['', '', 'Gesamtdauer:', f"{decimal_total_hours:.2f}".replace('.', ',')])

        
        # Tabelle erstellen
        table = Table(table_data, colWidths=[3*cm, 2*cm, 8*cm, 2*cm])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -2), colors.white),
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        story.append(table)
        story.append(Spacer(1, 20))
        
        # Fußzeile
        footer_style = styles['Normal']
        story.append(Paragraph(f"Erstellt am: {datetime.now().strftime('%d.%m.%Y %H:%M')}", footer_style))
        
        # PDF erstellen
        doc.build(story)
        print(f"PDF erstellt: {filename}")
        
        return filename
    
    def send_email(self, pdf_filename, customer_code):
        """Sendet das PDF per E-Mail"""
        try:
            # SMTP Konfiguration
            smtp_server = self.config['smtp']['server']
            smtp_port = int(self.config['smtp']['port'])
            sender_email = self.config['smtp']['sender_email']
            recipient_email = self.config['smtp']['recipient_email']
            
            # Optionale SMTP-Authentifizierung und Verschlüsselung
            use_starttls = self.config.getboolean('smtp', 'use_starttls', fallback=False)
            use_ssl = self.config.getboolean('smtp', 'use_ssl', fallback=False)
            smtp_username = self.config.get('smtp', 'smtp_username', fallback=None)
            smtp_password = self.config.get('smtp', 'smtp_password', fallback=None)
            
            print(f"Sende E-Mail über {smtp_server}:{smtp_port}")
            print(f"STARTTLS: {use_starttls}, SSL: {use_ssl}, Auth: {bool(smtp_username)}")
            
            # E-Mail erstellen
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Subject'] = f"Abrechnung {customer_code} - {self.month:02d}/{self.year}"
            
            body = f"""
            Hallo,
            
            anbei finden Sie die Abrechnung für {customer_code} vom {self.month:02d}/{self.year}.
            
            Mit freundlichen Grüßen
            """
            
            msg.attach(MIMEText(body, 'plain'))
            
            # PDF anhängen
            with open(pdf_filename, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename= {pdf_filename}'
                )
                msg.attach(part)
            
            # SMTP-Verbindung aufbauen
            if use_ssl:
                # Direkte SSL-Verbindung (meist Port 465)
                server = smtplib.SMTP_SSL(smtp_server, smtp_port)
                print("SSL-Verbindung hergestellt")
            else:
                # Standard SMTP-Verbindung (meist Port 25 oder 587)
                server = smtplib.SMTP(smtp_server, smtp_port)
                print("SMTP-Verbindung hergestellt")
                
                if use_starttls:
                    # STARTTLS für sichere Übertragung aktivieren
                    server.starttls()
                    print("STARTTLS aktiviert")
            
            # Authentifizierung falls erforderlich
            if smtp_username and smtp_password:
                server.login(smtp_username, smtp_password)
                print("SMTP-Authentifizierung erfolgreich")
            
            # E-Mail senden
            server.sendmail(sender_email, recipient_email, msg.as_string())
            server.quit()
            
            print(f"E-Mail für {customer_code} erfolgreich gesendet")
            
        except smtplib.SMTPAuthenticationError as e:
            print(f"Fehler bei der SMTP-Authentifizierung für {customer_code}: {e}")
            print("Überprüfen Sie SMTP-Benutzername und -Passwort")
        except smtplib.SMTPConnectError as e:
            print(f"Fehler beim Verbinden zum SMTP-Server für {customer_code}: {e}")
            print("Überprüfen Sie SMTP-Server und Port")
        except smtplib.SMTPException as e:
            print(f"SMTP-Fehler beim Senden der E-Mail für {customer_code}: {e}")
        except Exception as e:
            print(f"Allgemeiner Fehler beim Senden der E-Mail für {customer_code}: {e}")
    
    def generate_invoices(self):
        """Hauptfunktion zur Erstellung aller Rechnungen"""
        print(f"Lade iCal-Datei für {self.username}...")
        ical_content = self.download_ical()
        
        print("Parse iCal-Datei...")
        events = self.parse_ical(ical_content)
        
        if not events:
            print(f"Keine relevanten Termine für {self.month:02d}/{self.year} gefunden.")
            return
        
        print(f"Gefundene Termine: {len(events)}")
        
        # Gruppiere nach Kunden
        customer_events = self.group_by_customer(events)
        
        print(f"Kunden gefunden: {list(customer_events.keys())}")
        
        # Erstelle PDF für jeden Kunden
        for customer_code, events in customer_events.items():
            print(f"Erstelle Rechnung für {customer_code}...")
            pdf_filename = self.create_pdf(customer_code, events)
            
            # Sende E-Mail
            if self.config.has_section('smtp'):
                self.send_email(pdf_filename, customer_code)


def main():
    parser = argparse.ArgumentParser(description='Erstellt Rechnungen aus iCal-Dateien')
    parser.add_argument('username', help='Benutzername für Konfigurationsdatei')
    parser.add_argument('month', type=int, help='Monat (1-12)')
    parser.add_argument('year', type=int, help='Jahr (z.B. 2024)')
    
    args = parser.parse_args()
    
    # Validiere Eingaben
    if not 1 <= args.month <= 12:
        print("Fehler: Monat muss zwischen 1 und 12 liegen")
        sys.exit(1)
    
    if args.year < 2000 or args.year > 2100:
        print("Fehler: Jahr muss zwischen 2000 und 2100 liegen")
        sys.exit(1)
    
    # Erstelle Pfad zur Konfigurationsdatei
    base_path = os.path.dirname(os.path.abspath(__file__))
    config_file = os.path.join(base_path, 'conf', f'{args.username}.conf')
    
    if not os.path.exists(config_file):
        print(f"Fehler: Konfigurationsdatei {config_file} nicht gefunden")
        print("\nBeispiel-Konfigurationsdatei:")
        print("""
[ical]
url = https://calendar.example.com/calendar.ics
username = ihr_username
password = ihr_passwort
auth_method = basic
timeout = 60
max_retries = 3
retry_delay = 5
# Mögliche auth_method Werte:
# - basic: HTTP Basic Authentication (Standard)
# - digest: HTTP Digest Authentication
# - token: Bearer Token (password wird als Token verwendet)
# - url_params: Parameter in der URL
# 
# timeout: Sekunden bis zur Zeitüberschreitung (Standard: 60)
# max_retries: Maximale Anzahl der Wiederholungsversuche (Standard: 3)
# retry_delay: Wartezeit zwischen Versuchen in Sekunden (Standard: 5)

[smtp]
server = smtp.example.com
port = 587
sender_email = sender@example.com
recipient_email = recipient@example.com
use_starttls = true
use_ssl = false
smtp_username = ihr_smtp_username
smtp_password = ihr_smtp_passwort
# SMTP-Verschlüsselungsoptionen:
# - use_starttls: STARTTLS verwenden (empfohlen für Port 587)
# - use_ssl: Direkte SSL-Verbindung (meist Port 465)
# - smtp_username/smtp_password: Für SMTP-Authentifizierung (optional)

[pdf]
logo_path = logo.jpg

[invoice_header]
firma = Ihre Firma GmbH
adresse = Musterstraße 1, 12345 Musterstadt
telefon = +49 123 456789
email = info@ihre-firma.de
""")
        sys.exit(1)
    
    # Erstelle und starte Generator
    generator = InvoiceGenerator(config_file, args.username, args.month, args.year)
    generator.generate_invoices()


if __name__ == "__main__":
    main()
