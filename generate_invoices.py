import icalendar
import sys
import os
import re
import requests
import smtplib
import pytz
from datetime import datetime, timedelta, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

def ensure_directory_exists(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def read_default_config():
    file_path = os.path.join('conf', 'defaults.conf')
    try:
        with open(file_path, 'r') as f:
            line = f.readline().strip()
            sender_email = line
            return sender_email
    except FileNotFoundError:
        print(f"Default config file not found: {file_path}")
        return None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

def read_user_data(username):
    file_path = os.path.join('conf', f"{username}.conf")
    try:
        with open(file_path, 'r') as f:
            line = f.readline().strip()
            url, password, email, smtp_server, smtp_port = line.split(',')
            return url, password, email, smtp_server, smtp_port
    except FileNotFoundError:
        print(f"User data file not found: {file_path}")
        return None, None, None, None, None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None, None, None, None, None

def download_icalendar_file(url, username, password):
    response = requests.get(url, auth=(username, password))
    if response.status_code == 200:
        ensure_directory_exists('archive')
        file_path = os.path.join('archive', 'downloaded_calendar.ics')
        with open(file_path, 'wb') as f:
            f.write(response.content)
        return file_path
    else:
        raise Exception("Failed to download the iCalendar file")

def parse_icalendar_file(file_path):
    with open(file_path, 'rb') as f:
        cal = icalendar.Calendar.from_ical(f.read())
    return cal

def ensure_datetime(d):
    if isinstance(d, date) and not isinstance(d, datetime):
        return datetime.combine(d, datetime.min.time())
    return d

def convert_to_local_time(dt):
    if dt.tzinfo is None:
        return dt
    return dt.astimezone(pytz.timezone('Europe/Vienna')).replace(tzinfo=None)

def filter_events_by_month_and_year(cal, month, year):
    filtered_events = []
    start_date = datetime(year, month, 1)
    if month == 12:
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = datetime(year, month + 1, 1) - timedelta(days=1)

    for component in cal.walk():
        if component.name == 'VEVENT':
            event_start = ensure_datetime(component.get('dtstart').dt)
            event_end = ensure_datetime(component.get('dtend').dt)

            event_start = convert_to_local_time(event_start)
            event_end = convert_to_local_time(event_end)

            if start_date <= event_start <= end_date or start_date <= event_end <= end_date:
                filtered_events.append(component)
    return filtered_events

def generate_invoice(customer_code, events, username, month, year):
    ensure_directory_exists('archive')
    file_name = os.path.join('archive', f"{customer_code}_Abrechnung_{month}_{year}.pdf")
    doc = SimpleDocTemplate(file_name, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    # Logo einfügen
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        logo = ImageReader(logo_path)
        story.append(logo)

    # Titel setzen
    title = f"{customer_code} Abrechnung {month}/{year}"
    story.append(Paragraph(title, styles['Title']))
    story.append(Spacer(1, 12))

    # Tabelle für Events
    data = [["Beschreibung", "Datum", "Von", "Bis", "Dauer"]]
    for event in events:
        summary = str(event.get('summary'))
        start = convert_to_local_time(ensure_datetime(event.get('dtstart').dt))
        end = convert_to_local_time(ensure_datetime(event.get('dtend').dt))
        duration = end - start
        data.append([summary, start.date(), start.time(), end.time(), str(duration)])

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 14),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))

    story.append(table)
    doc.build(story)
    print(f"Invoice generated: {file_name}")
    return file_name

def send_email(sender_email, recipient_email, subject, body, attachment_path, smtp_server, smtp_port):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    with open(attachment_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {os.path.basename(attachment_path)}",
    )

    msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)
        server.quit()
        print(f"Email sent to {recipient_email} with attachment {attachment_path}")
    except Exception as e:
        print(f"Failed to send email: {e}")

def main(username, month, year):
    sender_email = read_default_config()
    if not sender_email:
        print("Default sender email not found.")
        return

    url, password, email, smtp_server, smtp_port = read_user_data(username)
    if not url or not password or not email or not smtp_server or not smtp_port:
        print("User data not found.")
        return

    ical_file = download_icalendar_file(url, username, password)
    cal = parse_icalendar_file(ical_file)
    filtered_events = filter_events_by_month_and_year(cal, month, year)

    customer_events = {}
    for event in filtered_events:
        description = str(event.get('description'))
        match = re.match(r'^[A-Z]{3}:', description)
        if match:
            customer_code = match.group()[:-1]
            if customer_code not in customer_events:
                customer_events[customer_code] = []
            customer_events[customer_code].append(event)

    for customer_code, events in customer_events.items():
        pdf_file = generate_invoice(customer_code, events, username, month, year)
        subject = f"Abrechnung für {customer_code} - {month}/{year}"
        body = f"Anbei finden Sie die Abrechnung für {customer_code} für den Zeitraum {month}/{year}."
        send_email(sender_email, email, subject, body, pdf_file, smtp_server, smtp_port)

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python generate_invoices.py <username> <month> <year>")
        sys.exit(1)

    username = sys.argv[1]
    month = int(sys.argv[2])
    year = int(sys.argv[3])

    main(username, month, year)
