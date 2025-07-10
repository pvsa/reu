import icalendar
import sys
import os
import requests
import smtplib
from datetime import datetime, timedelta, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader

def ensure_directory_exists(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

def read_user_data(username):
    file_path = os.path.join('conf', f"{username}.conf")
    try:
        with open(file_path, 'r') as f:
            line = f.readline().strip()
            url, password, email = line.split(',')
            return url, password, email
    except FileNotFoundError:
        print(f"User data file not found: {file_path}")
        return None, None, None
    except Exception as e:
        print(f"An error occurred: {e}")
        return None, None, None

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

            if isinstance(event_start, datetime) and event_start.tzinfo is not None:
                event_start = event_start.replace(tzinfo=None)
            if isinstance(event_end, datetime) and event_end.tzinfo is not None:
                event_end = event_end.replace(tzinfo=None)

            if start_date <= event_start <= end_date or start_date <= event_end <= end_date:
                filtered_events.append(component)
    return filtered_events

def generate_invoice(customer_code, events, username, month, year):
    ensure_directory_exists('archive')
    file_name = os.path.join('archive', f"{customer_code}_Abrechnung_{month}_{year}.pdf")
    c = canvas.Canvas(file_name, pagesize=letter)
    width, height = letter

    # Logo einfügen
    logo_path = "logo.png"
    if os.path.exists(logo_path):
        logo = ImageReader(logo_path)
        c.drawImage(logo, 50, height - 100, width=100, height=50, preserveAspectRatio=True)

    # Titel setzen
    title = f"{customer_code} Abrechnung {month}/{year}"
    c.setFont("Helvetica-Bold", 16)
    c.drawString(200, height - 60, title)

    # Überschrift im PDF
    c.setFont("Helvetica-Bold", 14)
    c.drawString(100, height - 100, title)

    # Weitere Informationen
    c.setFont("Helvetica", 12)
    c.drawString(100, height - 130, f"Generated for User: {username}")
    c.drawString(100, height - 150, "Events:")

    y_position = height - 170
    for event in events:
        summary = event.get('summary')
        start = event.get('dtstart').dt
        end = event.get('dtend').dt
        c.drawString(120, y_position, f"{start} to {end}: {summary}")
        y_position -= 20

    c.save()
    return file_name

def send_email(sender_email, recipient_email, subject, body, attachment_path):
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
        server = smtplib.SMTP('mail.pilarkto.net', 587)
        server.starttls()
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)
        server.quit()
        print(f"Email sent to {recipient_email} with attachment {attachment_path}")
    except Exception as e:
        print(f"Failed to send email: {e}")

def main(username, month, year):
    url, password, email = read_user_data(username)
    if not url or not password or not email:
        print("User data not found.")
        return

    ical_file = download_icalendar_file(url, username, password)
    cal = parse_icalendar_file(ical_file)
    filtered_events = filter_events_by_month_and_year(cal, month, year)

    customer_events = {}
    for event in filtered_events:
        description = str(event.get('description'))
        if ':' in description:
            customer_code = description.split(':')[0].strip()
            if customer_code.isupper() and len(customer_code) == 3:
                if customer_code not in customer_events:
                    customer_events[customer_code] = []
                customer_events[customer_code].append(event)

    sender_email = "your_email@example.com"  # Ersetzen Sie durch Ihre E-Mail-Adresse

    for customer_code, events in customer_events.items():
        pdf_file = generate_invoice(customer_code, events, username, month, year)
        subject = f"Invoice for {customer_code} - {month}/{year}"
        body = f"Please find attached the invoice for {customer_code} for the period {month}/{year}."
        send_email(sender_email, email, subject, body, pdf_file)

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python generate_invoices.py <username> <month> <year>")
        sys.exit(1)

    username = sys.argv[1]
    month = int(sys.argv[2])
    year = int(sys.argv[3])

    main(username, month, year)
