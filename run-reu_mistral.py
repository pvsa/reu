import icalendar
import requests
import configparser
import os
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import utils
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def download_ical_file(url, username, password):
    response = requests.get(url, auth=(username, password))
    if response.status_code == 200:
        return response.content
    else:
        raise Exception("Failed to download the iCalendar file")

def parse_ical_data(ical_content):
    cal = icalendar.Calendar.from_ical(ical_content)
    return cal.walk('VEVENT')

def filter_events_by_month_year(events, month, year):
    filtered_events = []
    for event in events:
        dt = event.get('dtstart').dt
        if dt.month == month and dt.year == year:
            filtered_events.append(event)
    return filtered_events

def read_user_config(username):
    config = configparser.ConfigParser()
    config_file = os.path.join('conf', f'{username}.conf')
    config.read(config_file)
    return config

def generate_pdf_invoice(customer_code, events, logo_path, month, year):
    filename = f"{customer_code}_Abrechnung_{month}_{year}.pdf"
    doc = SimpleDocTemplate(filename, pagesize=letter)
    styles = getSampleStyleSheet()
    story = []

    # Add logo
    logo = utils.ImageReader(logo_path)
    img = Image(logo, width=2*inch, height=2*inch)
    story.append(img)

    # Add title
    title = f"{customer_code} Abrechnung {month}/{year}"
    story.append(Paragraph(title, styles['Title']))

    # Add events
    for event in events:
        summary = event.get('summary')
        description = event.get('description')
        start_time = event.get('dtstart').dt
        story.append(Paragraph(f"<b>Datum:</b> {start_time.strftime('%d.%m.%Y %H:%M')}", styles['Normal']))
        story.append(Paragraph(f"<b>Zusammenfassung:</b> {summary}", styles['Normal']))
        story.append(Paragraph(f"<b>Beschreibung:</b> {description}", styles['Normal']))
        story.append(Spacer(1, 0.2*inch))

    doc.build(story)
    return filename

def send_email(sender_email, receiver_email, smtp_server, smtp_port, pdf_path):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = 'Ihre Abrechnung'

    body = 'Anbei finden Sie Ihre Abrechnung.'
    msg.attach(MIMEText(body, 'plain'))

    attachment = open(pdf_path, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {os.path.basename(pdf_path)}')
    msg.attach(part)

    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.send_message(msg)

def main(month, year, username, password):
    config = read_user_config(username)
    url = config.get('ICAL', 'url')
    ical_username = config.get('ICAL', 'username')
    ical_password = config.get('ICAL', 'password')
    logo_path = config.get('PDF', 'logo_path')
    smtp_server = config.get('SMTP', 'server')
    smtp_port = config.getint('SMTP', 'port')

    ical_content = download_ical_file(url, ical_username, ical_password)
    events = parse_ical_data(ical_content)
    filtered_events = filter_events_by_month_year(events, month, year)

    customers = set()
    for event in filtered_events:
        description = str(event.get('description'))
        if ':' in description:
            customer_code = description.split(':')[0]
            customers.add(customer_code)

    for customer in customers:
        customer_events = [event for event in filtered_events if str(event.get('description')).startswith(f"{customer}:")]
        pdf_path = generate_pdf_invoice(customer, customer_events, logo_path, month, year)
        send_email(ical_username, "customer@example.com", smtp_server, smtp_port, pdf_path)

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--month', type=int, required=True)
    parser.add_argument('--year', type=int, required=True)
    parser.add_argument('--username', type=str, required=True)
    parser.add_argument('--password', type=str, required=True)
    args = parser.parse_args()

    main(args.month, args.year, args.username, args.password)
