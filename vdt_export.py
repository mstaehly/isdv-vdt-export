#!/usr/bin/env python3
"""
Monatlicher VDT-Mitglieder-Export aus Campai
Filtert Mitglieder mit Tag "VDT" und versendet eine passwortgeschuetzte ZIP per E-Mail.
"""

import os
import json
import smtplib
import tempfile
import requests
from datetime import datetime, timezone
from dateutil.relativedelta import relativedelta
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import pyzipper

CAMPAI_API_KEY    = os.environ["CAMPAI_API_KEY"]
CAMPAI_ORG_ID     = "692f09995ca5c71a40524dc0"
CAMPAI_MANDATE_ID = "mzfbs"

SMTP_HOST         = os.environ.get("SMTP_HOST", "smtp.strato.de")
SMTP_PORT         = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER         = os.environ["SMTP_USER"]
SMTP_PASSWORD     = os.environ["SMTP_PASSWORD"]

SENDER_EMAIL      = "m.staehly@isdv.net"
RECIPIENT_EMAIL   = "info@ms-sounddesign.com"
ZIP_PASSWORD      = os.environ["ZIP_PASSWORD"]

TAG_VDT           = "VDT"


def get_date_range():
    today = datetime.now(timezone.utc)
    first_of_this_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    first_of_last_month = first_of_this_month - relativedelta(months=1)
    return first_of_last_month, first_of_this_month


def fetch_vdt_members(date_from, date_to):
    url = f"https://cloud.campai.com/api/{CAMPAI_ORG_ID}/{CAMPAI_MANDATE_ID}/crm/contacts/export"

    payload = {
        "format": "xlsx",
        "tags": [TAG_VDT],
        "types": ["member"],
        "organizationId": "692f09995ca5c71a40524dc0",
        "mandateId": "mzfbs",
        "formatOptions": {
            "xlsx": {
                "arrayFormat": "join"
            }
        }
    }

    headers = {
        "X-API-Key": CAMPAI_API_KEY,
        "Content-Type": "application/json"
    }

    print(f"Request an: {url}")
    print(f"Payload: {json.dumps(payload)}")

    response = requests.post(url, headers=headers, json=payload, timeout=60)

    if not response.ok:
        print(f"Fehler {response.status_code}: {response.text[:2000]}")
        response.raise_for_status()

    content_type = response.headers.get("Content-Type", "")
    print(f"Content-Type: {content_type}")

    if "application/json" in content_type:
        data = response.json()
        print(f"JSON: {json.dumps(data)[:500]}")
        if isinstance(data.get("data"), dict):
            file_url = data["data"].get("url") or data["data"].get("downloadUrl")
            if file_url:
                r2 = requests.get(file_url, timeout=60)
                r2.raise_for_status()
                return r2.content
        raise ValueError(f"Unerwartete Antwort: {json.dumps(data)[:500]}")

    return response.content


def create_encrypted_zip(excel_bytes, filename, password):
    with tempfile.NamedTemporaryFile(suffix=".zip", delete=False) as tmp:
        zip_path = tmp.name

    with pyzipper.AESZipFile(zip_path, "w",
                             compression=pyzipper.ZIP_DEFLATED,
                             encryption=pyzipper.WZ_AES) as zf:
        zf.setpassword(password.encode("utf-8"))
        zf.writestr(filename, excel_bytes)

    with open(zip_path, "rb") as f:
        zip_bytes = f.read()

    os.unlink(zip_path)
    return zip_bytes


def send_email(zip_bytes, zip_filename, month_label):
    msg = MIMEMultipart()
    msg["From"]    = SENDER_EMAIL
    msg["To"]      = RECIPIENT_EMAIL
    msg["Subject"] = f"isdv e.V. - VDT-Mitglieder Neuzugaenge {month_label}"

    body = (
        f"Liebe Kolleginnen und Kollegen,\n\n"
        f"anbei die Liste der isdv-Mitglieder mit VDT-Zugehoerigkeit "
        f"fuer {month_label}.\n\n"
        f"Die Excel-Datei ist passwortgeschuetzt.\n\n"
        f"Mit freundlichen Gruessen\n"
        f"isdv e.V.\n"
    )

    msg.attach(MIMEText(body, "plain", "utf-8"))

    attachment = MIMEBase("application", "zip")
    attachment.set_payload(zip_bytes)
    encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", f'attachment; filename="{zip_filename}"')
    msg.attach(attachment)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.ehlo()
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, msg.as_string())

    print(f"E-Mail versandt an {RECIPIENT_EMAIL}")


def main():
    date_from, date_to = get_date_range()
    month_label    = date_from.strftime("%B %Y")
    excel_filename = f"VDT_Mitglieder_{date_from.strftime('%Y-%m')}.xlsx"
    zip_filename   = f"VDT_Mitglieder_{date_from.strftime('%Y-%m')}.zip"

    print(f"Exportiere VDT-Mitglieder fuer {month_label}...")
    excel_bytes = fetch_vdt_members(date_from, date_to)
    print(f"{len(excel_bytes)} Bytes empfangen")

    zip_bytes = create_encrypted_zip(excel_bytes, excel_filename, ZIP_PASSWORD)
    print(f"ZIP erstellt ({len(zip_bytes)} Bytes)")

    send_email(zip_bytes, zip_filename, month_label)


if __name__ == "__main__":
    main()
