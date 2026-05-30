#!/usr/bin/env python3
"""
Monatlicher VDT-Mitglieder-Export aus Campai
Filtert Mitglieder mit Tag "VDT", die im letzten Monat beigetreten sind,
und versendet eine passwortgeschützte Excel-Datei per E-Mail.
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
import pyzipper  # für AES-verschlüsselte ZIP mit Excel

# ── Konfiguration aus Umgebungsvariablen ────────────────────────────────────
CAMPAI_API_KEY     = os.environ["CAMPAI_API_KEY"]
CAMPAI_ORG_ID      = "mzfbs"
CAMPAI_MANDATE_ID  = "mzfbs"  # anpassen falls abweichend

SMTP_HOST          = os.environ.get("SMTP_HOST", "smtp.strato.de")
SMTP_PORT          = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER          = os.environ["SMTP_USER"]        # m.staehly@isdv.net
SMTP_PASSWORD      = os.environ["SMTP_PASSWORD"]

SENDER_EMAIL       = "m.staehly@isdv.net"
RECIPIENT_EMAIL    = "info@ms-sounddesign.com"      # TEST — produktiv: grommes@tonmeisterverband.org
ZIP_PASSWORD       = os.environ["ZIP_PASSWORD"]     # Passwort für die ZIP

TAG_VDT            = "VDT"
# ────────────────────────────────────────────────────────────────────────────


def get_date_range():
    """Ersten und letzten Tag des Vormonats berechnen."""
    today = datetime.now(timezone.utc)
    first_of_this_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    first_of_last_month = first_of_this_month - relativedelta(months=1)
    last_of_last_month  = first_of_this_month
    return first_of_last_month, last_of_last_month


def fetch_vdt_members(date_from: datetime, date_to: datetime) -> bytes:
    """
    Ruft den Campai Export-Endpunkt auf und filtert nach:
    - Tag "VDT"
    - member.entryAt im letzten Monat
    Gibt den rohen Excel-Dateiinhalt zurück.
    """
    url = (
        f"https://cloud.campai.com/api/{CAMPAI_ORG_ID}/{CAMPAI_MANDATE_ID}"
        f"/crm/contacts/export"
    )

    payload = {
        "format": "xlsx",
        "tags": [TAG_VDT],
        "types": ["member"],
        # userFilter temporär deaktiviert — erst alle VDT-Mitglieder exportieren
        # um die korrekten Feldnamen für den Datumsfilter zu ermitteln
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

    response = requests.post(url, headers=headers, json=payload, timeout=60)
    response.raise_for_status()

    # Campai gibt entweder direkt die Datei zurück oder ein JSON mit URL
    content_type = response.headers.get("Content-Type", "")
    if "application/json" in content_type:
        data = response.json()
        # Falls asynchron: Download-URL aus Response holen
        if "data" in data and isinstance(data["data"], dict):
            file_url = data["data"].get("url") or data["data"].get("downloadUrl")
            if file_url:
                file_response = requests.get(file_url, timeout=60)
                file_response.raise_for_status()
                return file_response.content
        raise ValueError(f"Unerwartete JSON-Antwort: {json.dumps(data)[:500]}")
    else:
        return response.content


def create_encrypted_zip(excel_bytes: bytes, filename: str, password: str) -> bytes:
    """Erstellt eine AES-verschlüsselte ZIP-Datei mit der Excel-Datei."""
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


def send_email(zip_bytes: bytes, zip_filename: str, month_label: str):
    """Versendet die verschlüsselte ZIP per E-Mail."""
    msg = MIMEMultipart()
    msg["From"]    = SENDER_EMAIL
    msg["To"]      = RECIPIENT_EMAIL
    msg["Subject"] = f"isdv e.V. - VDT-Mitglieder Neuzugänge {month_label}"

    body = f"""Liebe Kolleginnen und Kollegen,

anbei die Liste der isdv-Mitglieder mit VDT-Zugehörigkeit, die im {month_label} neu beigetreten sind.

Die Excel-Datei ist passwortgeschützt. Das Passwort erhalten Sie auf dem üblichen Weg.

Mit freundlichen Grüßen
isdv e.V. – Interessengemeinschaft der selbständigen DienstleisterInnen in der Veranstaltungswirtschaft e.V.
"""

    msg.attach(MIMEText(body, "plain", "utf-8"))

    attachment = MIMEBase("application", "zip")
    attachment.set_payload(zip_bytes)
    encoders.encode_base64(attachment)
    attachment.add_header(
        "Content-Disposition",
        f'attachment; filename="{zip_filename}"'
    )
    msg.attach(attachment)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.ehlo()
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, msg.as_string())

    print(f"✓ E-Mail erfolgreich versandt an {RECIPIENT_EMAIL}")


def main():
    date_from, date_to = get_date_range()
    month_label = date_from.strftime("%B %Y")  # z.B. "April 2026"
    excel_filename = f"VDT_Mitglieder_{date_from.strftime('%Y-%m')}.xlsx"
    zip_filename   = f"VDT_Mitglieder_{date_from.strftime('%Y-%m')}.zip"

    print(f"Exportiere VDT-Mitglieder für {month_label}...")
    excel_bytes = fetch_vdt_members(date_from, date_to)
    print(f"✓ {len(excel_bytes)} Bytes empfangen")

    print("Erstelle verschlüsselte ZIP...")
    zip_bytes = create_encrypted_zip(excel_bytes, excel_filename, ZIP_PASSWORD)
    print(f"✓ ZIP erstellt ({len(zip_bytes)} Bytes)")

    print("Versende E-Mail...")
    send_email(zip_bytes, zip_filename, month_label)


if __name__ == "__main__":
    main()
