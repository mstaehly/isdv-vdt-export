#!/usr/bin/env python3
"""
Monatlicher VDT-Mitglieder-Export aus Campai
Nutzt /crm/contacts/list, baut selbst eine Excel und versendet sie verschluesselt.
"""

import os
import json
import smtplib
import tempfile
import requests
import openpyxl
from datetime import datetime, timezone
from dateutil.relativedelta import relativedelta
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
import pyzipper

CAMPAI_API_KEY    = os.environ["CAMPAI_API_KEY"]
CAMPAI_ORG_ID     = "692f09995ca5c71a40524db3"
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


def fetch_vdt_contacts():
    """Alle Kontakte mit Tag VDT aus Campai abrufen (paginiert)."""
    url = f"https://cloud.campai.com/api/{CAMPAI_ORG_ID}/{CAMPAI_MANDATE_ID}/crm/contacts/list"
    headers = {
        "X-API-Key": CAMPAI_API_KEY,
        "Content-Type": "application/json"
    }

    all_contacts = []
    offset = 0
    limit = 100

    while True:
        payload = {
            "types": ["member"],
            "tags": [TAG_VDT],
            "limit": limit,
            "offset": offset,
            "returnCount": True
        }

        print(f"Abruf Offset {offset}...")
        response = requests.post(url, headers=headers, json=payload, timeout=60)

        if not response.ok:
            print(f"Fehler {response.status_code}: {response.text[:2000]}")
            response.raise_for_status()

        data = response.json()
        contacts = data.get("contacts", [])
        total = data.get("count", 0)
        all_contacts.extend(contacts)

        print(f"{len(all_contacts)} / {total} Kontakte abgerufen")

        if len(all_contacts) >= total or len(contacts) == 0:
            break
        offset += limit

    return all_contacts


def filter_new_members(contacts, date_from, date_to):
    """Nur Mitglieder filtern, die im Vormonat beigetreten sind."""
    new_members = []
    for c in contacts:
        member = c.get("member") or {}
        entry_at = member.get("entryAt") if member else None

        if not entry_at:
            # Fallback: createdAt aus dem Kontakt
            entry_at = c.get("createdAt")

        if entry_at:
            try:
                entry_date = datetime.fromisoformat(entry_at.replace("Z", "+00:00"))
                if date_from <= entry_date < date_to:
                    new_members.append(c)
            except Exception:
                pass

    print(f"{len(new_members)} neue VDT-Mitglieder im Zeitraum gefunden")
    return new_members


def build_excel(contacts, date_from):
    """Excel-Datei aus Kontaktdaten bauen."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VDT-Mitglieder"

    # Header
    ws.append(["Vorname", "Nachname", "E-Mail", "Eintrittsdatum"])

    for c in contacts:
        person = c.get("person") or {}
        vorname = person.get("firstName", "")
        nachname = person.get("lastName", "")

        # Name als Fallback
        if not vorname and not nachname:
            name = c.get("name", "")
            parts = name.split(" ", 1)
            vorname = parts[0] if parts else ""
            nachname = parts[1] if len(parts) > 1 else ""

        communication = c.get("communication") or {}
        email = communication.get("email", "")

        member = c.get("member") or {}
        entry_at = member.get("entryAt", "") if member else ""
        if entry_at:
            try:
                entry_at = datetime.fromisoformat(entry_at.replace("Z", "+00:00")).strftime("%d.%m.%Y")
            except Exception:
                pass

        ws.append([vorname, nachname, email, entry_at])

    # Spaltenbreite anpassen
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

    # In Bytes speichern
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        xlsx_path = tmp.name
    wb.save(xlsx_path)
    with open(xlsx_path, "rb") as f:
        xlsx_bytes = f.read()
    os.unlink(xlsx_path)
    return xlsx_bytes


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


def send_email(zip_bytes, zip_filename, month_label, count):
    msg = MIMEMultipart()
    msg["From"]    = SENDER_EMAIL
    msg["To"]      = RECIPIENT_EMAIL
    msg["Subject"] = f"isdv e.V. - VDT-Mitglieder Neuzugaenge {month_label}"

    body = (
        f"Liebe Kolleginnen und Kollegen,\n\n"
        f"anbei die Liste der {count} isdv-Mitglieder mit VDT-Zugehoerigkeit, "
        f"die im {month_label} neu beigetreten sind.\n\n"
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

    all_contacts = fetch_vdt_contacts()
    new_members = filter_new_members(all_contacts, date_from, date_to)

    if not new_members:
        print("Keine neuen VDT-Mitglieder im Vormonat — keine E-Mail wird versandt.")
        return

    excel_bytes = build_excel(new_members, date_from)
    print(f"Excel erstellt ({len(excel_bytes)} Bytes)")

    zip_bytes = create_encrypted_zip(excel_bytes, excel_filename, ZIP_PASSWORD)
    print(f"ZIP erstellt ({len(zip_bytes)} Bytes)")

    send_email(zip_bytes, zip_filename, month_label, len(new_members))


if __name__ == "__main__":
    main()
