#!/usr/bin/env python3
"""
Monatlicher VDT-Mitglieder-Export aus Campai
Sendet E-Mail via Microsoft Graph API (OAuth2).
"""

import os
import json
import base64
import tempfile
import requests
import openpyxl
from datetime import datetime, timezone
from dateutil.relativedelta import relativedelta
import pyzipper

CAMPAI_API_KEY    = os.environ["CAMPAI_API_KEY"]
CAMPAI_ORG_ID     = "692f09955ca5c71a40524db3"

MS_CLIENT_ID      = os.environ["MS_CLIENT_ID"]
MS_TENANT_ID      = os.environ["MS_TENANT_ID"]
MS_CLIENT_SECRET  = os.environ["MS_CLIENT_SECRET"]

SENDER_EMAIL      = "m.staehly@isdv.net"
RECIPIENT_EMAIL   = "info@ms-sounddesign.com"
ZIP_PASSWORD      = os.environ["ZIP_PASSWORD"]

TAG_VDT           = "VDT"


def get_date_range():
    today = datetime.now(timezone.utc)
    first_of_this_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    first_of_last_month = first_of_this_month - relativedelta(months=1)
    return first_of_last_month, first_of_this_month


def get_ms_token():
    """OAuth2 Token von Microsoft holen."""
    url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": MS_CLIENT_ID,
        "client_secret": MS_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(url, data=data, timeout=30)
    if not response.ok:
        print(f"Token-Fehler: {response.text[:500]}")
        response.raise_for_status()
    return response.json()["access_token"]


def fetch_vdt_contacts():
    """Alle Kontakte mit Tag VDT aus der alten Campai API abrufen."""
    url = "https://api.campai.com/contacts"
    headers = {
        "Authorization": CAMPAI_API_KEY,
        "Content-Type": "application/json"
    }

    all_contacts = []
    skip = 0
    limit = 100

    while True:
        params = {
            "organisation": CAMPAI_ORG_ID,
            "tags": TAG_VDT,
            "type": "member",
            "limit": limit,
            "skip": skip
        }

        print(f"Abruf Skip {skip}...")
        response = requests.get(url, headers=headers, params=params, timeout=60)

        if not response.ok:
            print(f"Fehler {response.status_code}: {response.text[:2000]}")
            response.raise_for_status()

        data = response.json()

        if isinstance(data, list):
            contacts = data
        elif isinstance(data, dict):
            contacts = data.get("data") or data.get("items") or data.get("contacts") or []
        else:
            contacts = []

        all_contacts.extend(contacts)
        print(f"{len(all_contacts)} Kontakte bisher")

        if len(contacts) < limit:
            break
        skip += limit

    return all_contacts


def filter_new_members(contacts, date_from, date_to):
    """Nur Mitglieder filtern, die im Vormonat beigetreten sind."""
    new_members = []
    for c in contacts:
        if not isinstance(c, dict):
            continue
        merge_tags = c.get("mergeTags") or {}
        entry_at = (
            merge_tags.get("enterDate") or
            c.get("membership", {}).get("enterDate") or
            c.get("createdAt")
        )

        if entry_at:
            try:
                # Format DD/MM/YYYY
                if "/" in str(entry_at):
                    parts = entry_at.split("/")
                    if len(parts) == 3:
                        entry_date = datetime(
                            int(parts[2]), int(parts[1]), int(parts[0]),
                            tzinfo=timezone.utc
                        )
                else:
                    entry_date = datetime.fromisoformat(entry_at.replace("Z", "+00:00"))

                if date_from <= entry_date < date_to:
                    new_members.append(c)
            except Exception as e:
                print(f"Datum-Parse-Fehler: {entry_at} -> {e}")

    print(f"{len(new_members)} neue VDT-Mitglieder im Zeitraum gefunden")
    return new_members


def build_excel(contacts):
    """Excel-Datei aus Kontaktdaten bauen."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "VDT-Mitglieder"
    ws.append(["Vorname", "Nachname", "E-Mail", "Eintrittsdatum"])

    for c in contacts:
        if not isinstance(c, dict):
            continue

        merge_tags = c.get("mergeTags") or {}
        vorname = merge_tags.get("personFirstName") or ""
        nachname = merge_tags.get("personLastName") or ""

        if not vorname and not nachname:
            name = merge_tags.get("name") or c.get("name", "")
            parts = name.split(" ", 1)
            vorname = parts[0] if parts else ""
            nachname = parts[1] if len(parts) > 1 else ""

        email = merge_tags.get("email") or ""

        entry_at = merge_tags.get("enterDate") or ""
        if entry_at and "/" in entry_at:
            try:
                parts = entry_at.split("/")
                entry_at = f"{parts[0]}.{parts[1]}.{parts[2]}"
            except Exception:
                pass

        ws.append([vorname, nachname, email, entry_at])

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_len + 4

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


def send_email_graph(zip_bytes, zip_filename, month_label, count, token):
    """E-Mail via Microsoft Graph API senden."""
    zip_b64 = base64.b64encode(zip_bytes).decode("utf-8")

    message = {
        "message": {
            "subject": f"isdv e.V. - VDT-Mitglieder Neuzugaenge {month_label}",
            "body": {
                "contentType": "Text",
                "content": (
                    f"Liebe Kolleginnen und Kollegen,\n\n"
                    f"anbei die Liste der {count} isdv-Mitglieder mit VDT-Zugehoerigkeit, "
                    f"die im {month_label} neu beigetreten sind.\n\n"
                    f"Die Excel-Datei ist passwortgeschuetzt.\n\n"
                    f"Mit freundlichen Gruessen\n"
                    f"isdv e.V.\n"
                )
            },
            "toRecipients": [
                {"emailAddress": {"address": RECIPIENT_EMAIL}}
            ],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": zip_filename,
                    "contentBytes": zip_b64
                }
            ]
        }
    }

    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    response = requests.post(url, headers=headers, json=message, timeout=60)

    if not response.ok:
        print(f"Graph-Fehler {response.status_code}: {response.text[:1000]}")
        response.raise_for_status()

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

    excel_bytes = build_excel(new_members)
    print(f"Excel erstellt ({len(excel_bytes)} Bytes)")

    zip_bytes = create_encrypted_zip(excel_bytes, excel_filename, ZIP_PASSWORD)
    print(f"ZIP erstellt ({len(zip_bytes)} Bytes)")

    print("Hole Microsoft Graph Token...")
    token = get_ms_token()
    print("Token erhalten, sende E-Mail...")

    send_email_graph(zip_bytes, zip_filename, month_label, len(new_members), token)


if __name__ == "__main__":
    main()
