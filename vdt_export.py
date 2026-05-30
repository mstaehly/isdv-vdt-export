#!/usr/bin/env python3

import os
import base64
import tempfile
import requests
from datetime import datetime, timezone
from dateutil.relativedelta import relativedelta

from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table
from reportlab.lib import pdfencrypt

import secrets
import string


# ---------------- CONFIG ----------------

CAMPAI_API_KEY = os.environ["CAMPAI_API_KEY"]
CAMPAI_ORG_ID = "692f09955ca5c71a40524db3"

MS_CLIENT_ID = os.environ["MS_CLIENT_ID"]
MS_TENANT_ID = os.environ["MS_TENANT_ID"]
MS_CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]

SENDER_EMAIL = "m.staehly@isdv.net"
RECIPIENT_EMAIL = "info@ms-sounddesign.com"

TAG_VDT = "VDT"

# ----------------------------------------


def generate_password(length=12):
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))


def get_date_range():
    today = datetime.now(timezone.utc)
    first_of_this_month = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    first_of_last_month = first_of_this_month - relativedelta(months=1)
    return first_of_last_month, first_of_this_month


def get_ms_token():
    url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": MS_CLIENT_ID,
        "client_secret": MS_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }

    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()["access_token"]


def fetch_vdt_contacts():
    url = "https://api.campai.com/contacts"
    headers = {"Authorization": CAMPAI_API_KEY}

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

        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()

        contacts = response.json()
        all_contacts.extend(contacts)

        if len(contacts) < limit:
            break

        skip += limit

    return all_contacts


def filter_new_members(contacts, date_from, date_to):
    new_members = []

    for c in contacts:
        merge_tags = c.get("mergeTags", {})
        entry = merge_tags.get("enterDate")

        if not entry:
            continue

        try:
            d, m, y = entry.split("/")
            entry_date = datetime(int(y), int(m), int(d), tzinfo=timezone.utc)

            if date_from <= entry_date < date_to:
                new_members.append(c)
        except:
            pass

    return new_members


def build_pdf_encrypted(contacts, password):
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        pdf_path = tmp.name

    data = [["Vorname", "Nachname", "E-Mail", "Eintrittsdatum"]]

    for c in contacts:
        merge = c.get("mergeTags", {})
        data.append([
            merge.get("personFirstName", ""),
            merge.get("personLastName", ""),
            merge.get("email", ""),
            merge.get("enterDate", "")
        ])

    encrypt = pdfencrypt.StandardEncryption(password)

    doc = SimpleDocTemplate(pdf_path, pagesize=A4, encrypt=encrypt)
    table = Table(data)
    doc.build([table])

    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()

    os.unlink(pdf_path)
    return pdf_bytes


def send_email(file_bytes, filename, subject, body, token):
    file_b64 = base64.b64encode(file_bytes).decode()

    message = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": body
            },
            "toRecipients": [
                {"emailAddress": {"address": RECIPIENT_EMAIL}}
            ],
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": filename,
                    "contentBytes": file_b64
                }
            ]
        }
    }

    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    response = requests.post(url, headers=headers, json=message)
    response.raise_for_status()


def send_password_email(password, month_label, token):
    message = {
        "message": {
            "subject": f"Passwort VDT Export {month_label}",
            "body": {
                "contentType": "Text",
                "content": f"Das Passwort für die PDF-Datei lautet:\n\n{password}"
            },
            "toRecipients": [
                {"emailAddress": {"address": RECIPIENT_EMAIL}}
            ]
        }
    }

    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    response = requests.post(url, headers=headers, json=message)
    response.raise_for_status()


def main():
    date_from, date_to = get_date_range()
    month_label = date_from.strftime("%B %Y")

    contacts = fetch_vdt_contacts()
    new_members = filter_new_members(contacts, date_from, date_to)

    if not new_members:
        print("Keine neuen Mitglieder")
        return

    password = generate_password()

    pdf_bytes = build_pdf_encrypted(new_members, password)
    filename = f"VDT_Mitglieder_{date_from.strftime('%Y-%m')}.pdf"

    token = get_ms_token()

    send_email(
        pdf_bytes,
        filename,
        f"VDT Mitglieder {month_label}",
        f"Anbei die Liste der neuen Mitglieder ({len(new_members)}).",
        token
    )

    send_password_email(password, month_label, token)

    print("✅ Fertig")


if __name__ == "__main__":
    main()
