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
RECIPIENT_EMAIL = "info@ms-sounddesign.com"  # TEST

TAG_VDT = "VDT"

# ----------------------------------------


def generate_password(length=12):
    alphabet = string.ascii_letters + string.digits
    return ''.join(secrets.choice(alphabet) for _ in range(length))


def get_date_range():
    today = datetime.now(timezone.utc)
    first = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last = first - relativedelta(months=1)
    return last, first


def get_ms_token():
    url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": MS_CLIENT_ID,
        "client_secret": MS_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }

    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]


def fetch_vdt_contacts():
    url = "https://api.campai.com/contacts"
    headers = {"Authorization": CAMPAI_API_KEY}

    contacts = []
    skip = 0

    while True:
        params = {
            "organisation": CAMPAI_ORG_ID,
            "tags": TAG_VDT,
            "type": "member",
            "limit": 100,
            "skip": skip
        }

        r = requests.get(url, headers=headers, params=params)
        r.raise_for_status()

        data = r.json()
        contacts.extend(data)

        if len(data) < 100:
            break

        skip += 100

    return contacts


def filter_new_members(contacts, date_from, date_to):
    result = []

    for c in contacts:
        entry = c.get("mergeTags", {}).get("enterDate")
        if not entry:
            continue

        try:
            d, m, y = entry.split("/")
            dt = datetime(int(y), int(m), int(d), tzinfo=timezone.utc)

            if date_from <= dt < date_to:
                result.append(c)
        except:
            pass

    return result


def build_pdf_encrypted(contacts, password):
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        path = tmp.name

    data = [["Vorname", "Nachname", "E-Mail", "Eintrittsdatum"]]

    for c in contacts:
        m = c.get("mergeTags", {})
        data.append([
            m.get("personFirstName", ""),
            m.get("personLastName", ""),
            m.get("email", ""),
            m.get("enterDate", "")
        ])

    encrypt = pdfencrypt.StandardEncryption(password)

    doc = SimpleDocTemplate(path, pagesize=A4, encrypt=encrypt)
    table = Table(data)
    doc.build([table])

    with open(path, "rb") as f:
        pdf_bytes = f.read()

    os.unlink(path)
    return pdf_bytes


# ✅ HTML SIGNATURE (zentral, für beide Mails)
def build_signature():
    return """
<p>Viele Grüße<br>
Marc</p>

<p>-<br>
isdv e.V.<br>
-Sirius Office Center-<br>
Hanauer Landstr. 328-330<br>
60314 Frankfurt am Main</p>

<p>Marc Stähly<br>
2. Vorsitzender</p>

<p>
T.: +49 178 2 44 76 25<br>
M.: <a href="mailto:m.staehly@isdv.net">m.staehly@isdv.net</a><br>
W.: <a href="https://www.isdv.net">www.isdv.net</a>
</p>

<p>
Registernummer im Lobbyregister des Bundestages: R000099<br>
Amtsgericht Frankfurt am Main, VR16763<br>
Vertretungsberechtigter Vorstand: Marcus Pohl, Marc Stähly
</p>
"""


def send_email(file_bytes, filename, subject, token):
    print("📧 Sende Hauptmail an:", RECIPIENT_EMAIL)

    body = f"""
<html><body>

<p>Hallo Alex,</p>

<p>anbei die aktuelle Liste der VDT-Mitglieder.</p>

{build_signature()}

</body></html>
"""

    attachment = base64.b64encode(file_bytes).decode()

    message = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": [{"emailAddress": {"address": RECIPIENT_EMAIL}}],
            "attachments": [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": filename,
                "contentBytes": attachment
            }]
        }
    }

    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    requests.post(url, headers=headers, json=message).raise_for_status()


def send_password_email(password, token):
    print("🔑 Sende Passwort-Mail")

    body = f"""
<html><body>

<p>Hallo Alex,</p>

<p>anbei das Passwort für die PDF-Datei.</p>

<p><b>Passwort:</b><br>{password}</p>

{build_signature()}

</body></html>
"""

    message = {
        "message": {
            "subject": "isdv e.V. - Passwort für Mitglieder-Datei",
            "body": {"contentType": "HTML", "content": body},
            "toRecipients": [{"emailAddress": {"address": RECIPIENT_EMAIL}}]
        }
    }

    url = f"https://graph.microsoft.com/v1.0/users/{SENDER_EMAIL}/sendMail"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    requests.post(url, headers=headers, json=message).raise_for_status()


def main():
    date_from, date_to = get_date_range()

    month_label = date_from.strftime("%m/%Y")
    filename = f"VDT_Mitglieder_{date_from.strftime('%Y-%m')}.pdf"
    subject = f"isdv e.V. - VDT-Mitgliederliste {month_label}"

    contacts = fetch_vdt_contacts()
    new_members = filter_new_members(contacts, date_from, date_to)

    if not new_members:
        print("⚠️ Keine neuen Mitglieder – sende Testdaten")
        new_members = contacts[:5]

    password = generate_password()
    print("🔑 Passwort:", password)

    pdf_bytes = build_pdf_encrypted(new_members, password)

    token = get_ms_token()

    send_email(pdf_bytes, filename, subject, token)
    send_password_email(password, token)

    print("✅ Fertig")


if __name__ == "__main__":
    main()
