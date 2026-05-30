name: Monatlicher VDT-Mitglieder-Export

on:
  schedule:
    - cron: '0 8 1 * *'   # jeden 1. des Monats um 08:00 UTC
  workflow_dispatch:      # manueller Start möglich

jobs:
  export-and-send:
    runs-on: ubuntu-latest

    steps:
      - name: Repository auschecken
        uses: actions/checkout@v4

      - name: Python installieren
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Abhängigkeiten installieren
        run: |
          pip install requests
          pip install pandas
          pip install openpyxl
          pip install pyzipper
          pip install python-dateutil

      - name: Script ausführen
        env:
          CAMPAI_API_KEY: ${{ secrets.CAMPAI_API_KEY }}
          SMTP_HOST: ${{ secrets.SMTP_HOST }}
          SMTP_PORT: ${{ secrets.SMTP_PORT }}
          SMTP_USER: ${{ secrets.SMTP_USER }}
          SMTP_PASSWORD: ${{ secrets.SMTP_PASSWORD }}

          MS_CLIENT_ID: ${{ secrets.MS_CLIENT_ID }}
          MS_TENANT_ID: ${{ secrets.MS_TENANT_ID }}
          MS_CLIENT_SECRET: ${{ secrets.MS_CLIENT_SECRET }}

          ZIP_PASSWORD: ${{ secrets.ZIP_PASSWORD }}
        run: python vdt_export.py
``
