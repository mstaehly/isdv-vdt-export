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
