#!/usr/bin/env python3
"""Fetches metrics from Google Sheets Dashboard_data tab and writes data/metrics.json"""
import json
import os
import sys
import base64
import platform
from datetime import datetime

import httplib2
import google_auth_httplib2
from google.oauth2 import service_account
from googleapiclient.discovery import build

SHEET_NAME = "Dashboard_data"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
METRIC_START = 5  # Cols: MTD, Last_30d, Last_7d, Yesterday, Date, then metrics

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def build_service():
    # GitHub Actions supplies credentials as base64-encoded JSON secret
    if os.environ.get("GOOGLE_CREDENTIALS"):
        info = json.loads(base64.b64decode(os.environ["GOOGLE_CREDENTIALS"]))
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        h = httplib2.Http()
    else:
        creds_path = os.path.join(ROOT, "credentials.json")
        creds = service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        # Windows machines often fail SSL cert verification locally
        disable_ssl = platform.system() == "Windows"
        h = httplib2.Http(disable_ssl_certificate_validation=disable_ssl)

    auth_http = google_auth_httplib2.AuthorizedHttp(creds, http=h)
    return build("sheets", "v4", http=auth_http)


def parse_num(val):
    if not val or str(val).strip() in ("", "#N/A", "-"):
        return None
    s = str(val).replace(",", "").replace(" ", "").strip()
    if s.endswith("%"):
        s = s[:-1]
    try:
        return float(s)
    except ValueError:
        return None


def fmt_avg(num, sample_val=""):
    """Format an averaged number to match the style of the sample value."""
    if num is None:
        return "-"
    sample = str(sample_val).strip() if sample_val else ""
    if sample.endswith("%"):
        return f"{num:.1f}%"
    if num >= 1000:
        return f"{num:,.0f}"
    if num == int(num):
        return str(int(num))
    return f"{num:.1f}"


def avg_col(rows, col_idx):
    vals = [parse_num(r[col_idx]) for r in rows if col_idx < len(r)]
    vals = [v for v in vals if v is not None]
    return sum(vals) / len(vals) if vals else None


def flag_rows(all_rows, headers, flag_name):
    idx = headers.index(flag_name)
    return [r for r in all_rows if len(r) > idx and r[idx] == "Y"]


def main():
    with open(os.path.join(ROOT, "config.json")) as f:
        config = json.load(f)

    service = build_service()

    result = service.spreadsheets().values().get(
        spreadsheetId=config["spreadsheetId"],
        range=f"{SHEET_NAME}!A1:AZ500",
    ).execute()

    raw = result.get("values", [])
    if not raw:
        print("ERROR: Sheet is empty", file=sys.stderr)
        sys.exit(1)

    headers = raw[0]
    rows = raw[1:]

    rows_yesterday = flag_rows(rows, headers, "Yesterday")
    rows_7d = flag_rows(rows, headers, "Last_7d")
    rows_30d = flag_rows(rows, headers, "Last_30d")
    rows_mtd = flag_rows(rows, headers, "MTD")

    y_row = rows_yesterday[0] if rows_yesterday else None
    y_date = y_row[4] if (y_row and len(y_row) > 4) else ""

    metrics = []
    for i, name in enumerate(headers[METRIC_START:], METRIC_START):
        y_val = y_row[i] if (y_row and i < len(y_row)) else ""
        metrics.append({
            "name": name,
            "yesterday": y_val if y_val else "-",
            "avg7d": fmt_avg(avg_col(rows_7d, i), y_val),
            "avg30d": fmt_avg(avg_col(rows_30d, i), y_val),
            "mtdAvg": fmt_avg(avg_col(rows_mtd, i), y_val),
        })

    output = {
        "lastUpdated": datetime.now(__import__("datetime").timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "yesterdayDate": y_date,
        "metrics": metrics,
    }

    os.makedirs(os.path.join(ROOT, "data"), exist_ok=True)
    out_path = os.path.join(ROOT, "data", "metrics.json")
    with open(out_path, "w") as f:
        json.dump(output, f, indent=2)

    print(f"OK: {len(metrics)} metrics written. Yesterday = {y_date}")


if __name__ == "__main__":
    main()
