#!/usr/bin/env python3
"""
Analyse whether more sessions per day generates more revenue per session
(scale effect) or dilutes it (cannibalization).

Source  : 'excluding kavyal tamanna and offline' sheet
          in spreadsheet 158xyZv2gd7b7bN11WrWb1_V7EfW45seLkyy_zWoENi0
Filter  : rows where column H != 'N'
Output  : 'claude analysis' sheet (same spreadsheet)
"""

import base64
import json
import os
import platform
import sys
from collections import defaultdict
from datetime import datetime

import httplib2
import google_auth_httplib2
from google.oauth2 import service_account
from googleapiclient.discovery import build

SPREADSHEET_ID = "158xyZv2gd7b7bN11WrWb1_V7EfW45seLkyy_zWoENi0"
SOURCE_SHEET   = "excluding kavyal tamanna and offline"
OUTPUT_SHEET   = "claude analysis"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def build_service():
    if os.environ.get("GOOGLE_CREDENTIALS"):
        info = json.loads(base64.b64decode(os.environ["GOOGLE_CREDENTIALS"]))
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        h = httplib2.Http()
    else:
        creds_path = os.path.join(ROOT, "credentials.json")
        creds = service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        disable_ssl = platform.system() == "Windows"
        h = httplib2.Http(disable_ssl_certificate_validation=disable_ssl)

    auth_http = google_auth_httplib2.AuthorizedHttp(creds, http=h)
    return build("sheets", "v4", http=auth_http)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_price(val):
    if not val or str(val).strip() in ("", "-", "#N/A", "N/A"):
        return None
    s = str(val).replace(",", "").replace("₹", "").replace("$", "").strip()
    try:
        return float(s)
    except ValueError:
        return None


def col_letter_to_idx(letter):
    """Convert column letter (A, B, ... Z, AA, ...) to 0-based index."""
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


def safe_get(row, idx, default=""):
    return row[idx] if idx < len(row) else default


def mean(lst):
    return sum(lst) / len(lst) if lst else 0


def median(lst):
    if not lst:
        return 0
    s = sorted(lst)
    n = len(s)
    mid = n // 2
    return s[mid] if n % 2 else (s[mid - 1] + s[mid]) / 2


def pearson_r(xs, ys):
    """Simple Pearson correlation coefficient."""
    n = len(xs)
    if n < 2:
        return None
    mx, my = mean(xs), mean(ys)
    num = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    dx = sum((x - mx) ** 2 for x in xs) ** 0.5
    dy = sum((y - my) ** 2 for y in ys) ** 0.5
    return num / (dx * dy) if dx * dy else None


# ---------------------------------------------------------------------------
# Fetch & parse
# ---------------------------------------------------------------------------

def fetch_source(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SOURCE_SHEET}'!A1:Z2000",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    return result.get("values", [])


def identify_columns(headers):
    """
    Try to find column indices for date, price, and the H-filter column.
    We use heuristics based on common header names.
    """
    h_lower = [str(h).lower().strip() for h in headers]
    col_h_idx = 7  # default: 8th column (0-based = 7)

    # Guess date column
    date_idx = None
    for kw in ("session date", "date", "session_date"):
        for i, h in enumerate(h_lower):
            if kw in h:
                date_idx = i
                break
        if date_idx is not None:
            break
    if date_idx is None:
        date_idx = 0  # fallback

    # Guess price column
    price_idx = None
    for kw in ("price", "revenue", "amount", "fee", "total"):
        for i, h in enumerate(h_lower):
            if kw in h:
                price_idx = i
                break
        if price_idx is not None:
            break
    if price_idx is None:
        price_idx = 4  # fallback

    # Type column
    type_idx = None
    for kw in ("type",):
        for i, h in enumerate(h_lower):
            if kw in h:
                type_idx = i
                break

    # Leader column
    leader_idx = None
    for kw in ("leader", "facilitator", "host", "instructor"):
        for i, h in enumerate(h_lower):
            if kw in h:
                leader_idx = i
                break

    # Course column
    course_idx = None
    for kw in ("course", "program", "product", "service"):
        for i, h in enumerate(h_lower):
            if kw in h:
                course_idx = i
                break

    return {
        "date": date_idx,
        "price": price_idx,
        "type": type_idx,
        "leader": leader_idx,
        "course": course_idx,
        "filter_h": col_h_idx,
    }


def parse_date(val):
    """Try multiple date formats."""
    if not val:
        return None
    val = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y",
                "%d %b %Y", "%d-%b-%Y", "%B %d, %Y"):
        try:
            return datetime.strptime(val, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    # If it's a Google Sheets serial number (integer)
    try:
        serial = int(float(val))
        # Google Sheets epoch: Dec 30 1899
        from datetime import timedelta
        base = datetime(1899, 12, 30)
        return (base + timedelta(days=serial)).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    return None


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

def analyse(rows_raw, cols):
    """
    rows_raw: list of row lists (no header row), already filtered for col H != 'N'
    cols    : dict of column indices
    Returns analysis dict.
    """
    # Per-date aggregates
    date_data = defaultdict(lambda: {"sessions": 0, "revenue": 0.0, "prices": []})

    skipped_no_date = 0
    skipped_no_price = 0
    total_rows = 0

    for row in rows_raw:
        total_rows += 1
        raw_date  = safe_get(row, cols["date"])
        raw_price = safe_get(row, cols["price"])

        date = parse_date(raw_date)
        if not date:
            skipped_no_date += 1
            continue

        price = parse_price(raw_price)
        if price is None:
            skipped_no_price += 1
            continue

        date_data[date]["sessions"] += 1
        date_data[date]["revenue"]  += price
        date_data[date]["prices"].append(price)

    if not date_data:
        return None

    # Build per-day records
    records = []
    for date in sorted(date_data):
        d = date_data[date]
        n = d["sessions"]
        rev = d["revenue"]
        rps = rev / n if n else 0
        records.append({
            "date": date,
            "sessions": n,
            "revenue": round(rev, 2),
            "rev_per_session": round(rps, 2),
        })

    sessions_list = [r["sessions"] for r in records]
    rps_list      = [r["rev_per_session"] for r in records]
    rev_list      = [r["revenue"] for r in records]

    corr_sessions_vs_rps = pearson_r(sessions_list, rps_list)
    corr_sessions_vs_rev = pearson_r(sessions_list, rev_list)

    # Bucket analysis
    max_s = max(sessions_list) if sessions_list else 1
    buckets = {}
    if max_s <= 5:
        boundaries = [(1, 1), (2, 2), (3, 3), (4, 4), (5, 100)]
    else:
        # Dynamic buckets roughly in thirds
        t1 = max_s // 3
        t2 = (max_s * 2) // 3
        boundaries = [(1, t1), (t1+1, t2), (t2+1, 10000)]

    def bucket_label(lo, hi):
        if hi >= 10000:
            return f"{lo}+ sessions/day"
        elif lo == hi:
            return f"{lo} session/day"
        else:
            return f"{lo}–{hi} sessions/day"

    bucket_records = {bucket_label(lo, hi): [] for lo, hi in boundaries}
    for r in records:
        for lo, hi in boundaries:
            if lo <= r["sessions"] <= hi:
                bucket_records[bucket_label(lo, hi)].append(r)
                break

    bucket_stats = []
    for label in bucket_records:
        recs = bucket_records[label]
        if not recs:
            continue
        days = len(recs)
        avg_sessions  = mean([r["sessions"] for r in recs])
        avg_rps       = mean([r["rev_per_session"] for r in recs])
        avg_daily_rev = mean([r["revenue"] for r in recs])
        total_rev     = sum(r["revenue"] for r in recs)
        bucket_stats.append({
            "label": label,
            "days": days,
            "avg_sessions_per_day": round(avg_sessions, 1),
            "avg_rev_per_session": round(avg_rps, 2),
            "avg_daily_revenue": round(avg_daily_rev, 2),
            "total_revenue": round(total_rev, 2),
        })

    # Overall stats
    overall = {
        "total_days": len(records),
        "total_sessions": sum(sessions_list),
        "total_revenue": round(sum(rev_list), 2),
        "avg_sessions_per_day": round(mean(sessions_list), 2),
        "median_sessions_per_day": median(sessions_list),
        "avg_rev_per_session_overall": round(mean(rps_list), 2),
        "median_rev_per_session": round(median(rps_list), 2),
        "avg_daily_revenue": round(mean(rev_list), 2),
    }

    # Verdict
    r = corr_sessions_vs_rps
    if r is None:
        verdict = "Insufficient data to determine."
    elif r >= 0.25:
        verdict = (
            f"SCALE EFFECT: More sessions/day is associated with HIGHER revenue per session "
            f"(correlation = {r:.2f}). Adding sessions appears to grow revenue quality."
        )
    elif r <= -0.25:
        verdict = (
            f"DILUTION EFFECT: More sessions/day is associated with LOWER revenue per session "
            f"(correlation = {r:.2f}). Adding sessions appears to dilute per-session revenue."
        )
    else:
        verdict = (
            f"NO CLEAR RELATIONSHIP: Correlation between sessions/day and revenue/session = {r:.2f}. "
            f"Session count doesn't strongly predict per-session revenue."
        )

    return {
        "records": records,
        "overall": overall,
        "bucket_stats": bucket_stats,
        "corr_sessions_vs_rps": corr_sessions_vs_rps,
        "corr_sessions_vs_rev": corr_sessions_vs_rev,
        "verdict": verdict,
        "skipped_no_date": skipped_no_date,
        "skipped_no_price": skipped_no_price,
        "total_rows": total_rows,
    }


# ---------------------------------------------------------------------------
# Write to Google Sheets
# ---------------------------------------------------------------------------

def clear_and_write(service, data_2d, sheet_name=OUTPUT_SHEET, start="A1"):
    """Clear the output sheet then write data_2d starting at start."""
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'",
    ).execute()

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!{start}",
        valueInputOption="USER_ENTERED",
        body={"values": data_2d},
    ).execute()


def build_output(analysis, cols_detected, headers):
    """Compose the full 2D array to write to the output sheet."""
    rows = []
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    def blank(n=1):
        return [[""] * 1] * n

    def section(title):
        rows.append([title])
        rows.append([""])

    def hr():
        rows.append(["——————————————————————————————————————————"])

    # ── Title ──────────────────────────────────────────────────────────────
    rows.append(["SESSION VOLUME vs REVENUE ANALYSIS"])
    rows.append([f"Generated by Claude  |  {now}"])
    rows.append([f"Source: {SOURCE_SHEET}  |  Filter: Column H ≠ 'N'"])
    rows.append([""])
    hr()

    # ── Detected columns ───────────────────────────────────────────────────
    rows.append([""])
    rows.append(["DETECTED COLUMN MAPPING"])
    rows.append(["Field", "Header Name", "Column Index (0-based)"])
    for field, idx in cols_detected.items():
        hname = headers[idx] if idx is not None and idx < len(headers) else "?"
        rows.append([field.capitalize(), hname, idx])
    rows.append([""])

    # ── Data quality ───────────────────────────────────────────────────────
    o = analysis["overall"]
    rows.append(["DATA QUALITY"])
    rows.append(["Total rows after H-filter",          analysis["total_rows"]])
    rows.append(["Rows skipped (no parseable date)",   analysis["skipped_no_date"]])
    rows.append(["Rows skipped (no parseable price)",  analysis["skipped_no_price"]])
    rows.append(["Days with usable data",              o["total_days"]])
    rows.append([""])

    # ── Overall stats ──────────────────────────────────────────────────────
    hr()
    rows.append(["OVERALL STATISTICS"])
    rows.append([""])
    rows.append(["Metric", "Value"])
    rows.append(["Total sessions (filtered)",    o["total_sessions"]])
    rows.append(["Total revenue (filtered)",     o["total_revenue"]])
    rows.append(["Avg sessions per day",         o["avg_sessions_per_day"]])
    rows.append(["Median sessions per day",      o["median_sessions_per_day"]])
    rows.append(["Avg revenue per session",      o["avg_rev_per_session_overall"]])
    rows.append(["Median revenue per session",   o["median_rev_per_session"]])
    rows.append(["Avg daily revenue",            o["avg_daily_revenue"]])
    rows.append([""])

    # ── Key finding ────────────────────────────────────────────────────────
    hr()
    rows.append(["KEY FINDING"])
    rows.append([""])
    rows.append([analysis["verdict"]])
    rows.append([""])
    r1 = analysis["corr_sessions_vs_rps"]
    r2 = analysis["corr_sessions_vs_rev"]
    rows.append(["Correlation: sessions/day  vs  revenue/session",
                 round(r1, 4) if r1 is not None else "N/A"])
    rows.append(["Correlation: sessions/day  vs  total daily revenue",
                 round(r2, 4) if r2 is not None else "N/A"])
    rows.append([""])
    rows.append(["Interpretation guide:"])
    rows.append([" +0.25 to +1.0 → scale effect (more sessions = higher rev/session)"])
    rows.append([" -0.25 to -1.0 → dilution effect (more sessions = lower rev/session)"])
    rows.append([" -0.25 to +0.25 → no clear relationship"])
    rows.append([""])

    # ── Bucket analysis ────────────────────────────────────────────────────
    hr()
    rows.append(["BUCKET ANALYSIS  (sessions per day grouped)"])
    rows.append([""])
    rows.append([
        "Sessions/Day Bucket",
        "# Days",
        "Avg Sessions/Day",
        "Avg Rev/Session",
        "Avg Daily Revenue",
        "Total Revenue",
    ])
    for b in analysis["bucket_stats"]:
        rows.append([
            b["label"],
            b["days"],
            b["avg_sessions_per_day"],
            b["avg_rev_per_session"],
            b["avg_daily_revenue"],
            b["total_revenue"],
        ])
    rows.append([""])

    # ── Raw daily data (for charting) ──────────────────────────────────────
    hr()
    rows.append(["DAILY DETAIL  (use this table to create charts)"])
    rows.append([""])
    rows.append([
        "Date",
        "Sessions",
        "Total Revenue",
        "Revenue per Session",
    ])
    for r in analysis["records"]:
        rows.append([
            r["date"],
            r["sessions"],
            r["revenue"],
            r["rev_per_session"],
        ])

    return rows


# ---------------------------------------------------------------------------
# Ensure output sheet exists
# ---------------------------------------------------------------------------

def ensure_output_sheet(service):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    existing = {s["properties"]["title"] for s in meta.get("sheets", [])}
    if OUTPUT_SHEET not in existing:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": OUTPUT_SHEET}}}]},
        ).execute()
        print(f"Created sheet: {OUTPUT_SHEET}")
    else:
        print(f"Sheet already exists: {OUTPUT_SHEET}")


# ---------------------------------------------------------------------------
# Add a scatter chart: Sessions vs Revenue-per-Session
# ---------------------------------------------------------------------------

def get_sheet_id(service, title):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == title:
            return s["properties"]["sheetId"]
    return None


def add_scatter_chart(service, analysis_sheet_id, data_start_row, num_data_rows):
    """
    data_start_row: 0-based row index where the daily-detail header row sits.
    Columns in the daily detail table:
      Col A (0): Date
      Col B (1): Sessions
      Col C (2): Total Revenue
      Col D (3): Revenue per Session
    We'll plot Sessions (col B) on X, Revenue per Session (col D) on Y.
    """
    requests = [
        {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Sessions per Day vs Revenue per Session",
                        "basicChart": {
                            "chartType": "SCATTER",
                            "legendPosition": "BOTTOM_LEGEND",
                            "axis": [
                                {"position": "BOTTOM_AXIS", "title": "Sessions per Day"},
                                {"position": "LEFT_AXIS",   "title": "Revenue per Session (₹)"},
                            ],
                            "series": [
                                {
                                    "series": {
                                        "sourceRange": {
                                            "sources": [
                                                {
                                                    "sheetId": analysis_sheet_id,
                                                    "startRowIndex": data_start_row + 1,  # skip header
                                                    "endRowIndex":   data_start_row + 1 + num_data_rows,
                                                    "startColumnIndex": 3,  # col D: rev/session
                                                    "endColumnIndex":   4,
                                                }
                                            ]
                                        }
                                    },
                                    "targetAxis": "LEFT_AXIS",
                                }
                            ],
                            "domains": [
                                {
                                    "domain": {
                                        "sourceRange": {
                                            "sources": [
                                                {
                                                    "sheetId": analysis_sheet_id,
                                                    "startRowIndex": data_start_row + 1,
                                                    "endRowIndex":   data_start_row + 1 + num_data_rows,
                                                    "startColumnIndex": 1,  # col B: sessions
                                                    "endColumnIndex":   2,
                                                }
                                            ]
                                        }
                                    }
                                }
                            ],
                        },
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {
                                "sheetId": analysis_sheet_id,
                                "rowIndex": data_start_row + 1,
                                "columnIndex": 6,
                            },
                            "widthPixels":  700,
                            "heightPixels": 420,
                        }
                    },
                }
            }
        }
    ]
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests},
    ).execute()
    print("Scatter chart added.")


def add_bucket_bar_chart(service, analysis_sheet_id, bucket_header_row, num_buckets):
    """Bar chart: bucket labels (col A) vs avg rev/session (col D)."""
    requests = [
        {
            "addChart": {
                "chart": {
                    "spec": {
                        "title": "Avg Revenue per Session by Session-Count Bucket",
                        "basicChart": {
                            "chartType": "BAR",
                            "legendPosition": "BOTTOM_LEGEND",
                            "axis": [
                                {"position": "BOTTOM_AXIS", "title": "Avg Revenue per Session (₹)"},
                                {"position": "LEFT_AXIS",   "title": "Sessions/Day Bucket"},
                            ],
                            "series": [
                                {
                                    "series": {
                                        "sourceRange": {
                                            "sources": [
                                                {
                                                    "sheetId": analysis_sheet_id,
                                                    "startRowIndex": bucket_header_row + 1,
                                                    "endRowIndex":   bucket_header_row + 1 + num_buckets,
                                                    "startColumnIndex": 3,  # col D: avg rev/session
                                                    "endColumnIndex":   4,
                                                }
                                            ]
                                        }
                                    },
                                    "targetAxis": "BOTTOM_AXIS",
                                }
                            ],
                            "domains": [
                                {
                                    "domain": {
                                        "sourceRange": {
                                            "sources": [
                                                {
                                                    "sheetId": analysis_sheet_id,
                                                    "startRowIndex": bucket_header_row + 1,
                                                    "endRowIndex":   bucket_header_row + 1 + num_buckets,
                                                    "startColumnIndex": 0,  # col A: label
                                                    "endColumnIndex":   1,
                                                }
                                            ]
                                        }
                                    }
                                }
                            ],
                        },
                    },
                    "position": {
                        "overlayPosition": {
                            "anchorCell": {
                                "sheetId": analysis_sheet_id,
                                "rowIndex": bucket_header_row + 1,
                                "columnIndex": 6,
                            },
                            "widthPixels":  650,
                            "heightPixels": 380,
                        }
                    },
                }
            }
        }
    ]
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests},
    ).execute()
    print("Bar chart added.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Building Google Sheets service …")
    service = build_service()

    print(f"Fetching '{SOURCE_SHEET}' …")
    raw = fetch_source(service)

    if not raw:
        print("ERROR: Sheet is empty or not accessible.", file=sys.stderr)
        sys.exit(1)

    headers = [str(h).strip() for h in raw[0]] if raw else []
    data_rows = raw[1:]

    print(f"Headers ({len(headers)}): {headers}")
    print(f"Total data rows: {len(data_rows)}")

    # Identify columns
    cols = identify_columns(headers)
    print(f"Column mapping: {cols}")
    filter_col = cols["filter_h"]

    # Filter: column H (index 7) != 'N'
    filtered = []
    removed = 0
    for row in data_rows:
        h_val = str(safe_get(row, filter_col, "")).strip().upper()
        if h_val == "N":
            removed += 1
        else:
            filtered.append(row)

    print(f"Rows after removing col-H='N': {len(filtered)}  (removed {removed})")

    if not filtered:
        print("ERROR: No rows remain after filtering.", file=sys.stderr)
        sys.exit(1)

    # Run analysis
    print("Running analysis …")
    analysis = analyse(filtered, cols)
    if not analysis:
        print("ERROR: Analysis produced no results.", file=sys.stderr)
        sys.exit(1)

    # Print key findings to console
    print("\n" + "=" * 60)
    print("KEY FINDING:")
    print(analysis["verdict"])
    print(f"Correlation (sessions vs rev/session): {analysis['corr_sessions_vs_rps']}")
    print(f"Correlation (sessions vs daily rev):   {analysis['corr_sessions_vs_rev']}")
    print(f"Overall stats: {analysis['overall']}")
    print("=" * 60 + "\n")

    # Ensure output sheet exists
    ensure_output_sheet(service)

    # Build 2D output
    output_rows = build_output(analysis, cols, headers)

    # Write to sheet
    print(f"Writing to '{OUTPUT_SHEET}' …")
    clear_and_write(service, output_rows)

    # Find where the daily-detail table starts (to anchor charts)
    # Count rows before it: walk output_rows looking for "Date" header
    detail_header_row = None
    bucket_header_row = None
    for i, row in enumerate(output_rows):
        if row and row[0] == "Date":
            detail_header_row = i
        if row and row[0] == "Sessions/Day Bucket":
            bucket_header_row = i

    sheet_id = get_sheet_id(service, OUTPUT_SHEET)

    if sheet_id is not None:
        # Delete existing charts on the output sheet first
        meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        del_requests = []
        for s in meta.get("sheets", []):
            if s["properties"]["sheetId"] == sheet_id:
                for ch in s.get("charts", []):
                    del_requests.append({"deleteEmbeddedObject": {"objectId": ch["chartId"]}})
        if del_requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID, body={"requests": del_requests}
            ).execute()
            print(f"Deleted {len(del_requests)} existing chart(s).")

        num_days = len(analysis["records"])
        num_buckets = len(analysis["bucket_stats"])

        if detail_header_row is not None and num_days > 1:
            add_scatter_chart(service, sheet_id, detail_header_row, num_days)

        if bucket_header_row is not None and num_buckets > 1:
            add_bucket_bar_chart(service, sheet_id, bucket_header_row, num_buckets)
    else:
        print("WARNING: Could not find sheet ID; charts skipped.")

    print(f"\nDone. Open '{OUTPUT_SHEET}' in your Google Sheet to see results.")
    print(f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")


if __name__ == "__main__":
    main()
