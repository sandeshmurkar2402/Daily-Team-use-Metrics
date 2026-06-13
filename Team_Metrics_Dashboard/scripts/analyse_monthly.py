#!/usr/bin/env python3
"""
Month-wise analysis:
  - Monthly avg revenue per session
  - Each day classified: ABOVE / BELOW monthly avg
  - Exact drivers behind each group: session type, leader, course, price
Appends to 'claude analysis' sheet.
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
        h = httplib2.Http(disable_ssl_certificate_validation=(platform.system() == "Windows"))
    auth_http = google_auth_httplib2.AuthorizedHttp(creds, http=h)
    return build("sheets", "v4", http=auth_http)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def safe_get(row, idx, default=""):
    return row[idx] if idx < len(row) else default

def parse_price(val):
    if not val or str(val).strip() in ("", "-", "#N/A", "N/A"):
        return None
    s = str(val).replace(",", "").replace("₹", "").replace("$", "").strip()
    try:
        return float(s)
    except ValueError:
        return None

def parse_date(val):
    if not val:
        return None
    val = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y",
                "%d %b %Y", "%d-%b-%Y", "%B %d, %Y"):
        try:
            return datetime.strptime(val, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    try:
        from datetime import timedelta
        serial = int(float(val))
        return (datetime(1899, 12, 30) + timedelta(days=serial)).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    return None

def mean(lst):
    return sum(lst) / len(lst) if lst else 0

def ym(date_str):
    """Return 'YYYY-MM' from 'YYYY-MM-DD'."""
    return date_str[:7]

def month_label(ym_str):
    """'2024-03' → 'Mar 2024'."""
    try:
        return datetime.strptime(ym_str, "%Y-%m").strftime("%b %Y")
    except Exception:
        return ym_str

def top_n(counter_dict, n=5):
    """Return top-n items sorted by value desc."""
    return sorted(counter_dict.items(), key=lambda x: -x[1])[:n]


# ---------------------------------------------------------------------------
# Fetch & parse
# ---------------------------------------------------------------------------

def fetch_records(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SOURCE_SHEET}'!A1:Z2000",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    raw = result.get("values", [])
    if not raw:
        return []

    data_rows = raw[1:]

    date_idx   = 14   # Session Date
    price_idx  = 16   # Price
    leader_idx = 12   # Leader Name
    type_idx   = 10   # Type
    course_idx = 15   # Course Name
    filter_idx = 7    # col H

    records = []
    for row in data_rows:
        h_val = str(safe_get(row, filter_idx, "")).strip().upper()
        if h_val == "N":
            continue
        date  = parse_date(safe_get(row, date_idx))
        price = parse_price(safe_get(row, price_idx))
        if not date or price is None:
            continue
        records.append({
            "date":   date,
            "month":  ym(date),
            "price":  price,
            "type":   str(safe_get(row, type_idx,   "Unknown")).strip() or "Unknown",
            "leader": str(safe_get(row, leader_idx, "Unknown")).strip() or "Unknown",
            "course": str(safe_get(row, course_idx, "Unknown")).strip() or "Unknown",
        })
    return records


# ---------------------------------------------------------------------------
# Build day-level table
# ---------------------------------------------------------------------------

def build_day_table(records):
    """
    day_table[date] = {
        month, sessions, revenue, avg_rps,
        sessions_by_type   {type:  {sessions, revenue}},
        sessions_by_leader {leader:{sessions, revenue}},
        sessions_by_course {course:{sessions, revenue}},
    }
    """
    days = defaultdict(lambda: {
        "month": None, "sessions": 0, "revenue": 0.0,
        "by_type":   defaultdict(lambda: {"sessions": 0, "revenue": 0.0}),
        "by_leader": defaultdict(lambda: {"sessions": 0, "revenue": 0.0}),
        "by_course": defaultdict(lambda: {"sessions": 0, "revenue": 0.0}),
    })
    for r in records:
        d = r["date"]
        days[d]["month"]     = r["month"]
        days[d]["sessions"] += 1
        days[d]["revenue"]  += r["price"]
        for key, grp in [("by_type", r["type"]),
                         ("by_leader", r["leader"]),
                         ("by_course", r["course"])]:
            days[d][key][grp]["sessions"] += 1
            days[d][key][grp]["revenue"]  += r["price"]

    for d in days:
        n = days[d]["sessions"]
        days[d]["avg_rps"] = days[d]["revenue"] / n if n else 0

    return days


# ---------------------------------------------------------------------------
# Month-level analysis
# ---------------------------------------------------------------------------

def analyse_months(days):
    """
    For each month:
      - monthly_avg_rps
      - per-day classification: ABOVE / BELOW / AT
      - aggregate drivers for above-days and below-days
    """
    # Group days by month
    month_days = defaultdict(list)
    for d in sorted(days):
        month_days[days[d]["month"]].append(d)

    months_analysis = {}

    for month in sorted(month_days):
        m_days  = month_days[month]
        rps_list = [days[d]["avg_rps"] for d in m_days]
        monthly_avg = mean(rps_list)

        above_days = [d for d in m_days if days[d]["avg_rps"] > monthly_avg]
        below_days = [d for d in m_days if days[d]["avg_rps"] < monthly_avg]
        at_days    = [d for d in m_days if days[d]["avg_rps"] == monthly_avg]

        def aggregate_drivers(day_list):
            """For a set of days, aggregate session counts/revenue by type/leader/course."""
            by_type   = defaultdict(lambda: {"sessions": 0, "revenue": 0.0})
            by_leader = defaultdict(lambda: {"sessions": 0, "revenue": 0.0})
            by_course = defaultdict(lambda: {"sessions": 0, "revenue": 0.0})
            total_sessions = 0
            total_revenue  = 0.0

            for d in day_list:
                total_sessions += days[d]["sessions"]
                total_revenue  += days[d]["revenue"]
                for k, v in days[d]["by_type"].items():
                    by_type[k]["sessions"]   += v["sessions"]
                    by_type[k]["revenue"]    += v["revenue"]
                for k, v in days[d]["by_leader"].items():
                    by_leader[k]["sessions"] += v["sessions"]
                    by_leader[k]["revenue"]  += v["revenue"]
                for k, v in days[d]["by_course"].items():
                    by_course[k]["sessions"] += v["sessions"]
                    by_course[k]["revenue"]  += v["revenue"]

            def enrich(d):
                return {k: {
                    "sessions": v["sessions"],
                    "revenue":  round(v["revenue"], 2),
                    "avg_price": round(v["revenue"] / v["sessions"], 2) if v["sessions"] else 0,
                } for k, v in d.items()}

            return {
                "total_sessions": total_sessions,
                "total_revenue":  round(total_revenue, 2),
                "avg_rps":        round(total_revenue / total_sessions, 2) if total_sessions else 0,
                "by_type":        enrich(by_type),
                "by_leader":      enrich(by_leader),
                "by_course":      enrich(by_course),
            }

        # Per-day detail for this month
        day_detail = []
        for d in sorted(m_days):
            day = days[d]
            diff = day["avg_rps"] - monthly_avg
            flag = "ABOVE" if diff > 0 else ("BELOW" if diff < 0 else "AT AVG")

            # Top driver for this day (leader with highest avg price if above, lowest if below)
            leaders_sorted = sorted(
                day["by_leader"].items(),
                key=lambda x: x[1]["revenue"] / x[1]["sessions"] if x[1]["sessions"] else 0,
                reverse=(diff >= 0)
            )
            top_leader = leaders_sorted[0][0] if leaders_sorted else "—"

            types_sorted = sorted(
                day["by_type"].items(),
                key=lambda x: x[1]["sessions"],
                reverse=True
            )
            dominant_type = types_sorted[0][0] if types_sorted else "—"

            courses_sorted = sorted(
                day["by_course"].items(),
                key=lambda x: x[1]["revenue"] / x[1]["sessions"] if x[1]["sessions"] else 0,
                reverse=(diff >= 0)
            )
            top_course = courses_sorted[0][0] if courses_sorted else "—"

            day_detail.append({
                "date":         d,
                "sessions":     day["sessions"],
                "avg_rps":      round(day["avg_rps"], 2),
                "vs_monthly":   round(diff, 2),
                "flag":         flag,
                "top_leader":   top_leader,
                "dominant_type": dominant_type,
                "top_course":   top_course,
            })

        months_analysis[month] = {
            "monthly_avg":    round(monthly_avg, 2),
            "total_days":     len(m_days),
            "above_days":     len(above_days),
            "below_days":     len(below_days),
            "total_sessions": sum(days[d]["sessions"] for d in m_days),
            "total_revenue":  round(sum(days[d]["revenue"] for d in m_days), 2),
            "above_drivers":  aggregate_drivers(above_days),
            "below_drivers":  aggregate_drivers(below_days),
            "day_detail":     day_detail,
        }

    return months_analysis


# ---------------------------------------------------------------------------
# Build output rows
# ---------------------------------------------------------------------------

def top_by_sessions(driver_dict, n=4):
    return sorted(driver_dict.items(), key=lambda x: -x[1]["sessions"])[:n]

def top_by_avg_price(driver_dict, n=4, reverse=True):
    return sorted(driver_dict.items(),
                  key=lambda x: x[1]["avg_price"],
                  reverse=reverse)[:n]


def build_output(months_analysis):
    rows = []
    now  = datetime.now().strftime("%Y-%m-%d %H:%M")

    def hr():
        rows.append(["══════════════════════════════════════════════════════════════"])

    rows.append([""])
    rows.append([""])
    hr()
    rows.append(["MONTH-WISE ANALYSIS — Above vs Below Monthly Avg Revenue per Session"])
    rows.append([f"Generated: {now}"])
    rows.append([
        "For each month: days where avg rev/session was ABOVE or BELOW the monthly average,",
        "with breakdown by session type, leader, and course."
    ])
    rows.append([""])

    # ── Summary table across all months ────────────────────────────────────
    hr()
    rows.append(["MONTH SUMMARY"])
    rows.append([""])
    rows.append([
        "Month", "Total Days", "Total Sessions", "Total Revenue",
        "Monthly Avg RPS", "Days ABOVE Avg", "Days BELOW Avg",
        "Above Avg RPS", "Below Avg RPS", "Gap (Above - Below)",
    ])
    for month in sorted(months_analysis):
        ma = months_analysis[month]
        above_rps = ma["above_drivers"]["avg_rps"]
        below_rps = ma["below_drivers"]["avg_rps"]
        gap       = round(above_rps - below_rps, 2) if above_rps and below_rps else ""
        rows.append([
            month_label(month),
            ma["total_days"],
            ma["total_sessions"],
            ma["total_revenue"],
            ma["monthly_avg"],
            ma["above_days"],
            ma["below_days"],
            above_rps,
            below_rps,
            gap,
        ])
    rows.append([""])

    # ── Per-month deep-dive ─────────────────────────────────────────────────
    for month in sorted(months_analysis):
        ma  = months_analysis[month]
        lbl = month_label(month)
        hr()
        rows.append([f"MONTH: {lbl}   |   Monthly Avg RPS = Rs{ma['monthly_avg']}   |   "
                     f"{ma['total_days']} days, {ma['total_sessions']} sessions, Rs{ma['total_revenue']} revenue"])
        rows.append([""])

        # ── Above-average days breakdown ────────────────────────────────
        ad = ma["above_drivers"]
        rows.append([f"  ABOVE-AVERAGE DAYS ({ma['above_days']} days)  —  avg RPS = Rs{ad['avg_rps']}"])
        rows.append([""])

        if ad["total_sessions"]:
            # Session type breakdown
            rows.append(["    By Session Type:"])
            rows.append(["    ", "Type", "Sessions", "Avg Price", "Reason it lifts avg"])
            for t, v in top_by_sessions(ad["by_type"]):
                reason = (
                    "Premium price type" if v["avg_price"] > ma["monthly_avg"]
                    else "Below-avg price but dominant in volume"
                )
                rows.append(["    ", t, v["sessions"], f"Rs{v['avg_price']}", reason])

            rows.append([""])
            rows.append(["    By Leader (top 5 by avg price):"])
            rows.append(["    ", "Leader", "Sessions", "Avg Price", "vs Monthly Avg"])
            for ldr, v in top_by_avg_price(ad["by_leader"], n=5, reverse=True):
                vs = round(v["avg_price"] - ma["monthly_avg"], 0)
                rows.append(["    ", ldr, v["sessions"],
                             f"Rs{v['avg_price']}", f"{vs:+.0f}"])

            rows.append([""])
            rows.append(["    By Course (top 5 by avg price):"])
            rows.append(["    ", "Course", "Sessions", "Avg Price", "vs Monthly Avg"])
            for crs, v in top_by_avg_price(ad["by_course"], n=5, reverse=True):
                vs = round(v["avg_price"] - ma["monthly_avg"], 0)
                rows.append(["    ", crs, v["sessions"],
                             f"Rs{v['avg_price']}", f"{vs:+.0f}"])
        else:
            rows.append(["    (No above-average days this month)"])

        rows.append([""])

        # ── Below-average days breakdown ────────────────────────────────
        bd = ma["below_drivers"]
        rows.append([f"  BELOW-AVERAGE DAYS ({ma['below_days']} days)  —  avg RPS = Rs{bd['avg_rps']}"])
        rows.append([""])

        if bd["total_sessions"]:
            rows.append(["    By Session Type:"])
            rows.append(["    ", "Type", "Sessions", "Avg Price", "Why it drags avg down"])
            for t, v in top_by_sessions(bd["by_type"]):
                reason = (
                    "Below-avg price — dilutes the day's avg" if v["avg_price"] < ma["monthly_avg"]
                    else "Decent price but low volume"
                )
                rows.append(["    ", t, v["sessions"], f"Rs{v['avg_price']}", reason])

            rows.append([""])
            rows.append(["    By Leader (bottom 5 by avg price — lowest first):"])
            rows.append(["    ", "Leader", "Sessions", "Avg Price", "vs Monthly Avg"])
            for ldr, v in top_by_avg_price(bd["by_leader"], n=5, reverse=False):
                vs = round(v["avg_price"] - ma["monthly_avg"], 0)
                rows.append(["    ", ldr, v["sessions"],
                             f"Rs{v['avg_price']}", f"{vs:+.0f}"])

            rows.append([""])
            rows.append(["    By Course (bottom 5 by avg price — lowest first):"])
            rows.append(["    ", "Course", "Sessions", "Avg Price", "vs Monthly Avg"])
            for crs, v in top_by_avg_price(bd["by_course"], n=5, reverse=False):
                vs = round(v["avg_price"] - ma["monthly_avg"], 0)
                rows.append(["    ", crs, v["sessions"],
                             f"Rs{v['avg_price']}", f"{vs:+.0f}"])
        else:
            rows.append(["    (No below-average days this month)"])

        rows.append([""])

        # ── Day-by-day detail for this month ────────────────────────────
        rows.append(["  DAY-BY-DAY DETAIL"])
        rows.append(["  ", "Date", "Sessions", "Avg RPS", "vs Monthly Avg",
                     "Flag", "Dominant Type", "Key Leader", "Top Course"])
        for dd in ma["day_detail"]:
            rows.append([
                "  ",
                dd["date"],
                dd["sessions"],
                dd["avg_rps"],
                f"{dd['vs_monthly']:+.0f}",
                dd["flag"],
                dd["dominant_type"],
                dd["top_leader"],
                dd["top_course"],
            ])
        rows.append([""])

    return rows


# ---------------------------------------------------------------------------
# Write to sheet
# ---------------------------------------------------------------------------

def expand_sheet_if_needed(service, sheet_id, needed_rows):
    """Append blank rows to the sheet if current row count is too small."""
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta.get("sheets", []):
        if s["properties"]["sheetId"] == sheet_id:
            current = s["properties"]["gridProperties"]["rowCount"]
            if needed_rows > current:
                extra = needed_rows - current + 500
                service.spreadsheets().batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body={"requests": [{"appendDimension": {
                        "sheetId":   sheet_id,
                        "dimension": "ROWS",
                        "length":    extra,
                    }}]},
                ).execute()
                print(f"Expanded sheet by {extra} rows (was {current}).")
            break


def append_to_sheet(service, new_rows):
    # Find current last used row (read only col A)
    existing = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{OUTPUT_SHEET}'!A1:A2000",
    ).execute()
    next_row = len(existing.get("values", [])) + 1
    needed   = next_row + len(new_rows) + 10

    sheet_id = get_sheet_id(service, OUTPUT_SHEET)
    if sheet_id is not None:
        expand_sheet_if_needed(service, sheet_id, needed)

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{OUTPUT_SHEET}'!A{next_row}",
        valueInputOption="USER_ENTERED",
        body={"values": new_rows},
    ).execute()
    return next_row


def get_sheet_id(service, title):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == title:
            return s["properties"]["sheetId"]
    return None


def add_monthly_summary_chart(service, sheet_id, summary_header_row, num_months):
    """Grouped bar: monthly avg RPS vs above-avg-days RPS vs below-avg-days RPS."""
    requests = [{
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Monthly Avg RPS: Overall vs Above-Avg Days vs Below-Avg Days",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Month"},
                            {"position": "LEFT_AXIS",   "title": "Avg Revenue per Session (Rs)"},
                        ],
                        "series": [
                            # Monthly Avg RPS (col E, index 4)
                            {
                                "series": {"sourceRange": {"sources": [{
                                    "sheetId": sheet_id,
                                    "startRowIndex":  summary_header_row + 1,
                                    "endRowIndex":    summary_header_row + 1 + num_months,
                                    "startColumnIndex": 4,
                                    "endColumnIndex":   5,
                                }]}},
                                "targetAxis": "LEFT_AXIS",
                            },
                            # Above Avg RPS (col H, index 7)
                            {
                                "series": {"sourceRange": {"sources": [{
                                    "sheetId": sheet_id,
                                    "startRowIndex":  summary_header_row + 1,
                                    "endRowIndex":    summary_header_row + 1 + num_months,
                                    "startColumnIndex": 7,
                                    "endColumnIndex":   8,
                                }]}},
                                "targetAxis": "LEFT_AXIS",
                            },
                            # Below Avg RPS (col I, index 8)
                            {
                                "series": {"sourceRange": {"sources": [{
                                    "sheetId": sheet_id,
                                    "startRowIndex":  summary_header_row + 1,
                                    "endRowIndex":    summary_header_row + 1 + num_months,
                                    "startColumnIndex": 8,
                                    "endColumnIndex":   9,
                                }]}},
                                "targetAxis": "LEFT_AXIS",
                            },
                        ],
                        "domains": [{"domain": {"sourceRange": {"sources": [{
                            "sheetId": sheet_id,
                            "startRowIndex":  summary_header_row + 1,
                            "endRowIndex":    summary_header_row + 1 + num_months,
                            "startColumnIndex": 0,
                            "endColumnIndex":   1,
                        }]}}}],
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId":     sheet_id,
                            "rowIndex":    summary_header_row + 1,
                            "columnIndex": 11,
                        },
                        "widthPixels": 800, "heightPixels": 420,
                    }
                },
            }
        }
    }]
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body={"requests": requests}
    ).execute()
    print("Monthly summary chart added.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Building service ...")
    service = build_service()

    print("Fetching data ...")
    records = fetch_records(service)
    print(f"{len(records)} usable records.")

    days = build_day_table(records)
    print(f"{len(days)} unique days across {len(set(d[:7] for d in days))} months.")

    months_analysis = analyse_months(days)

    # Print console summary
    print(f"\n{'Month':<12} {'Avg RPS':>8} {'Days':>5} {'Above':>6} {'Below':>6} {'Above RPS':>10} {'Below RPS':>10} {'Gap':>8}")
    print("-" * 70)
    for month in sorted(months_analysis):
        ma = months_analysis[month]
        above_rps = ma["above_drivers"]["avg_rps"]
        below_rps = ma["below_drivers"]["avg_rps"]
        gap = f"{above_rps - below_rps:+.0f}" if above_rps and below_rps else "N/A"
        print(f"{month_label(month):<12} {ma['monthly_avg']:>8.0f} {ma['total_days']:>5} "
              f"{ma['above_days']:>6} {ma['below_days']:>6} "
              f"{above_rps:>10.0f} {below_rps:>10.0f} {gap:>8}")

    print("\nBuilding output rows ...")
    new_rows = build_output(months_analysis)

    print("Appending to sheet ...")
    start_row = append_to_sheet(service, new_rows)
    print(f"Written at row {start_row}.")

    sheet_id = get_sheet_id(service, OUTPUT_SHEET)
    if sheet_id is not None:
        summary_header_rel = None
        for i, row in enumerate(new_rows):
            if row and row[0] == "Month" and len(row) > 4 and row[4] == "Monthly Avg RPS":
                summary_header_rel = i
                break

        if summary_header_rel is not None:
            offset = start_row - 1
            add_monthly_summary_chart(
                service, sheet_id,
                offset + summary_header_rel,
                len(months_analysis)
            )

    print("\nDone.")
    print(f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")


if __name__ == "__main__":
    main()
