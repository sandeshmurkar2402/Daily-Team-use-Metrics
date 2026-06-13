#!/usr/bin/env python3
"""
Deep analysis: which leader drives/masks the sessions-vs-avg-rev-per-session relationship?

For each leader:
  1. Their own avg revenue per session vs the overall baseline
  2. Their per-day session count correlation with that day's avg rev/session
  3. "Remove-one" test: overall correlation WITH vs WITHOUT that leader
  4. Day-level avg rev per session trend split by low/medium/high session count days

Output → 'claude analysis' sheet (appended after existing content)
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
# Auth (same pattern as fetch.py)
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
        serial = int(float(val))
        from datetime import timedelta
        return (datetime(1899, 12, 30) + timedelta(days=serial)).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        pass
    return None

def mean(lst):
    return sum(lst) / len(lst) if lst else 0

def median(lst):
    if not lst: return 0
    s = sorted(lst); n = len(s); mid = n // 2
    return s[mid] if n % 2 else (s[mid-1] + s[mid]) / 2

def pearson_r(xs, ys):
    n = len(xs)
    if n < 3:
        return None
    mx, my = mean(xs), mean(ys)
    num = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    dx  = sum((x - mx)**2 for x in xs)**0.5
    dy  = sum((y - my)**2 for y in ys)**0.5
    return num / (dx * dy) if dx * dy else None

def r_label(r):
    if r is None: return "N/A (too few days)"
    if r >=  0.5: return f"{r:+.2f}  ↑ strong scale effect"
    if r >=  0.25: return f"{r:+.2f}  ↑ mild scale effect"
    if r <= -0.5: return f"{r:+.2f}  ↓ strong dilution"
    if r <= -0.25: return f"{r:+.2f}  ↓ mild dilution"
    return f"{r:+.2f}  → no clear relationship"


# ---------------------------------------------------------------------------
# Fetch & filter
# ---------------------------------------------------------------------------

def fetch_and_filter(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SOURCE_SHEET}'!A1:Z2000",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    raw = result.get("values", [])
    if not raw:
        return [], []

    headers = [str(h).strip() for h in raw[0]]
    data_rows = raw[1:]

    # Column indices (identified in previous run)
    date_idx   = 14   # "Session Date"
    price_idx  = 16   # "Price"
    leader_idx = 12   # "Leader Name"
    filter_idx = 7    # col H: TBC_Non_Array

    filtered = []
    for row in data_rows:
        h_val = str(safe_get(row, filter_idx, "")).strip().upper()
        if h_val == "N":
            continue
        date  = parse_date(safe_get(row, date_idx))
        price = parse_price(safe_get(row, price_idx))
        if not date or price is None:
            continue
        leader = str(safe_get(row, leader_idx, "Unknown")).strip() or "Unknown"
        filtered.append({"date": date, "price": price, "leader": leader})

    return headers, filtered


# ---------------------------------------------------------------------------
# Build day-level aggregates
# ---------------------------------------------------------------------------

def build_day_table(records):
    """
    Returns dict: date → {
        sessions, revenue, avg_rps,
        leaders: {leader_name: {sessions, revenue}}
    }
    """
    days = defaultdict(lambda: {"sessions": 0, "revenue": 0.0, "leaders": defaultdict(lambda: {"sessions": 0, "revenue": 0.0})})
    for r in records:
        d = r["date"]; l = r["leader"]; p = r["price"]
        days[d]["sessions"] += 1
        days[d]["revenue"]  += p
        days[d]["leaders"][l]["sessions"] += 1
        days[d]["leaders"][l]["revenue"]  += p

    for d in days:
        n = days[d]["sessions"]
        days[d]["avg_rps"] = days[d]["revenue"] / n if n else 0

    return days


# ---------------------------------------------------------------------------
# Per-leader analysis
# ---------------------------------------------------------------------------

def analyse_leaders(records, days):
    # Collect all leaders
    leaders = defaultdict(lambda: {"sessions": 0, "revenue": 0.0, "prices": []})
    for r in records:
        l = r["leader"]
        leaders[l]["sessions"] += 1
        leaders[l]["revenue"]  += r["price"]
        leaders[l]["prices"].append(r["price"])

    overall_avg_rps = mean([days[d]["avg_rps"] for d in days])
    all_sessions    = sorted(days.keys())

    xs_all = [days[d]["sessions"] for d in all_sessions]
    ys_all = [days[d]["avg_rps"]  for d in all_sessions]
    r_overall = pearson_r(xs_all, ys_all)

    leader_stats = []

    for leader in sorted(leaders, key=lambda l: -leaders[l]["sessions"]):
        info = leaders[leader]
        n_sessions   = info["sessions"]
        avg_price    = info["revenue"] / n_sessions if n_sessions else 0
        price_vs_avg = avg_price - overall_avg_rps

        # Days this leader appeared
        days_with_leader    = [d for d in all_sessions if leader in days[d]["leaders"]]
        days_without_leader = [d for d in all_sessions if leader not in days[d]["leaders"]]

        # 1. On days the leader appears: how does THEIR session count correlate with that day's avg RPS?
        xs_leader_days = [days[d]["leaders"][leader]["sessions"] for d in days_with_leader]
        ys_leader_days = [days[d]["avg_rps"]                    for d in days_with_leader]
        r_leader_own   = pearson_r(xs_leader_days, ys_leader_days)

        # 2. Remove-one test: overall correlation excluding this leader's sessions
        #    Rebuild per-day aggregates without this leader
        days_excl = {}
        for d in all_sessions:
            sessions_excl = days[d]["sessions"] - days[d]["leaders"][leader]["sessions"] \
                            if leader in days[d]["leaders"] else days[d]["sessions"]
            revenue_excl  = days[d]["revenue"]  - days[d]["leaders"][leader]["revenue"] \
                            if leader in days[d]["leaders"] else days[d]["revenue"]
            if sessions_excl > 0:
                days_excl[d] = {
                    "sessions": sessions_excl,
                    "avg_rps":  revenue_excl / sessions_excl,
                }

        xs_excl = [days_excl[d]["sessions"] for d in sorted(days_excl)]
        ys_excl = [days_excl[d]["avg_rps"]  for d in sorted(days_excl)]
        r_excl  = pearson_r(xs_excl, ys_excl)

        corr_shift = (r_excl - r_overall) if (r_excl is not None and r_overall is not None) else None

        leader_stats.append({
            "leader":              leader,
            "total_sessions":      n_sessions,
            "days_appeared":       len(days_with_leader),
            "avg_sessions_per_appearance": round(mean(xs_leader_days), 2) if xs_leader_days else 0,
            "avg_price":           round(avg_price, 2),
            "price_vs_avg":        round(price_vs_avg, 2),
            "r_own_sessions_vs_day_rps":   r_leader_own,
            "r_excl":              r_excl,
            "corr_shift":          corr_shift,
        })

    return leader_stats, r_overall, overall_avg_rps


# ---------------------------------------------------------------------------
# Session-count buckets → avg RPS trend
# ---------------------------------------------------------------------------

def bucket_trend(days):
    """Shows how avg-RPS changes as sessions/day increases, bucket by bucket."""
    counts = sorted(set(days[d]["sessions"] for d in days))
    if not counts:
        return []

    # Create ~6 even buckets
    lo, hi = counts[0], counts[-1]
    span = hi - lo
    if span == 0:
        buckets = [(lo, hi)]
    else:
        step = max(1, span // 6)
        buckets = []
        cur = lo
        while cur <= hi:
            buckets.append((cur, min(cur + step - 1, hi)))
            cur += step

    result = []
    for blo, bhi in buckets:
        matched = [d for d in days if blo <= days[d]["sessions"] <= bhi]
        if not matched:
            continue
        n_days       = len(matched)
        avg_sessions = mean([days[d]["sessions"] for d in matched])
        avg_rps      = mean([days[d]["avg_rps"]  for d in matched])
        avg_daily_rev= mean([days[d]["revenue"]  for d in matched])
        result.append({
            "label":         f"{blo}–{bhi} sessions/day",
            "days":          n_days,
            "avg_sessions":  round(avg_sessions, 1),
            "avg_rps":       round(avg_rps, 2),
            "avg_daily_rev": round(avg_daily_rev, 2),
        })
    return result


# ---------------------------------------------------------------------------
# Build 2-D output rows
# ---------------------------------------------------------------------------

def build_output(leader_stats, r_overall, overall_avg_rps, trend, days):
    rows = []
    now  = datetime.now().strftime("%Y-%m-%d %H:%M")

    def hr():
        rows.append(["══════════════════════════════════════════════════════════════"])

    rows.append([""])
    rows.append([""])
    hr()
    rows.append(["LEADER IMPACT ANALYSIS  —  Who drives the sessions↔avg-RPS relationship?"])
    rows.append([f"Generated: {now}"])
    rows.append([""])

    # ── Overall baseline ───────────────────────────────────────────────────
    rows.append(["OVERALL BASELINE"])
    rows.append(["Overall avg revenue per session (filtered)", round(overall_avg_rps, 2)])
    rows.append(["Overall corr: sessions/day vs avg rev/session", round(r_overall, 4) if r_overall else "N/A"])
    rows.append(["Total analysis days", len(days)])
    rows.append([""])

    # ── Avg RPS trend by session count ────────────────────────────────────
    hr()
    rows.append(["AVG REVENUE PER SESSION  as sessions/day increases"])
    rows.append(["(Does avg rev/session go up or down as we add more sessions?)"])
    rows.append([""])
    rows.append(["Sessions/Day Range", "# Days", "Avg Sessions/Day", "Avg Rev/Session (₹)", "Avg Daily Revenue (₹)", "Trend vs prev bucket"])
    prev_rps = None
    for b in trend:
        if prev_rps is None:
            trend_vs_prev = "—  (baseline)"
        else:
            delta = b["avg_rps"] - prev_rps
            pct   = (delta / prev_rps * 100) if prev_rps else 0
            arrow = "▲" if delta >= 0 else "▼"
            trend_vs_prev = f"{arrow} {abs(delta):.0f} ({pct:+.1f}%)"
        rows.append([
            b["label"], b["days"], b["avg_sessions"], b["avg_rps"], b["avg_daily_rev"], trend_vs_prev
        ])
        prev_rps = b["avg_rps"]
    rows.append([""])

    # ── Leader summary table ───────────────────────────────────────────────
    hr()
    rows.append(["PER-LEADER BREAKDOWN"])
    rows.append([""])
    rows.append([
        "Leader",
        "Total Sessions",
        "Days Appeared",
        "Avg Sessions per Appearance",
        "Avg Rev/Session (₹)",
        "vs Overall Avg (₹)",
        "Corr: own sessions ↔ day avg RPS",
        "Corr WITHOUT this leader",
        "Corr Shift (removing them)",
        "Interpretation",
    ])

    for s in leader_stats:
        r_own  = s["r_own_sessions_vs_day_rps"]
        r_excl = s["r_excl"]
        shift  = s["corr_shift"]

        # Interpretation
        interp_parts = []
        if r_own is not None:
            if r_own <= -0.3:
                interp_parts.append("More of their sessions → lower day avg RPS (diluter)")
            elif r_own >= 0.3:
                interp_parts.append("More of their sessions → higher day avg RPS (amplifier)")
            else:
                interp_parts.append("Neutral effect on day avg RPS")
        if shift is not None:
            if abs(shift) >= 0.1:
                direction = "correlation rises" if shift > 0 else "correlation falls"
                interp_parts.append(f"removing them {direction} by {abs(shift):.2f}")
        if s["price_vs_avg"] < -200:
            interp_parts.append(f"low-price leader ({s['avg_price']:.0f} vs avg {overall_avg_rps:.0f})")
        elif s["price_vs_avg"] > 200:
            interp_parts.append(f"high-price leader ({s['avg_price']:.0f} vs avg {overall_avg_rps:.0f})")

        rows.append([
            s["leader"],
            s["total_sessions"],
            s["days_appeared"],
            s["avg_sessions_per_appearance"],
            s["avg_price"],
            s["price_vs_avg"],
            r_label(r_own),
            r_label(r_excl),
            f"{shift:+.3f}" if shift is not None else "N/A",
            " | ".join(interp_parts) if interp_parts else "—",
        ])

    rows.append([""])

    # ── Top movers ────────────────────────────────────────────────────────
    hr()
    rows.append(["KEY LEADERS TO WATCH  (largest correlation shift when removed)"])
    rows.append([""])
    rows.append(["Leader", "Corr Shift", "What it means"])

    sortable = [s for s in leader_stats if s["corr_shift"] is not None and s["total_sessions"] >= 5]
    top_movers = sorted(sortable, key=lambda s: abs(s["corr_shift"]), reverse=True)[:10]

    for s in top_movers:
        shift = s["corr_shift"]
        if shift > 0.05:
            meaning = f"Their presence SUPPRESSES the correlation. Without them, overall corr = {s['r_excl']:+.2f}."
        elif shift < -0.05:
            meaning = f"Their presence BOOSTS the correlation. Without them, overall corr = {s['r_excl']:+.2f}."
        else:
            meaning = "Minimal influence on overall pattern."
        rows.append([s["leader"], f"{shift:+.3f}", meaning])

    rows.append([""])

    # ── Daily detail (chart source) ────────────────────────────────────────
    hr()
    rows.append(["DAILY DETAIL — for charting (sessions/day, avg RPS, daily revenue)"])
    rows.append([""])
    rows.append(["Date", "Sessions", "Total Revenue", "Avg Rev/Session"])
    for d in sorted(days):
        rows.append([d, days[d]["sessions"], round(days[d]["revenue"], 2), round(days[d]["avg_rps"], 2)])

    return rows


# ---------------------------------------------------------------------------
# Write to sheet (append after existing content)
# ---------------------------------------------------------------------------

def append_to_sheet(service, new_rows):
    """Read current last row, then append below it."""
    existing = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{OUTPUT_SHEET}'!A1:A2000",
    ).execute()
    current_data = existing.get("values", [])
    next_row = len(current_data) + 1

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


def add_trend_chart(service, sheet_id, header_row_0based, num_rows):
    """Line/column chart: sessions/day bucket vs avg RPS."""
    requests = [{
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Avg Revenue per Session as Sessions/Day Increases",
                    "basicChart": {
                        "chartType": "COLUMN",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Sessions/Day Bucket"},
                            {"position": "LEFT_AXIS",   "title": "Avg Rev/Session (₹)"},
                        ],
                        "series": [{
                            "series": {
                                "sourceRange": {"sources": [{
                                    "sheetId":        sheet_id,
                                    "startRowIndex":  header_row_0based + 1,
                                    "endRowIndex":    header_row_0based + 1 + num_rows,
                                    "startColumnIndex": 3,
                                    "endColumnIndex":   4,
                                }]}
                            },
                            "targetAxis": "LEFT_AXIS",
                        }],
                        "domains": [{"domain": {"sourceRange": {"sources": [{
                            "sheetId":        sheet_id,
                            "startRowIndex":  header_row_0based + 1,
                            "endRowIndex":    header_row_0based + 1 + num_rows,
                            "startColumnIndex": 0,
                            "endColumnIndex":   1,
                        }]}}}],
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {"sheetId": sheet_id, "rowIndex": header_row_0based + 1, "columnIndex": 7},
                        "widthPixels": 680, "heightPixels": 380,
                    }
                },
            }
        }
    }]
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body={"requests": requests}
    ).execute()
    print("Trend chart added.")


def add_leader_scatter(service, sheet_id, detail_header_row, num_days):
    """Scatter: sessions/day vs avg RPS, using daily detail table."""
    requests = [{
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Daily Sessions vs Avg Revenue per Session (all leaders combined)",
                    "basicChart": {
                        "chartType": "SCATTER",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Sessions that Day"},
                            {"position": "LEFT_AXIS",   "title": "Avg Rev/Session (₹)"},
                        ],
                        "series": [{
                            "series": {
                                "sourceRange": {"sources": [{
                                    "sheetId":        sheet_id,
                                    "startRowIndex":  detail_header_row + 1,
                                    "endRowIndex":    detail_header_row + 1 + num_days,
                                    "startColumnIndex": 3,
                                    "endColumnIndex":   4,
                                }]}
                            },
                            "targetAxis": "LEFT_AXIS",
                        }],
                        "domains": [{"domain": {"sourceRange": {"sources": [{
                            "sheetId":        sheet_id,
                            "startRowIndex":  detail_header_row + 1,
                            "endRowIndex":    detail_header_row + 1 + num_days,
                            "startColumnIndex": 1,
                            "endColumnIndex":   2,
                        }]}}}],
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {"sheetId": sheet_id, "rowIndex": detail_header_row + 1, "columnIndex": 6},
                        "widthPixels": 700, "heightPixels": 420,
                    }
                },
            }
        }
    }]
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body={"requests": requests}
    ).execute()
    print("Leader scatter chart added.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Building service …")
    service = build_service()

    print("Fetching & filtering data …")
    headers, records = fetch_and_filter(service)
    print(f"{len(records)} usable records after filter.")

    days = build_day_table(records)
    print(f"{len(days)} unique days.")

    leader_stats, r_overall, overall_avg_rps = analyse_leaders(records, days)
    trend = bucket_trend(days)

    # Print leader highlights
    print(f"\nOverall correlation (sessions/day vs avg RPS): {r_overall:.4f}")
    print(f"Overall avg RPS: Rs{overall_avg_rps:.2f}")
    print("\nTop 5 leaders by session count:")
    for s in leader_stats[:5]:
        r_own  = f"{s['r_own_sessions_vs_day_rps']:.2f}" if s['r_own_sessions_vs_day_rps'] is not None else "N/A"
        shift  = f"{s['corr_shift']:+.3f}"               if s['corr_shift']                is not None else "N/A"
        print(f"  {s['leader']:30s}  sessions={s['total_sessions']:4d}  "
              f"avg_price=Rs{s['avg_price']:.0f}  r_own={r_own}  shift={shift}")

    print("\nBuilding output …")
    new_rows = build_output(leader_stats, r_overall, overall_avg_rps, trend, days)

    print("Appending to 'claude analysis' sheet …")
    start_row = append_to_sheet(service, new_rows)
    print(f"Written starting at row {start_row}.")

    sheet_id = get_sheet_id(service, OUTPUT_SHEET)
    if sheet_id is not None:
        # Find header rows inside new_rows
        trend_header_rel  = None
        detail_header_rel = None
        for i, row in enumerate(new_rows):
            if row and row[0] == "Sessions/Day Range":
                trend_header_rel = i
            if row and row[0] == "Date":
                detail_header_rel = i

        offset = start_row - 1  # convert to 0-based

        if trend_header_rel is not None and len(trend) > 1:
            add_trend_chart(service, sheet_id, offset + trend_header_rel, len(trend))

        if detail_header_rel is not None and len(days) > 1:
            add_leader_scatter(service, sheet_id, offset + detail_header_rel, len(days))

    print("\nDone.")
    print(f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")


if __name__ == "__main__":
    main()
