#!/usr/bin/env python3
"""
Analyse whether free/zero-revenue session types (e.g. 'free add to cart session')
inflate session count without lifting revenue, thereby diluting avg rev/session per day.

Steps:
  1. Show revenue distribution by session type (free vs paid)
  2. Tag each day: has_free_sessions Y/N
  3. Compare avg RPS on days with vs without free sessions
  4. Re-run the sessions-vs-avg-RPS correlation excluding free sessions
  5. Show per-type: avg price, session count, % days it appeared

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

FREE_KEYWORDS = ["free", "add to cart", "trial", "complimentary", "sample", "demo", "intro"]


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

def pearson_r(xs, ys):
    n = len(xs)
    if n < 3:
        return None
    mx, my = mean(xs), mean(ys)
    num = sum((x - mx) * (y - my) for x, y in zip(xs, ys))
    dx  = sum((x - mx)**2 for x in xs)**0.5
    dy  = sum((y - my)**2 for y in ys)**0.5
    return num / (dx * dy) if dx * dy else None

def is_free_type(type_str):
    t = str(type_str).lower().strip()
    return any(kw in t for kw in FREE_KEYWORDS)

def is_zero_price(price):
    return price is not None and price == 0.0


# ---------------------------------------------------------------------------
# Fetch & parse
# ---------------------------------------------------------------------------

def fetch_and_parse(service):
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SOURCE_SHEET}'!A1:Z2000",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    raw = result.get("values", [])
    if not raw:
        return []

    headers = [str(h).strip() for h in raw[0]]
    data_rows = raw[1:]

    date_idx   = 14   # Session Date
    price_idx  = 16   # Price
    leader_idx = 12   # Leader Name
    type_idx   = 10   # Type
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
        stype  = str(safe_get(row, type_idx,   "Unknown")).strip() or "Unknown"
        leader = str(safe_get(row, leader_idx, "Unknown")).strip() or "Unknown"
        records.append({
            "date":   date,
            "price":  price,
            "type":   stype,
            "leader": leader,
            "is_free_type":  is_free_type(stype),
            "is_zero_price": is_zero_price(price),
            # a session is "non-revenue" if its type name suggests free OR price is 0
            "is_non_revenue": is_free_type(stype) or is_zero_price(price),
        })

    return records


# ---------------------------------------------------------------------------
# Per-type summary
# ---------------------------------------------------------------------------

def type_summary(records):
    types = defaultdict(lambda: {"sessions": 0, "revenue": 0.0, "prices": [], "days": set()})
    for r in records:
        t = r["type"]
        types[t]["sessions"] += 1
        types[t]["revenue"]  += r["price"]
        types[t]["prices"].append(r["price"])
        types[t]["days"].add(r["date"])

    total_sessions = len(records)
    total_days     = len(set(r["date"] for r in records))

    result = []
    for t in sorted(types, key=lambda x: -types[x]["sessions"]):
        info = types[t]
        n    = info["sessions"]
        rev  = info["revenue"]
        prices = info["prices"]
        avg_p  = rev / n if n else 0
        zero_count = sum(1 for p in prices if p == 0)
        result.append({
            "type":            t,
            "sessions":        n,
            "pct_sessions":    round(n / total_sessions * 100, 1),
            "days_appeared":   len(info["days"]),
            "pct_days":        round(len(info["days"]) / total_days * 100, 1),
            "total_revenue":   round(rev, 2),
            "avg_price":       round(avg_p, 2),
            "zero_price_sessions": zero_count,
            "pct_zero_price":  round(zero_count / n * 100, 1) if n else 0,
            "flagged_free":    is_free_type(t),
        })
    return result


# ---------------------------------------------------------------------------
# Day-level analysis: with vs without free sessions
# ---------------------------------------------------------------------------

def day_analysis(records):
    # All sessions day table
    all_days  = defaultdict(lambda: {"sessions": 0, "revenue": 0.0,
                                      "paid_sessions": 0, "free_sessions": 0,
                                      "paid_revenue": 0.0})
    for r in records:
        d = r["date"]
        all_days[d]["sessions"] += 1
        all_days[d]["revenue"]  += r["price"]
        if r["is_non_revenue"]:
            all_days[d]["free_sessions"] += 1
        else:
            all_days[d]["paid_sessions"] += 1
            all_days[d]["paid_revenue"]  += r["price"]

    for d in all_days:
        n = all_days[d]["sessions"]
        p = all_days[d]["paid_sessions"]
        all_days[d]["avg_rps_all"]  = all_days[d]["revenue"] / n if n else 0
        all_days[d]["avg_rps_paid"] = (all_days[d]["paid_revenue"] / p) if p else 0
        all_days[d]["has_free"]     = all_days[d]["free_sessions"] > 0

    return all_days


def correlation_comparison(days):
    """
    Compare:
      A. sessions/day (all) vs avg RPS (all sessions)
      B. sessions/day (all) vs avg RPS (paid only)
      C. paid sessions/day   vs avg RPS (paid only)
    """
    sorted_days = sorted(days)

    xs_all   = [days[d]["sessions"]      for d in sorted_days]
    xs_paid  = [days[d]["paid_sessions"] for d in sorted_days]
    ys_all   = [days[d]["avg_rps_all"]   for d in sorted_days]
    ys_paid  = [days[d]["avg_rps_paid"]  for d in sorted_days if days[d]["paid_sessions"] > 0]
    xs_paid2 = [days[d]["paid_sessions"] for d in sorted_days if days[d]["paid_sessions"] > 0]

    r_all_sessions_vs_avg_rps_all   = pearson_r(xs_all, ys_all)
    r_all_sessions_vs_avg_rps_paid  = pearson_r(xs_all, ys_paid) if len(ys_paid) == len(xs_all) else \
                                       pearson_r([days[d]["sessions"] for d in sorted_days if days[d]["paid_sessions"] > 0], ys_paid)
    r_paid_sessions_vs_avg_rps_paid = pearson_r(xs_paid2, ys_paid)

    # Days with vs without free sessions
    days_with_free    = [d for d in sorted_days if days[d]["has_free"]]
    days_without_free = [d for d in sorted_days if not days[d]["has_free"]]

    avg_rps_with_free    = mean([days[d]["avg_rps_all"]  for d in days_with_free])    if days_with_free    else 0
    avg_rps_without_free = mean([days[d]["avg_rps_all"]  for d in days_without_free]) if days_without_free else 0
    avg_sessions_with    = mean([days[d]["sessions"]      for d in days_with_free])    if days_with_free    else 0
    avg_sessions_without = mean([days[d]["sessions"]      for d in days_without_free]) if days_without_free else 0

    return {
        "r_all_vs_all":     r_all_sessions_vs_avg_rps_all,
        "r_all_vs_paid":    r_all_sessions_vs_avg_rps_paid,
        "r_paid_vs_paid":   r_paid_sessions_vs_avg_rps_paid,
        "days_with_free":   len(days_with_free),
        "days_without_free": len(days_without_free),
        "avg_rps_with_free":    round(avg_rps_with_free, 2),
        "avg_rps_without_free": round(avg_rps_without_free, 2),
        "avg_sessions_with":    round(avg_sessions_with, 2),
        "avg_sessions_without": round(avg_sessions_without, 2),
        "rps_impact": round(avg_rps_with_free - avg_rps_without_free, 2),
    }


# ---------------------------------------------------------------------------
# Per-type impact on correlation
# ---------------------------------------------------------------------------

def type_impact_on_correlation(records, days, type_stats):
    sorted_days   = sorted(days)
    xs_base       = [days[d]["sessions"]    for d in sorted_days]
    ys_base       = [days[d]["avg_rps_all"] for d in sorted_days]
    r_base        = pearson_r(xs_base, ys_base)

    impact = []
    # Only analyse types with >= 5 sessions
    for ts in type_stats:
        if ts["sessions"] < 5:
            continue
        t = ts["type"]

        # Rebuild days excluding this type
        excl_days = defaultdict(lambda: {"sessions": 0, "revenue": 0.0})
        for r in records:
            if r["type"] == t:
                continue
            excl_days[r["date"]]["sessions"] += 1
            excl_days[r["date"]]["revenue"]  += r["price"]

        xs_excl = [excl_days[d]["sessions"] for d in sorted_days if excl_days[d]["sessions"] > 0]
        ys_excl = [excl_days[d]["revenue"] / excl_days[d]["sessions"]
                   for d in sorted_days if excl_days[d]["sessions"] > 0]
        r_excl  = pearson_r(xs_excl, ys_excl)
        shift   = (r_excl - r_base) if (r_excl is not None and r_base is not None) else None

        impact.append({
            "type":    t,
            "sessions": ts["sessions"],
            "avg_price": ts["avg_price"],
            "flagged_free": ts["flagged_free"],
            "r_excl":  r_excl,
            "shift":   shift,
        })

    return sorted(impact, key=lambda x: (abs(x["shift"]) if x["shift"] is not None else 0), reverse=True)


# ---------------------------------------------------------------------------
# Build output rows
# ---------------------------------------------------------------------------

def r_fmt(r):
    if r is None: return "N/A"
    if r >=  0.5: return f"{r:+.3f}  (strong scale)"
    if r >=  0.25: return f"{r:+.3f}  (mild scale)"
    if r <= -0.5: return f"{r:+.3f}  (strong dilution)"
    if r <= -0.25: return f"{r:+.3f}  (mild dilution)"
    return f"{r:+.3f}  (no clear pattern)"


def build_output(type_stats, corr, type_impact, days, records):
    rows = []
    now  = datetime.now().strftime("%Y-%m-%d %H:%M")

    def hr():
        rows.append(["══════════════════════════════════════════════════════════════"])

    rows.append([""])
    rows.append([""])
    hr()
    rows.append(["SESSION TYPE IMPACT ANALYSIS — Do free sessions dilute avg revenue per session?"])
    rows.append([f"Generated: {now}"])
    rows.append([""])

    # ── Per-type summary ───────────────────────────────────────────────────
    hr()
    rows.append(["ALL SESSION TYPES — Revenue Profile"])
    rows.append([""])
    rows.append([
        "Session Type",
        "Sessions",
        "% of All Sessions",
        "Days Appeared",
        "% of Days",
        "Total Revenue",
        "Avg Price",
        "Zero-Price Sessions",
        "% Zero-Price",
        "Flagged as Free Type?",
    ])
    for ts in type_stats:
        rows.append([
            ts["type"],
            ts["sessions"],
            ts["pct_sessions"],
            ts["days_appeared"],
            ts["pct_days"],
            ts["total_revenue"],
            ts["avg_price"],
            ts["zero_price_sessions"],
            ts["pct_zero_price"],
            "YES" if ts["flagged_free"] else "",
        ])
    rows.append([""])

    # ── Correlation comparison ────────────────────────────────────────────
    hr()
    rows.append(["CORRELATION COMPARISON — Effect of stripping free sessions"])
    rows.append([""])
    rows.append(["Scenario", "Correlation", "What it tells you"])
    rows.append([
        "All sessions/day  vs  avg rev/session (all)",
        r_fmt(corr["r_all_vs_all"]),
        "Baseline: includes free sessions in both numerator and denominator",
    ])
    rows.append([
        "All sessions/day  vs  avg rev/session (paid only)",
        r_fmt(corr["r_all_vs_paid"]),
        "Does total supply correlate with quality of paid sessions specifically?",
    ])
    rows.append([
        "Paid sessions/day  vs  avg rev/session (paid only)",
        r_fmt(corr["r_paid_vs_paid"]),
        "Pure paid supply signal — free sessions removed from both axes",
    ])
    rows.append([""])

    # ── Days with vs without free sessions ────────────────────────────────
    hr()
    rows.append(["DAYS WITH FREE SESSIONS vs DAYS WITHOUT"])
    rows.append([""])
    rows.append(["Metric", "Days WITH free sessions", "Days WITHOUT free sessions", "Difference"])
    rows.append(["Number of days",
                 corr["days_with_free"], corr["days_without_free"], ""])
    rows.append(["Avg total sessions/day",
                 corr["avg_sessions_with"], corr["avg_sessions_without"],
                 round(corr["avg_sessions_with"] - corr["avg_sessions_without"], 2)])
    rows.append(["Avg rev/session (all sessions)",
                 corr["avg_rps_with_free"], corr["avg_rps_without_free"],
                 corr["rps_impact"]])

    rps_delta = corr["rps_impact"]
    if rps_delta < -50:
        verdict = f"Free sessions DRAG avg RPS down by Rs{abs(rps_delta):.0f} on days they appear — clear dilution signal."
    elif rps_delta > 50:
        verdict = f"Days with free sessions actually have HIGHER avg RPS by Rs{rps_delta:.0f} — free sessions may attract more paid ones."
    else:
        verdict = f"Negligible difference ({rps_delta:+.0f}) — free sessions don't meaningfully shift avg RPS."
    rows.append(["VERDICT", verdict, "", ""])
    rows.append([""])

    # ── Per-type impact on overall correlation ────────────────────────────
    hr()
    rows.append(["PER-TYPE IMPACT ON CORRELATION (remove-one test, types with >= 5 sessions)"])
    rows.append(["(Positive shift = removing this type improves the correlation = type was suppressing it)"])
    rows.append([""])
    rows.append([
        "Session Type",
        "Sessions",
        "Avg Price",
        "Free-Type Flag",
        "Corr WITHOUT this type",
        "Shift in correlation",
        "Interpretation",
    ])
    for ti in type_impact:
        shift = ti["shift"]
        if shift is None:
            interp = "Insufficient data"
        elif shift > 0.05:
            interp = "Removing this type IMPROVES correlation → it was diluting the sessions↔RPS signal"
        elif shift < -0.05:
            interp = "Removing this type WEAKENS correlation → it was amplifying the sessions↔RPS signal"
        else:
            interp = "Minimal impact on overall correlation"

        rows.append([
            ti["type"],
            ti["sessions"],
            ti["avg_price"],
            "YES" if ti["flagged_free"] else "",
            r_fmt(ti["r_excl"]),
            f"{shift:+.4f}" if shift is not None else "N/A",
            interp,
        ])
    rows.append([""])

    # ── Day detail: paid vs free split (chart source) ─────────────────────
    hr()
    rows.append(["DAILY DETAIL — Paid vs Free split (use for charting)"])
    rows.append([""])
    rows.append(["Date", "Total Sessions", "Paid Sessions", "Free Sessions",
                 "Total Revenue", "Avg RPS (all)", "Avg RPS (paid only)", "Has Free Sessions?"])
    for d in sorted(days):
        day = days[d]
        rows.append([
            d,
            day["sessions"],
            day["paid_sessions"],
            day["free_sessions"],
            round(day["revenue"], 2),
            round(day["avg_rps_all"], 2),
            round(day["avg_rps_paid"], 2),
            "Y" if day["has_free"] else "N",
        ])

    return rows


# ---------------------------------------------------------------------------
# Write to sheet (append)
# ---------------------------------------------------------------------------

def append_to_sheet(service, new_rows):
    existing = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{OUTPUT_SHEET}'!A1:A3000",
    ).execute()
    next_row = len(existing.get("values", [])) + 1

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


def add_paid_vs_all_rps_chart(service, sheet_id, detail_header_row, num_rows):
    """Scatter: total sessions vs avg RPS all, and total sessions vs avg RPS paid."""
    requests = [{
        "addChart": {
            "chart": {
                "spec": {
                    "title": "Sessions vs Avg Rev/Session: All vs Paid-only",
                    "basicChart": {
                        "chartType": "SCATTER",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Total Sessions that Day"},
                            {"position": "LEFT_AXIS",   "title": "Avg Rev/Session (Rs)"},
                        ],
                        "series": [
                            {
                                "series": {"sourceRange": {"sources": [{
                                    "sheetId": sheet_id,
                                    "startRowIndex":  detail_header_row + 1,
                                    "endRowIndex":    detail_header_row + 1 + num_rows,
                                    "startColumnIndex": 5,   # col F: avg RPS all
                                    "endColumnIndex":   6,
                                }]}},
                                "targetAxis": "LEFT_AXIS",
                            },
                            {
                                "series": {"sourceRange": {"sources": [{
                                    "sheetId": sheet_id,
                                    "startRowIndex":  detail_header_row + 1,
                                    "endRowIndex":    detail_header_row + 1 + num_rows,
                                    "startColumnIndex": 6,   # col G: avg RPS paid
                                    "endColumnIndex":   7,
                                }]}},
                                "targetAxis": "LEFT_AXIS",
                            },
                        ],
                        "domains": [{"domain": {"sourceRange": {"sources": [{
                            "sheetId": sheet_id,
                            "startRowIndex":  detail_header_row + 1,
                            "endRowIndex":    detail_header_row + 1 + num_rows,
                            "startColumnIndex": 1,   # col B: total sessions
                            "endColumnIndex":   2,
                        }]}}}],
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {"sheetId": sheet_id,
                                       "rowIndex": detail_header_row + 1,
                                       "columnIndex": 9},
                        "widthPixels": 720, "heightPixels": 440,
                    }
                },
            }
        }
    }]
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body={"requests": requests}
    ).execute()
    print("Paid vs All RPS scatter chart added.")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    print("Building service ...")
    service = build_service()

    print("Fetching & parsing data ...")
    records = fetch_and_parse(service)
    print(f"{len(records)} usable records.")

    # Session type summary
    type_stats = type_summary(records)
    print(f"\n{len(type_stats)} unique session types found:")
    for ts in type_stats:
        flag = " <-- FREE/ZERO" if (ts["flagged_free"] or ts["pct_zero_price"] > 50) else ""
        print(f"  [{ts['sessions']:4d} sessions] {ts['type'][:55]:55s}  avg=Rs{ts['avg_price']:.0f}  zero={ts['pct_zero_price']:.0f}%{flag}")

    # Day-level analysis
    days = day_analysis(records)

    corr = correlation_comparison(days)
    print(f"\nCorrelation — all sessions vs avg RPS (all):  {corr['r_all_vs_all']}")
    print(f"Correlation — all sessions vs avg RPS (paid): {corr['r_all_vs_paid']}")
    print(f"Correlation — paid sessions vs avg RPS (paid):{corr['r_paid_vs_paid']}")
    print(f"\nDays WITH free sessions:    {corr['days_with_free']}  avg RPS={corr['avg_rps_with_free']}")
    print(f"Days WITHOUT free sessions: {corr['days_without_free']}  avg RPS={corr['avg_rps_without_free']}")
    print(f"RPS delta (with - without): {corr['rps_impact']}")

    # Per-type impact on correlation
    type_impact = type_impact_on_correlation(records, days, type_stats)
    print("\nTop type impacts on correlation:")
    for ti in type_impact[:8]:
        s = f"{ti['shift']:+.4f}" if ti["shift"] is not None else "N/A"
        print(f"  [{ti['sessions']:4d}] {ti['type'][:45]:45s}  shift={s}  avg=Rs{ti['avg_price']:.0f}")

    # Build and write output
    print("\nBuilding output rows ...")
    new_rows = build_output(type_stats, corr, type_impact, days, records)

    print("Appending to sheet ...")
    start_row = append_to_sheet(service, new_rows)
    print(f"Written at row {start_row}.")

    sheet_id = get_sheet_id(service, OUTPUT_SHEET)
    if sheet_id is not None:
        detail_header_rel = None
        for i, row in enumerate(new_rows):
            if row and row[0] == "Date" and len(row) > 1 and row[1] == "Total Sessions":
                detail_header_rel = i
                break

        if detail_header_rel is not None:
            offset = start_row - 1
            add_paid_vs_all_rps_chart(service, sheet_id,
                                       offset + detail_header_rel, len(days))

    print("\nDone.")
    print(f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")


if __name__ == "__main__":
    main()
