#!/usr/bin/env python3
"""
Creates one stacked-column chart per month in 'stack_charts' worksheet.

Layout per month (written to the sheet):
  Col A        : Date
  Cols B..I    : Top-7 leaders by revenue (value = leader_rev / day_total_sessions)
                 8th col = "Others" (all remaining leaders combined)
  Col J        : Upper control limit  (monthly_avg + 1 std-dev, constant per month)
  Col K        : Lower control limit  (monthly_avg - 1 std-dev, constant per month)
  Col L        : Day Avg RPS          (for reference / cross-check)
  Col M        : [blank separator]
  Cols N..U    : % label strings "XX%" per leader (same rows as B..I, used as chart data labels)

Chart (combo):
  - Leader cols (B..I) : COLUMN, STACKED
  - Upper / Lower (J,K): LINE (float at absolute value, not stacked)
  - customLabelData per leader series → % label cols (N..U)
"""

import base64, json, math, os, platform, sys
from collections import defaultdict
from datetime import datetime

import httplib2, google_auth_httplib2
from google.oauth2 import service_account
from googleapiclient.discovery import build

SPREADSHEET_ID = "158xyZv2gd7b7bN11WrWb1_V7EfW45seLkyy_zWoENi0"
SOURCE_SHEET   = "excluding kavyal tamanna and offline"
OUTPUT_SHEET   = "stack_charts"
SCOPES         = ["https://www.googleapis.com/auth/spreadsheets"]
MAX_LEADERS    = 7      # top-N leaders per month; rest → "Others"
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


# ─── Auth ───────────────────────────────────────────────────────────────────

def build_service():
    if os.environ.get("GOOGLE_CREDENTIALS"):
        info  = json.loads(base64.b64decode(os.environ["GOOGLE_CREDENTIALS"]))
        creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
        h     = httplib2.Http()
    else:
        creds = service_account.Credentials.from_service_account_file(
            os.path.join(ROOT, "credentials.json"), scopes=SCOPES)
        h = httplib2.Http(disable_ssl_certificate_validation=(platform.system() == "Windows"))
    return build("sheets", "v4", http=google_auth_httplib2.AuthorizedHttp(creds, http=h))


# ─── Helpers ────────────────────────────────────────────────────────────────

def safe_get(row, idx, default=""):
    return row[idx] if idx < len(row) else default

def parse_price(val):
    if not val or str(val).strip() in ("", "-", "#N/A", "N/A"):
        return None
    try:
        return float(str(val).replace(",", "").replace("Rs", "").replace("$", "").strip())
    except ValueError:
        return None

def parse_date(val):
    if not val: return None
    val = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%d/%Y",
                "%d %b %Y", "%d-%b-%Y", "%B %d, %Y"):
        try: return datetime.strptime(val, fmt).strftime("%Y-%m-%d")
        except ValueError: pass
    try:
        from datetime import timedelta
        return (datetime(1899, 12, 30) + timedelta(days=int(float(val)))).strftime("%Y-%m-%d")
    except: return None

def mean(lst): return sum(lst) / len(lst) if lst else 0

def stdev(lst):
    if len(lst) < 2: return 0
    m = mean(lst)
    return math.sqrt(sum((x - m)**2 for x in lst) / len(lst))

def month_label(ym):
    try: return datetime.strptime(ym, "%Y-%m").strftime("%b %Y")
    except: return ym

def col_letter(n):          # 0-based → "A", "B", ... "Z", "AA", ...
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def a1(row0, col0):         # 0-based → "A1" style
    return f"{col_letter(col0)}{row0 + 1}"


# ─── Fetch & filter ─────────────────────────────────────────────────────────

def fetch_records(service):
    res = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SOURCE_SHEET}'!A1:Z2000",
        valueRenderOption="UNFORMATTED_VALUE",
    ).execute()
    raw = res.get("values", [])
    records = []
    for row in raw[1:]:
        if str(safe_get(row, 7, "")).strip().upper() == "N":
            continue
        date  = parse_date(safe_get(row, 14))
        price = parse_price(safe_get(row, 16))
        if not date or price is None: continue
        records.append({
            "date":   date,
            "month":  date[:7],
            "price":  price,
            "leader": str(safe_get(row, 12, "Unknown")).strip() or "Unknown",
        })
    return records


# ─── Build month pivot ───────────────────────────────────────────────────────

def build_month_pivots(records):
    """
    Returns dict:  month → {
        days: {date → {leader → revenue, __sessions → n, __revenue → total}},
        leaders_ranked: [leader names sorted by total month revenue desc],
        monthly_avg_rps, upper, lower
    }
    """
    month_data = defaultdict(lambda: defaultdict(lambda: {
        "__sessions": 0, "__revenue": 0.0
    }))
    for r in records:
        d = month_data[r["month"]][r["date"]]
        d["__sessions"] += 1
        d["__revenue"]  += r["price"]
        d[r["leader"]]   = d.get(r["leader"], 0.0) + r["price"]

    result = {}
    for month in sorted(month_data):
        days = month_data[month]

        # Daily avg RPS list
        rps_list = [days[d]["__revenue"] / days[d]["__sessions"]
                    for d in days if days[d]["__sessions"]]

        m_avg  = mean(rps_list)
        m_std  = stdev(rps_list)
        upper  = round(m_avg + m_std, 2)
        lower  = max(0, round(m_avg - m_std, 2))

        # Rank leaders by total revenue this month
        leader_rev = defaultdict(float)
        for d in days:
            for k, v in days[d].items():
                if not k.startswith("__"):
                    leader_rev[k] += v
        leaders_ranked = [l for l, _ in
                          sorted(leader_rev.items(), key=lambda x: -x[1])]

        result[month] = {
            "days":            dict(days),
            "leaders_ranked":  leaders_ranked,
            "monthly_avg_rps": round(m_avg, 2),
            "upper":           upper,
            "lower":           lower,
        }
    return result


# ─── Build data rows for one month ──────────────────────────────────────────

def month_rows(month, mpivot):
    """
    Returns (header_row, data_rows, label_rows, top_leaders)
    All leader values = leader_revenue / total_day_sessions  (contribution to avg RPS)
    Label rows contain "XX%" strings for data labels.
    """
    days_dict      = mpivot["days"]
    leaders_ranked = mpivot["leaders_ranked"]
    monthly_avg    = mpivot["monthly_avg_rps"]
    upper          = mpivot["upper"]
    lower          = mpivot["lower"]

    top_leaders  = leaders_ranked[:MAX_LEADERS]
    has_others   = len(leaders_ranked) > MAX_LEADERS
    all_cols     = top_leaders + (["Others"] if has_others else [])

    # Header
    header = ["Date"] + all_cols + ["Upper Limit", "Lower Limit", "Day Avg RPS",
              ""] + [f"{l} %" for l in all_cols]

    data_rows  = []
    label_rows = []

    for date in sorted(days_dict):
        day      = days_dict[date]
        n        = day["__sessions"]
        rev      = day["__revenue"]
        day_rps  = round(rev / n, 2) if n else 0

        # Leader contributions to avg RPS
        leader_vals = []
        others_rev  = 0.0
        for l in top_leaders:
            lr = day.get(l, 0.0)
            leader_vals.append(round(lr / n, 2) if n else 0)

        if has_others:
            others_rev = sum(day.get(l, 0.0)
                             for l in leaders_ranked[MAX_LEADERS:])
            leader_vals.append(round(others_rev / n, 2) if n else 0)

        # % of daily revenue per leader
        label_vals = []
        for idx, l in enumerate(all_cols):
            lr = day.get(l, 0.0) if l != "Others" else others_rev
            pct = round(lr / rev * 100) if rev else 0
            label_vals.append(f"{pct}%" if pct > 0 else "")

        data_rows.append(
            [date] + leader_vals + [upper, lower, day_rps, ""] + label_vals
        )
        label_rows.append(label_vals)  # parallel (same row, label cols only)

    return header, data_rows, label_rows, all_cols


# ─── Ensure sheet exists & has enough rows ───────────────────────────────────

def ensure_sheet(service, title, needed_rows=3000, needed_cols=30):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    existing = {s["properties"]["title"]: s["properties"]
                for s in meta.get("sheets", [])}

    requests = []
    if title not in existing:
        requests.append({"addSheet": {"properties": {
            "title": title,
            "gridProperties": {"rowCount": needed_rows, "columnCount": needed_cols}
        }}})
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body={"requests": requests}
        ).execute()
        print(f"Created sheet: {title}")
        # re-fetch to get sheetId
        meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        existing = {s["properties"]["title"]: s["properties"]
                    for s in meta.get("sheets", [])}
    else:
        props = existing[title]
        sid   = props["sheetId"]
        gr    = props["gridProperties"]
        expand_reqs = []
        if gr["rowCount"] < needed_rows:
            expand_reqs.append({"appendDimension": {
                "sheetId": sid, "dimension": "ROWS",
                "length": needed_rows - gr["rowCount"]
            }})
        if gr["columnCount"] < needed_cols:
            expand_reqs.append({"appendDimension": {
                "sheetId": sid, "dimension": "COLUMNS",
                "length": needed_cols - gr["columnCount"]
            }})
        if expand_reqs:
            service.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID, body={"requests": expand_reqs}
            ).execute()
            print(f"Expanded sheet '{title}'.")

    # Return sheetId
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == title:
            return s["properties"]["sheetId"]
    raise RuntimeError(f"Sheet '{title}' not found after creation.")


def clear_sheet(service, title):
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range=f"'{title}'"
    ).execute()


# ─── Delete existing charts on sheet ────────────────────────────────────────

def delete_charts(service, sheet_id):
    meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    del_reqs = []
    for s in meta.get("sheets", []):
        if s["properties"]["sheetId"] == sheet_id:
            for ch in s.get("charts", []):
                del_reqs.append({"deleteEmbeddedObject": {"objectId": ch["chartId"]}})
    if del_reqs:
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body={"requests": del_reqs}
        ).execute()
        print(f"Deleted {len(del_reqs)} existing chart(s).")


# ─── Write data to sheet ────────────────────────────────────────────────────

def write_range(service, range_a1, values):
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{OUTPUT_SHEET}'!{range_a1}",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()


# ─── Build chart request ────────────────────────────────────────────────────

def build_chart_request(sheet_id, month, monthly_avg, upper, lower,
                        header_row, data_row_start, data_row_end,
                        chart_anchor_row, chart_anchor_col,
                        num_leader_cols):
    """
    header_row     : 0-based row index of the column-header row
    data_row_start : 0-based row index of first data row
    data_row_end   : 0-based exclusive end row

    Column layout (0-based):
      0       : Date
      1..N    : Leader contribution cols  (N = num_leader_cols)
      N+1     : Upper limit
      N+2     : Lower limit
      N+3     : Day Avg RPS
      N+4     : blank
      N+5..   : % label strings per leader
    """
    N = num_leader_cols

    # Range from header row so Sheets picks up series names from col headers
    def src(start_col, end_col):
        return {
            "sheetId":          sheet_id,
            "startRowIndex":    header_row,          # include header
            "endRowIndex":      data_row_end,
            "startColumnIndex": start_col,
            "endColumnIndex":   end_col,
        }

    series = []

    # Leader COLUMN series (stacked)
    for i in range(N):
        leader_col = 1 + i

        series.append({
            "series":     {"sourceRange": {"sources": [src(leader_col, leader_col + 1)]}},
            "targetAxis": "LEFT_AXIS",
            "type":       "COLUMN",
            "dataLabel": {
                "type":      "DATA",
                "placement": "CENTER",
            },
        })

    # Upper limit — LINE series
    series.append({
        "series":     {"sourceRange": {"sources": [src(N + 1, N + 2)]}},
        "targetAxis": "LEFT_AXIS",
        "type":       "LINE",
    })

    # Lower limit — LINE series
    series.append({
        "series":     {"sourceRange": {"sources": [src(N + 2, N + 3)]}},
        "targetAxis": "LEFT_AXIS",
        "type":       "LINE",
    })

    return {
        "addChart": {
            "chart": {
                "spec": {
                    "title": (f"{month_label(month)} | Avg RPS: Rs{monthly_avg} | "
                              f"Upper: Rs{upper} | Lower: Rs{lower}"),
                    "basicChart": {
                        "chartType":     "COMBO",
                        "stackedType":   "STACKED",
                        "legendPosition": "BOTTOM_LEGEND",
                        "axis": [
                            {"position": "BOTTOM_AXIS", "title": "Date"},
                            {"position": "LEFT_AXIS",   "title": "Avg Rev/Session (Rs)"},
                        ],
                        "series":  series,
                        "domains": [{"domain": {
                            "sourceRange": {"sources": [src(0, 1)]}
                        }}],
                        "headerCount": 1,
                    },
                },
                "position": {
                    "overlayPosition": {
                        "anchorCell": {
                            "sheetId":     sheet_id,
                            "rowIndex":    chart_anchor_row,
                            "columnIndex": chart_anchor_col,
                        },
                        "widthPixels":  950,
                        "heightPixels": 500,
                    }
                },
            }
        }
    }


# ─── Conditional formatting (green > upper, yellow < lower) ─────────────────

def cf_request(sheet_id, data_row_start, data_row_end, rps_col, upper, lower):
    """Colour the Day Avg RPS column: green if > upper, yellow if < lower."""
    base = {
        "sheetId":          sheet_id,
        "startColumnIndex": rps_col,
        "endColumnIndex":   rps_col + 1,
        "startRowIndex":    data_row_start,
        "endRowIndex":      data_row_end,
    }
    return [
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [base],
                    "booleanRule": {
                        "condition": {
                            "type": "NUMBER_GREATER",
                            "values": [{"userEnteredValue": str(upper)}]
                        },
                        "format": {"backgroundColor": {"red": 0.56, "green": 0.93, "blue": 0.56}},
                    }
                },
                "index": 0,
            }
        },
        {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [base],
                    "booleanRule": {
                        "condition": {
                            "type": "NUMBER_LESS",
                            "values": [{"userEnteredValue": str(lower)}]
                        },
                        "format": {"backgroundColor": {"red": 1.0, "green": 0.95, "blue": 0.4}},
                    }
                },
                "index": 1,
            }
        },
    ]


# ─── Bold / colour month title row ──────────────────────────────────────────

def title_format_request(sheet_id, row0, num_cols):
    return [{
        "repeatCell": {
            "range": {
                "sheetId":          sheet_id,
                "startRowIndex":    row0,
                "endRowIndex":      row0 + 1,
                "startColumnIndex": 0,
                "endColumnIndex":   num_cols,
            },
            "cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 0.23, "green": 0.47, "blue": 0.85},
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                               "fontSize": 11},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }
    }]


def header_format_request(sheet_id, row0, num_cols):
    return [{
        "repeatCell": {
            "range": {
                "sheetId":          sheet_id,
                "startRowIndex":    row0,
                "endRowIndex":      row0 + 1,
                "startColumnIndex": 0,
                "endColumnIndex":   num_cols,
            },
            "cell": {"userEnteredFormat": {
                "backgroundColor": {"red": 0.85, "green": 0.90, "blue": 0.98},
                "textFormat": {"bold": True, "fontSize": 9},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)",
        }
    }]


# ─── Main ────────────────────────────────────────────────────────────────────

def main():
    print("Building service ...")
    svc = build_service()

    print("Fetching source data ...")
    records = fetch_records(svc)
    print(f"  {len(records)} usable records (col H != 'N')")

    pivots = build_month_pivots(records)
    months = sorted(pivots)
    print(f"  {len(months)} months: {', '.join(month_label(m) for m in months)}")

    print(f"Ensuring '{OUTPUT_SHEET}' sheet ...")
    sheet_id = ensure_sheet(svc, OUTPUT_SHEET, needed_rows=5000, needed_cols=40)

    print("Clearing existing content & charts ...")
    clear_sheet(svc, OUTPUT_SHEET)
    delete_charts(svc, sheet_id)

    # ── Write title row ─────────────────────────────────────────────────────
    write_range(svc, "A1", [[
        "Stack Charts: Leader Revenue Contribution to Daily Avg Session Revenue",
        "", "", "Filter: Column H != N", "",
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}",
    ]])

    # ── Per-month: write data, queue chart & format requests ────────────────
    all_chart_reqs  = []
    all_format_reqs = []

    current_row = 2          # 0-based; row 1 = row-index 1
    CHART_COL   = 22         # Column W (0-based = 22) — charts anchored here

    for month in months:
        mpivot = pivots[month]
        header, data_rows, label_rows, leader_names = month_rows(month, mpivot)

        N         = len(leader_names)   # number of leader + Others cols
        num_cols  = len(header)

        title_row  = current_row          # month title
        header_row = current_row + 1
        data_start = current_row + 2
        data_end   = data_start + len(data_rows)   # exclusive

        # ── Write month title ──────────────────────────────────────────────
        write_range(svc, f"A{title_row + 1}", [[
            f"{month_label(month)}",
            f"Monthly Avg RPS: Rs{mpivot['monthly_avg_rps']}",
            f"Upper (mean+1σ): Rs{mpivot['upper']}",
            f"Lower (mean-1σ): Rs{mpivot['lower']}",
            f"Days: {len(data_rows)}",
            f"Leaders: {', '.join(leader_names)}",
        ]])

        # ── Write header ───────────────────────────────────────────────────
        write_range(svc, f"A{header_row + 1}", [header])

        # ── Write data rows ────────────────────────────────────────────────
        if data_rows:
            write_range(svc, f"A{data_start + 1}", data_rows)

        print(f"  {month_label(month)}: {len(data_rows)} days, {N} leader cols, "
              f"rows {data_start+1}–{data_end}")

        # ── Queue chart ────────────────────────────────────────────────────
        if len(data_rows) > 1:
            all_chart_reqs.append(build_chart_request(
                sheet_id       = sheet_id,
                month          = month,
                monthly_avg    = mpivot["monthly_avg_rps"],
                upper          = mpivot["upper"],
                lower          = mpivot["lower"],
                header_row     = header_row,
                data_row_start = data_start,
                data_row_end   = data_end,
                chart_anchor_row  = title_row,
                chart_anchor_col  = CHART_COL,
                num_leader_cols   = N,
            ))

        # ── Queue conditional formatting (Day Avg RPS col = L = col index N+3) ──
        rps_col = N + 3   # 0-based col index of "Day Avg RPS"
        all_format_reqs += cf_request(
            sheet_id, data_start, data_end,
            rps_col, mpivot["upper"], mpivot["lower"]
        )
        all_format_reqs += title_format_request(sheet_id, title_row, 6)
        all_format_reqs += header_format_request(sheet_id, header_row, num_cols)

        current_row = data_end + 3   # 3 blank rows between months

    # ── Send all charts in one batchUpdate ──────────────────────────────────
    print(f"Creating {len(all_chart_reqs)} charts ...")
    CHUNK = 5
    for i in range(0, len(all_chart_reqs), CHUNK):
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": all_chart_reqs[i:i+CHUNK]}
        ).execute()
        print(f"  Charts {i+1}–{min(i+CHUNK, len(all_chart_reqs))} created.")

    # ── Send formatting ──────────────────────────────────────────────────────
    print(f"Applying {len(all_format_reqs)} formatting rules ...")
    CHUNK = 20
    for i in range(0, len(all_format_reqs), CHUNK):
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": all_format_reqs[i:i+CHUNK]}
        ).execute()

    print("\nDone.")
    print(f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit")


if __name__ == "__main__":
    main()
