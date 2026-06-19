# plot_charts.py
from datetime import datetime, date, time
from pathlib import Path
import math
from collections import defaultdict
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference, Series

MONTHS = [
    ("October_generated.xlsx",  "October"),
    ("November_generated.xlsx", "November"),
    ("December_generated.xlsx", "December"),
]
COMBINED_OUTPUT = "Summary_Charts.xlsx"

SERIES_LABELS = ["Under Modules inside Pots", "Under Modules", "Control"]
SERIES_COLORS = ["4472C4", "ED7D31", "70AD47"]  # blue, orange, green


def _avg(a, b):
    if a is None and b is None:
        return None
    if a is None:
        return b
    if b is None:
        return a
    return (a + b) / 2


def _safe_mean(vals):
    clean = [v for v in vals if v is not None]
    return sum(clean) / len(clean) if clean else 0.0


def _coerce_time(val):
    # Normalise various types openpyxl can return for a time cell
    if val is None:
        return None
    if isinstance(val, time):
        return val
    if isinstance(val, datetime):
        return val.time()
    if isinstance(val, (int, float)) and 0.0 <= val < 1.0:
        total_sec = round(val * 86400)
        h, rem = divmod(total_sec, 3600)
        m, s = divmod(rem, 60)
        return time(h % 24, m, s)
    return None


def read_month_data(path):
    # Returns a list of (date, time, g1_avg, g2_avg, g3_avg) from *_generated.xlsx
    # Reads raw sensor columns C+D, F+G, I+J, computes averages in Python
    # Skips header row + the summary row (col A is None on the summary row)
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 10:
            continue
        a = row[0]    # date
        b = row[1]    # time
        c = row[2]    # T01
        d = row[3]    # T02
        # row[4] = col E (AVERAGE formula) — skip
        f = row[5]    # T03
        g = row[6]    # T04
        # row[7] = col H (AVERAGE formula) — skip
        i_val = row[8]  # T05
        j_val = row[9]  # T06

        if a is None:   # summary row has no date
            continue

        if isinstance(a, datetime):
            row_date = a.date()
        elif isinstance(a, date):
            row_date = a
        else:
            continue

        row_time = _coerce_time(b)
        if row_time is None:
            continue

        rows.append((row_date, row_time, _avg(c, d), _avg(f, g), _avg(i_val, j_val)))
    return rows


def find_representative_day(rows):
    # Returns the date whose daily group averages are closest to the monthly averages
    # Uses Euclidean distance across the three group averages
    # Monthly averages are weighted by reading count, so days with more readings carry more influence
    if not rows:
        return None

    monthly = (
        _safe_mean([r[2] for r in rows]),
        _safe_mean([r[3] for r in rows]),
        _safe_mean([r[4] for r in rows]),
    )

    by_date = defaultdict(list)
    for row in rows:
        by_date[row[0]].append(row)

    best_date, best_dist = None, float('inf')
    for d, day_rows in by_date.items():
        day_avg = (
            _safe_mean([r[2] for r in day_rows]),
            _safe_mean([r[3] for r in day_rows]),
            _safe_mean([r[4] for r in day_rows]),
        )
        dist = math.sqrt(sum((a - b) ** 2 for a, b in zip(day_avg, monthly)))
        if dist < best_dist:
            best_dist, best_date = dist, d

    return best_date


def write_chartdata_sheet(wb, sheet_name, rep_rows):
    # Writes representative-day group averages to a supporting data sheet
    # Times are stored as "HH:MM:SS" strings so they render correctly as chart labels
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    ws.append(["Time"] + SERIES_LABELS)
    for _, t, g1, g2, g3 in rep_rows:
        ws.append([t.strftime("%H:%M:%S"), g1, g2, g3])
    return ws


def _make_bar_chart(ws, n_months):
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Monthly Average Temperatures by Treatment"
    chart.y_axis.title = "Temperature (°C)"
    chart.x_axis.title = "Month"
    chart.width = 22
    chart.height = 14
    for col_idx, color in enumerate(SERIES_COLORS, start=2):
        data = Reference(ws, min_col=col_idx, min_row=1, max_row=n_months + 1)
        series = Series(data, title_from_data=True)
        series.graphicalProperties.solidFill = color
        chart.series.append(series)
    cats = Reference(ws, min_col=1, min_row=2, max_row=n_months + 1)
    chart.set_categories(cats)
    return chart


def _make_line_chart(ws, title):
    chart = LineChart()
    chart.style = 10
    chart.title = title
    chart.y_axis.title = "Temperature (°C)"
    chart.x_axis.title = "Time of Day"
    chart.width = 22
    chart.height = 14
    n = ws.max_row
    for col_idx, color in enumerate(SERIES_COLORS, start=2):
        data = Reference(ws, min_col=col_idx, min_row=1, max_row=n)
        series = Series(data, title_from_data=True)
        series.graphicalProperties.line.solidFill = color
        series.graphicalProperties.line.width = 20000
        chart.series.append(series)
    cats = Reference(ws, min_col=1, min_row=2, max_row=n)
    chart.set_categories(cats)
    return chart


_DAY_SHEET_LABELS = {
    "October": "Oct Day",
    "November": "Nov Day",
    "December": "Dec Day",
}


def write_combined_workbook(month_data, output_path):
    # Writes Summary_Charts.xlsx w/ a bar chart and per-month line charts
    wb = openpyxl.Workbook()

    # Summary sheet: monthly averages table + bar chart
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.append(["Month"] + SERIES_LABELS)
    for m in month_data:
        ws_summary.append([m["name"]] + list(m["monthly_avgs"]))
    ws_summary.add_chart(_make_bar_chart(ws_summary, len(month_data)), "F2")

    # Per-month day sheets: daily data + line chart
    for m in month_data:
        sheet_label = _DAY_SHEET_LABELS.get(m["name"], f"{m['name']} Day")
        ws_day = write_chartdata_sheet(wb, sheet_label, m["rep_rows"])
        ws_day.add_chart(
            _make_line_chart(ws_day, f"{m['name']} - Representative Day"), "F2"
        )

    wb.save(output_path)


def main():
    # Must be run from the directory containing *_generated.xlsx files
    month_data = []
    for xlsx_path, month_name in MONTHS:
        p = Path(xlsx_path)
        if not p.exists():
            print(f"Warning: {p} not found - skipping {month_name}")
            continue
        print(f"Reading {p} ...")
        rows = read_month_data(xlsx_path)
        if not rows:
            print(f"Warning: no data rows in {p} - skipping {month_name}")
            continue
        rep_day = find_representative_day(rows)
        rep_rows = [r for r in rows if r[0] == rep_day]
        monthly_avgs = (
            _safe_mean([r[2] for r in rows]),
            _safe_mean([r[3] for r in rows]),
            _safe_mean([r[4] for r in rows]),
        )
        month_data.append({
            "name": month_name,
            "monthly_avgs": monthly_avgs,
            "rep_rows": rep_rows,
        })
        print(f"  -> representative day: {rep_day}  ({len(rep_rows)} readings)")

    if not month_data:
        raise SystemExit("Error: no months could be processed.")

    print(f"Writing {COMBINED_OUTPUT} ...")
    write_combined_workbook(month_data, COMBINED_OUTPUT)
    print("Done.")


if __name__ == "__main__":
    main()
