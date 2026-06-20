from datetime import datetime, date, time
from pathlib import Path
import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference, Series

# update these paths if your filenames differ
MONTHS = [
    ("October_generated.xlsx",  "October"),
    ("November_generated.xlsx", "November"),
    ("December_generated.xlsx", "December"),
]
OUT_FILE = "Summary_Charts.xlsx"

LABELS = ["Under Modules inside Pots", "Under Modules", "Control"]
COLORS = ["4472C4", "ED7D31", "70AD47"]  # blue, orange, green

DAY_TABS = {
    "October": "Oct Day",
    "November": "Nov Day",
    "December": "Dec Day",
}

DT_TABS = {
    "October": "Oct High dT",
    "November": "Nov High dT",
    "December": "Dec High dT",
}


def avg(vals):
    # skip None vals
    total = 0.0
    count = 0
    for v in vals:
        if v is not None:
            total += v
            count += 1
    if count == 0: return 0.0
    return total / count


def read_xlsx(path):
    # skip E/H/K (formula cols), recompute avgs from raw pairs
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if len(row) < 10 or row[0] == None:
            continue
        dval = row[0]
        rdate = None
        if isinstance(dval, datetime):
            rdate = dval.date()
        elif isinstance(dval, date):
            rdate = dval
        if rdate == None:
            continue
        tv = row[1]
        if isinstance(tv, time):
            rtime = tv
        elif isinstance(tv, datetime):
            rtime = tv.time()
        elif isinstance(tv, float):
            # time stored as fraction of day (0.5 = noon)
            sec = round(tv * 86400)
            rtime = time(sec // 3600 % 24, sec % 3600 // 60, sec % 60)
        else:
            continue

        t01, t02 = row[2], row[3]
        if t01 is not None and t02 is not None:
            pot = (t01 + t02) / 2
        else:
            pot = t01 if t01 is not None else t02

        t03, t04 = row[5], row[6]
        if t03 is not None and t04 is not None:
            air = (t03 + t04) / 2
        else:
            air = t03 if t03 is not None else t04

        t05, t06 = row[8], row[9]
        if t05 is not None and t06 is not None:
            ctrl = (t05 + t06) / 2
        else:
            ctrl = t05 if t05 is not None else t06

        rows.append((rdate, rtime, pot, air, ctrl))
    return rows


def rep_day(data):
    # which day is most "average" for the month
    if not data:
        return None
    pot = []
    air = []
    ctrl = []
    for r in data:
        pot.append(r[2])
        air.append(r[3])
        ctrl.append(r[4])
    mnthly = (avg(pot), avg(air), avg(ctrl))

    by_day = {}
    for r in data:
        if r[0] not in by_day:
            by_day[r[0]] = []
        by_day[r[0]].append(r)

    rdate = None
    mdist = None
    for d, day in by_day.items():
        pv = []
        av = []
        cv = []
        for r in day:
            pv.append(r[2])
            av.append(r[3])
            cv.append(r[4])
        dmeans = (avg(pv), avg(av), avg(cv))
        # euclidean dist from monthly avg across all 3 groups
        dsq = 0.0
        dsq += (dmeans[0] - mnthly[0]) ** 2
        dsq += (dmeans[1] - mnthly[1]) ** 2
        dsq += (dmeans[2] - mnthly[2]) ** 2
        dist = dsq ** 0.5
        if mdist is None or dist < mdist:
            mdist = dist
            rdate = d
    return rdate


def peak_dt_day(data):
    # find the day where pots ran hottest relative to control
    if not data:
        return None
    by_day = {}
    for r in data:
        if r[0] not in by_day:
            by_day[r[0]] = []
        by_day[r[0]].append(r)
    pdate = None
    mdt = None
    for d, day in by_day.items():
        pot = []
        ctrl = []
        for r in day:
            pot.append(r[2])
            ctrl.append(r[4])
        dt = avg(pot) - avg(ctrl)
        if mdt is None or dt > mdt:
            mdt = dt
            pdate = d
    return pdate


def day_sheet(wb, name, day_rows):
    # format time as string or Excel treats it as a number
    if name in wb.sheetnames:
        del wb[name]
    ws = wb.create_sheet(name)
    ws.append(["Time"] + LABELS)
    for _, t, g1, g2, g3 in day_rows:
        ws.append([t.strftime("%H:%M:%S"), g1, g2, g3])
    return ws


def bar_chart(ws, n):
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Monthly Average Temperatures by Treatment"
    chart.y_axis.title = "Temperature (°C)"
    chart.x_axis.title = "Month"
    chart.width = 22
    chart.height = 14
    col_idx = 2
    for color in COLORS:  # one series per treatment group
        data = Reference(ws, min_col=col_idx, min_row=1, max_row=n + 1)
        series = Series(data, title_from_data=True)
        series.graphicalProperties.solidFill = color
        chart.series.append(series)
        col_idx += 1
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n + 1))
    return chart


def line_chart(ws, title):
    chart = LineChart()
    chart.style = 10
    chart.title = title
    chart.y_axis.title = "Temperature (°C)"
    chart.x_axis.title = "Time of Day"
    chart.width = 22
    chart.height = 14
    n = ws.max_row
    col_idx = 2
    for color in COLORS:
        data = Reference(ws, min_col=col_idx, min_row=1, max_row=n)
        series = Series(data, title_from_data=True)
        series.graphicalProperties.line.solidFill = color
        series.graphicalProperties.line.width = 20000
        chart.series.append(series)
        col_idx += 1
    chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=n))
    return chart


def write_workbook(months, out):
    wb = openpyxl.Workbook()
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.append(["Month"] + LABELS)
    for m in months:
        ws_sum.append([m["name"], m["means"][0], m["means"][1], m["means"][2]])
    ws_sum.add_chart(bar_chart(ws_sum, len(months)), "F2")

    for m in months:
        tab = DAY_TABS.get(m["name"], "%s Day" % m["name"])
        ws = day_sheet(wb, tab, m["rep_rows"])
        ws.add_chart(line_chart(ws, "%s - Representative Day (%s)" % (m["name"], m["rep_date"])), "F2")
    for m in months:
        tab = DT_TABS.get(m["name"], "%s High dT" % m["name"])
        ws = day_sheet(wb, tab, m["dt_rows"])
        ws.add_chart(line_chart(ws, "%s - Highest ΔT Day (%s)" % (m["name"], m["peak"])), "F2")
    wb.save(out)


def main():
    months = []
    for fpath, mname in MONTHS:
        p = Path(fpath)
        if not p.exists():
            print("%s not found, skipping" % p)
            continue
        print("Reading %s ..." % p)
        data = read_xlsx(fpath)
        if not data:
            print("no data in %s, skipping" % p)
            continue

        rdate = rep_day(data)
        peak = peak_dt_day(data)

        # collect means + pull out readings for both special days in one pass
        pot = []
        air = []
        ctrl = []
        rep_rows = []
        dt_rows = []
        for r in data:
            pot.append(r[2])
            air.append(r[3])
            ctrl.append(r[4])
            if r[0] == rdate:
                rep_rows.append(r)
            if r[0] == peak:
                dt_rows.append(r)

        months.append({
            "name": mname,
            "means": (avg(pot), avg(air), avg(ctrl)),
            "rep_date": rdate,
            "rep_rows": rep_rows,
            "peak": peak,
            "dt_rows": dt_rows,
        })
        print("  -> rep day: %s (%d readings)" % (rdate, len(rep_rows)))
        print("  -> peak dT: %s (%d readings)" % (peak, len(dt_rows)))

    if not months:
        raise SystemExit("nothing to process, check your xlsx files")
    write_workbook(months, OUT_FILE)
    print("Done.")

if __name__ == "__main__":
    main()
