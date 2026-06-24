# agrivoltaic-cea-temperature-analysis
Companion code for *"Low-cost adaptation of agrivoltaic infrastructure for controlled-environment agriculture to extend the growing season in northern latitudes"* by Uzair Jamil, Beril Kayla Coban, and Joshua M. Pearce.

This repo contains the scripts used to process raw sensor and weather-station data, match them by timestamp, and produce the temperature charts in the paper.

## Requirements

Python 3 and openpyxl (handles both file I/O and chart generation):

```
pip install openpyxl
```

---

## Pipeline overview

```
USB0 raw TXT logs  ──┐
                     ├──► process_met.py ──► *_MET_matched.xlsx ──► plot_charts.py ──► Summary_Charts.xlsx
MET station CSVs   ──┘
```

Run `process_met.py` first to build the matched workbooks, then `plot_charts.py` to generate the charts:
```
py -3 process_met.py
py -3 plot_charts.py
```

---

## process_met.py

Reads raw USB0 sensor logs and MET station CSVs for each month, matches them by timestamp, and writes one `*_MET_matched.xlsx` per month with all four temperature series side by side.

**Configuration** (edit top of file):
```python
USB0_ROOT  = r"USB0"
MET_ROOT   = r"German-Solar"
WINDOW_MIN = 7
MONTHS = [
    ("2025-10", "October-German-Solar",  "October_MET_matched.xlsx"),
    ...
]
```

**Input:**
- USB0 TXT files named `RDL_YYYY-MM-DD_USB0.txt`, one per day, in `USB0\YYYY-MM\`
- MET station CSV files in `German-Solar\<Month>-German-Solar\`

**Output columns per `*_MET_matched.xlsx`:**

| Column | Source | Description |
|--------|--------|-------------|
| A | — | Date |
| B | — | Time |
| C | Data logger T01/T02 avg | Soil Under Rack (°C) |
| D | Data logger T03/T04 avg | Air Under Rack (°C) |
| E | Data logger T05/T06 avg | Soil Outside (°C) |
| F | MET station | MET Ambient (°C) |
| G | =D−F | ΔT Air-Ambient (°C) |

**MET matching:** Each USB0 reading (15-minute intervals) is matched to the mean of all MET station readings (1-minute intervals) within a ±7 minute window. This smooths short-term MET fluctuations around each USB0 timestamp.

**Sensor notes:** T07 and T08 are excluded — they produced non-physical readings throughout the study.

---

## plot_charts.py

Reads the three `*_MET_matched.xlsx` files and writes `Summary_Charts.xlsx` with seven sheets:

| Sheet | Content |
|-------|---------|
| **Summary** | Monthly average temperatures for all four series + bar chart |
| **Oct Day** | Representative-day readings for October + line chart |
| **Nov Day** | Representative-day readings for November + line chart |
| **Dec Day** | Representative-day readings for December + line chart |
| **Oct High dT** | Highest ΔT day for October + line chart |
| **Nov High dT** | Highest ΔT day for November + line chart |
| **Dec High dT** | Highest ΔT day for December + line chart |

### Four temperature series

Every chart tracks the same four temperatures:
1. **Soil Under Rack** — pot temperature under the plastic-covered rack (T01/T02 avg)
2. **Air Under Rack** — air temperature under the plastic-covered rack (T03/T04 avg)
3. **Soil Outside** — pot temperature in the uncovered control (T05/T06 avg)
4. **MET Ambient** — outdoor ambient air temperature from the MET station

### Representative day selection

The representative day is the calendar date whose **mean MET ambient temperature is closest to the monthly mean MET ambient temperature**.

**Rationale:** Ambient temperature is the primary driver of all other series (under-rack air and both soil temps). Anchoring the selection to ambient grounds the representative day in typical outdoor conditions, and when plotted, the diurnal profile shows how all four sensors respond on a normal day. This pairs with the peak ΔT chart to tell two distinct stories: typical conditions vs. best-case thermal benefit.

### Peak ΔT day selection

The peak ΔT day is the calendar date with the **highest mean daily ΔT**, where ΔT = Air Under Rack − MET Ambient.

**Rationale:** The paper's central question is whether enclosing the agrivoltaic rack improves the thermal environment. ΔT between inside air and outside ambient is the direct measure of that benefit. The peak ΔT day shows the rack performing at its best.

### Manual formatting applied after running

The version used in the paper had these adjustments made by hand in Excel:

- **Chart style 13** on all seven charts (Chart Design > Chart Styles > style 13)
- **Charts resized** slightly to fit the paper layout
- **Horizontal axis crosses at minimum value** on all charts (right-click Y axis > Format Axis > Horizontal axis crosses > minimum value)

---

## process_month.py (manual reference only)

Not part of the main pipeline. Converts raw USB0 TXT logs for a single month into a structured Excel workbook — useful for manually inspecting individual sensor readings without digging through raw TXT files.

Edit the two lines at the top before running:

```python
INPUT_FOLDER = r"USB0\2025-11"
OUTPUT_FILE  = r"November_generated.xlsx"
```

**Output:** One `.xlsx` workbook with columns A–K: date, time, T01–T06 with per-row average formulas for each sensor pair (E = avg T01/T02, H = avg T03/T04, K = avg T05/T06), and a summary row at the bottom.

---

## Tests

```
py -3 -m pytest tests/ -v
```

Covers MET CSV parsing, USB0 log parsing (all six sensors), timestamp deduplication, and the window averaging function.
