# agrivoltaic-cea-temperature-analysis
Companion code for *"Low-cost adaptation of agrivoltaic infrastructure for controlled-environment agriculture to extend the growing season in northern latitudes"* by Uzair Jamil, Beril Kayla Coban, and Joshua M. Pearce. Two scripts: one that converts raw TXT log files from the rack greenhouse study into a structured Excel workbook per month, and one that reads those workbooks and generates summary charts.

## Requirements

Python 3 and openpyxl:

```
pip install openpyxl
```

---

## process_month.py

Takes daily TXT sensor log files for a single month and outputs one formatted Excel workbook. Handles filtering of repeated headers, extracts T01–T06, adds per-row average formulas, and puts a summary row at the bottom.

Edit the two lines at the top before running:

```python
INPUT_FOLDER = r"USB0\2025-11"            # folder with the TXT files for the month
OUTPUT_FILE  = r"November_generated.xlsx" # output filename
```

Then run:

```
py -3 process_month.py
```

**Input:** Files named `RDL_YYYY-MM-DD_USB0.txt`, one per day. Boot messages and repeated column headers are filtered automatically. Sensors T07 and T08 are excluded — they produced non-physical readings throughout the study.

**Output:** A single `.xlsx` workbook with one sheet named after the month. Columns A–B are date and time; C–K are sensors T01–T06 with averaged group columns (E, H, K) for Under Modules inside Pots, Under Modules, and Control respectively. The last row holds month-wide averages.

---

## plot_charts.py

Reads the three `*_generated.xlsx` workbooks and writes `Summary_Charts.xlsx` with seven sheets:

| Sheet | Content |
|-------|---------|
| **Summary** | Monthly average temperatures per treatment group + bar chart |
| **Oct Day** | Representative-day readings for October + line chart |
| **Nov Day** | Representative-day readings for November + line chart |
| **Dec Day** | Representative-day readings for December + line chart |
| **Oct High dT** | Highest ΔT day for October + line chart |
| **Nov High dT** | Highest ΔT day for November + line chart |
| **Dec High dT** | Highest ΔT day for December + line chart |

The representative day is the calendar date whose daily group means are closest to the monthly means (Euclidean distance across all three treatment groups) — this preserves the diurnal pattern that monthly averages flatten. The highest ΔT day is the day with the largest mean temperature gap between the under-rack pot and the control.

Run from the folder containing the `*_generated.xlsx` files:

```
py -3 plot_charts.py
```

### Manual formatting applied after running

The version used in the paper had these changes applied by hand in Excel:

- **Chart style 13** on all seven charts (Chart Design > Chart Styles > style 13)
- **Charts resized** slightly for the paper layout
- **Horizontal axis crosses at minimum value** on all charts (right-click Y axis > Format Axis > Horizontal axis crosses > minimum value)
