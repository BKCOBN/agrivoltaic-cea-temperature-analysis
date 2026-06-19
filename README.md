# agrivoltaic-cea-temperature-analysis
Scripts for cleaning and processing temperature sensor data from a rack greenhouse experiment. There are two scripts: one that converts raw TXT log files into a structured Excel workbook for each month, and one that reads those workbooks and generates summary charts.

---

## Requirements

Python 3 and openpyxl. Install the library with:

```
pip install openpyxl
```

---

## process_month.py

Combines daily TXT sensor log files for a single month into one formatted Excel workbook. It strips repeated column headers that the logger writes at the start of each 30-minute block, extracts sensor columns T01 through T06, inserts per-row average formulas, and adds a month-wide summary row at the bottom.

### Setup

Open `process_month.py` in any text editor. At the very top you'll see two lines. Update these to match your folder and the desired output name before running:

```python
INPUT_FOLDER = r"USB0\2025-11"           # folder containing the TXT files for the month
OUTPUT_FILE  = r"November_generated.xlsx" # where to save the output
```

- `INPUT_FOLDER` should point to the folder holding the `.txt` files for one month. Use a raw string (`r"..."`) so Windows backslashes are handled correctly.
- `OUTPUT_FILE` is the name (or full path) for the output workbook. If you give just a filename it will be created in the same folder as the script.

### Running

Open a terminal in the folder that contains `process_month.py` and run:

```
py -3 process_month.py
```

You should see output like:

```
Reading TXT files from: C:\...\USB0\2025-11
Loaded 2881 data rows  (2025-11-01 -> 2025-11-30)
Writing: C:\...\November_generated.xlsx
Done.
```

The script will exit with a clear error if the input folder can't be found or contains no valid data.

### Processing a different month

Update the two config lines and re-run:

```python
INPUT_FOLDER = r"USB0\2025-10"
OUTPUT_FILE  = r"October_generated.xlsx"
```

### Input format

- The input folder must contain files named `RDL_YYYY-MM-DD_USB0.txt`, one per day.
- Files are processed in alphabetical order, which matches chronological order given the naming convention.
- Repeated column-header rows are automatically skipped.
- Sensor columns T07 and T08 are excluded because they frequently contain invalid readings.

### Output format

The output is a single `.xlsx` workbook with one sheet named after the month (e.g. `November`):

| Column | Content |
|--------|---------|
| A | Date (`mm-dd-yy`) |
| B | Time of reading (`h:mm:ss`) |
| C | T01 |
| D | T02 |
| E | `=AVERAGE(C:D)` — Under Modules inside Pots |
| F | T03 |
| G | T04 |
| H | `=AVERAGE(F:G)` — Under Modules |
| I | T05 |
| J | T06 |
| K | `=AVERAGE(I:J)` — Control |

- **Row 1** — header row. Columns A and B show the timestamp of the first log entry; C through K show sensor and group labels.
- **Rows 2 to N** — one row per sensor reading.
- **Last row** — month-wide column averages, with no date or time values.

---

## plot_charts.py

Reads the three `*_generated.xlsx` workbooks produced by `process_month.py` and writes a single `Summary_Charts.xlsx` containing all charts.

### What it produces

`Summary_Charts.xlsx` has four sheets:

| Sheet | Content |
|-------|---------|
| **Summary** | Monthly average temperatures per treatment group + bar chart |
| **Oct Day** | Representative-day readings for October + line chart |
| **Nov Day** | Representative-day readings for November + line chart |
| **Dec Day** | Representative-day readings for December + line chart |

The three treatment groups are Under Modules inside Pots, Under Modules, and Control, corresponding to sensor pairs T01/T02, T03/T04, and T05/T06 respectively.

### Representative day

Rather than plotting a full month of data, `plot_charts.py` picks one representative day per month. This is the calendar date whose average temperatures across all three treatment groups are closest to the monthly averages. Proximity is measured as Euclidean distance across the three group averages. That day's readings become the x-axis of each line chart.

### Setup

Open `plot_charts.py` and confirm the `MONTHS` list at the top matches your generated files:

```python
MONTHS = [
    ("October_generated.xlsx",  "October"),
    ("November_generated.xlsx", "November"),
    ("December_generated.xlsx", "December"),
]
```

Each entry is a `(filename, display_name)` pair. Add or remove entries if you have more or fewer months. The filenames must match what `process_month.py` produced.

### Running

Open a terminal in the folder that contains `plot_charts.py` and the `*_generated.xlsx` files, then run:

```
py -3 plot_charts.py
```

You should see output like:

```
Reading October_generated.xlsx ...
  -> representative day: 2025-10-24  (96 readings)
Reading November_generated.xlsx ...
  -> representative day: 2025-11-24  (96 readings)
Reading December_generated.xlsx ...
  -> representative day: 2025-12-20  (96 readings)
Writing Summary_Charts.xlsx ...
Done.
```

The script skips any month whose file is missing and exits with an error if no months can be processed at all. Re-running overwrites `Summary_Charts.xlsx` from scratch.

### Manual formatting applied after running

The version used in the paper had the following changes applied by hand in Excel after the script ran:

- **Chart style 13** applied to all four charts (the monthly averages bar chart and all three representative-day line charts). Select each chart, go to Chart Design > Chart Styles, and pick style 13.
- **Charts resized** slightly for better readability in the paper layout.
- **Horizontal axis crosses at minimum value** set on all four charts. Right-click the value (Y) axis, Format Axis, and under Axis Options set "Horizontal axis crosses" to the minimum axis value rather than automatic.
