# Hedge fund analytics

A standalone desktop tool for analysing monthly hedge fund returns from an
Excel workbook. It replaces three separate scripts with one, driven by a single
settings window, and adds a workbook format check that runs before any chart is
produced.

| Module | What it produces | Layout it needs |
| --- | --- | --- |
| 1 - Cone charts | Expected-vs-actual excess return cone, rolling total return with rolling volatility, rolling Sharpe, drawdown - plus a compiled PDF | Cone layout only |
| 2 - Correlation & clustering | Correlation heatmaps, Ward dendrograms, cluster assignments, exclusion audit, per-lookback PDFs | Either layout |
| 3 - Performance & drawdowns | NAV and drawdown lines (interactive HTML or static), rolling drawdown / performance / risk-adjusted-return heatmaps, combined PDF | Either layout |

Any combination of modules can run together, as long as the sheet's layout
supports each one.

## Install and run

```
pip install -r requirements.txt
python hedge_fund_analytics.py
```

Running it with no arguments opens the settings window. Every option lives in
that window, and the choices are saved to `hf_analytics_settings.json` beside
the script, then reloaded next time. If the script's folder is read-only the
file falls back to the per-user configuration folder, and the window shows
which file is in use.

Command-line use skips the window and takes the rest of its settings from the
saved file:

```
python hedge_fund_analytics.py --check-only "C:\data\returns.xlsx"
python hedge_fund_analytics.py "C:\data\returns.xlsx" --sheet "Raw Data" \
    --modules clustering,performance --output-dir "C:\data\output"
```

Command-line runs do not overwrite the saved settings, so they cannot change
what the window opens with next time.

## Workbook layouts

Both layouts put month-end dates in column A, strategy names on row 1 and fund
names on row 2, with one fund per column from column B.

**Cone layout** adds two header rows, so returns start on Excel row 5:

| | A | B | C | D |
| --- | --- | --- | --- | --- |
| 1 | | Equity L/S | Equity L/S | Risk Free Rate |
| 2 | | Alpha Capital | Bravo Partners | SONIA |
| 3 | | 5.0% | 4.0% | |
| 4 | | 8.0% | 10.0% | |
| 5 | 31/01/2020 | 1.20% | 0.80% | 0.35% |

Row 3 is the annualised expected **excess** return and row 4 the annualised
expected volatility.

**Plain layout** has no expected figures, so returns start on Excel row 3.

Other rules:

* Returns must be entered as Excel percentages - format the cells as `%` so
  that 5% is stored underneath as 0.05. Text such as `5%` is also accepted. A
  plain `5` means 500% and is rejected as a units error.
* A column whose row 1 starts with `Risk Free` is the risk-free series. It can
  sit anywhere from column B onwards. Module 1 uses it; modules 2 and 3 exclude
  it, because it is not a fund.
* Blanks before a fund launches or after it closes are expected. Blanks inside
  a fund's live history are a data error.

## The format check

The check runs first and reports every problem at once, rather than stopping at
the first one. It refuses to run on duplicate months, unreadable dates,
non-numeric cells, missing strategy or fund names, duplicate fund names inside
a strategy, returns of -100% or worse, returns large enough to imply the wrong
units, and funds whose return never changes.

Two of the checks can be relaxed in the window, for workbooks that are known to
be incomplete:

* **Stop if a month is missing from the date column** - when unticked, absent
  months are inserted as missing data instead.
* **Stop if a fund has blank months inside its live history** - when unticked,
  those months are left as missing data.

With either relaxed, the affected months are reported in the run report, and
any rolling window covering them produces no value rather than a misleading
one.

Funds left out of module 2 or module 3 for having too short a track record are
not errors. They are recorded with their reason in
`excluded_funds_audit.csv` and on the excluded-funds page of module 3's PDF.

## Output

Each run writes to a date-and-time sub-folder of the output folder (this can be
switched off), with one sub-folder per module and a `run_report.txt` recording
the workbook summary, what each module produced and the exact settings used.
