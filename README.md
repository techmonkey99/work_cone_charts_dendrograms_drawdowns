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
| 4 - Portfolio risk allocation | Risk budget by cluster, the dendrogram with weight and risk in the margin, an add-or-trim chart, concentration measures, per-lookback PDFs and CSVs | Either layout, plus a weights row |

Any combination of modules can run together, as long as the sheet's layout
supports each one.

## Install and run

```
pip install -r requirements.txt
python hedge_fund_analytics.py
```

Running it with no arguments opens the settings window. Every option lives in
that window, and the choices are saved to `hf_analytics_settings.json` beside
the script, then reloaded next time.

Once a workbook is chosen, the **Worksheet** box becomes a drop-down list of
that workbook's tabs, so the name never has to be typed correctly. It keeps the
tab you used last if the new workbook still has it, and otherwise selects the
first tab. If the workbook cannot be opened - it is missing, or locked by Excel
- the box stays typeable and says why. If the script's folder is read-only the
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

**The weights row** is optional and works with either layout. Put `Weight` in
column A of the row directly above the first month, and each fund's share of
the portfolio in its own column. Everything below moves down one row, so the
plain layout then starts on row 4 and the cone layout on row 6:

| | A | B | C | D |
| --- | --- | --- | --- | --- |
| 1 | | Equity L/S | Equity L/S | Risk Free Rate |
| 2 | | Alpha Capital | Bravo Partners | SONIA |
| 3 | | 5.0% | 4.0% | |
| 4 | | 8.0% | 10.0% | |
| 5 | Weight | 6.0% | 4.0% | |
| 6 | 31/01/2020 | 1.20% | 0.80% | 0.35% |

Only module 4 reads it; the other three ignore it. Weights may be typed as
percentages or as whole numbers, since their total tells the two apart, and the
run report says which was assumed. A weight left blank means the fund is not
held, which is how a candidate is measured before it is bought - those funds
still appear on the add-or-trim chart. Weights totalling less than 100% leave
the rest in cash, which carries no risk; more than 100% is read as gearing.
Leave the risk-free column's weight blank.

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

The check runs automatically as soon as you pick a workbook, and again whenever
you change the tab, the modules or either data check, so the panel at the
bottom of the window always reflects what a run would do. It reports every
problem at once, rather than stopping at the first one. Automatic checks only
write to that panel; the buttons are there to re-check on demand, after editing
the workbook in Excel. It refuses to run on duplicate months, unreadable dates,
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

## Module 4 - portfolio risk allocation

Weight is not risk. Three managed-futures funds at 5% each are a smaller block
of risk than their 15% of the book suggests, because they diversify one
another; three multi-strategy funds at 5% each are a larger one. Module 4
splits the portfolio's volatility across the clusters module 2 finds, so a
cluster on these charts is the same cluster on that dendrogram, and it uses
module 2's lookbacks, clustering method and cluster count.

It produces four charts per lookback:

* **Cluster risk budget** - each cluster's share of the capital against its
  share of the risk, with the funds inside it, and a multiplier saying how much
  risk it carries per unit of capital.
* **The dendrogram with weight and risk in the margin** - the clustering
  picture, with what the portfolio actually owns beside it.
* **Add or trim** - each fund's marginal contribution to portfolio volatility
  against its expected return. The diagonal is the portfolio's own return per
  unit of risk: above it, the next pound improves that ratio; below it,
  trimming does. Funds with no weight are drawn as open rings, so a candidate
  can be judged before it is bought.
* **Concentration** - portfolio volatility, the effective number of bets, the
  diversification ratio, and every fund's gap between its share of the risk and
  its share of the capital.

The decomposition is the standard one: with weights `w` and covariance `S`, the
portfolio's volatility is `sqrt(w'Sw)`, fund *i*'s marginal contribution is
`(Sw)i / vol` and its component contribution is `wi` times that. Those
components add up to the portfolio's volatility exactly, so they group by
cluster with nothing left over.

Two settings change the risk numbers, and both are named in the chart footers
and the run report:

* **Covariance estimator** - Ledoit-Wolf shrinkage by default, because sixty
  months of twenty-odd funds is enough to estimate a covariance but not enough
  to trust every entry of it. `sample` gives the plain sample covariance.
* **Reverse return smoothing** - off by default. Monthly marks are stale,
  especially in credit and event, which understates volatility and correlation.
  Turning it on applies a Geltner adjustment to funds with positive
  autocorrelation, which costs the window its first month.

## Output

Each run writes to a date-and-time sub-folder of the output folder (this can be
switched off), with one sub-folder per module and a `run_report.txt` recording
the workbook summary, what each module produced and the exact settings used.
