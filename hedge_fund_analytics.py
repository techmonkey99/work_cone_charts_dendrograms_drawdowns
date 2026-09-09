"""Hedge fund analytics suite - cone charts, correlation clustering and
performance/drawdown analysis, driven from a single settings window.

This file replaces three separate scripts:

    Module 1  Cone charts            expected-vs-actual excess return cones,
                                     rolling total return, rolling Sharpe,
                                     drawdown.
    Module 2  Correlation clustering correlation heatmaps, Ward dendrograms,
                                     cluster assignments, exclusion audit.
    Module 3  Performance/drawdowns  NAV and drawdown lines plus rolling
                                     drawdown, rolling performance and
                                     risk-adjusted-return heatmaps.

===============================================================================
EXPECTED WORKBOOK LAYOUTS
===============================================================================
Two layouts are recognised. Both share the same first two rows and put month
end dates in column A. The script detects which one a sheet uses and reports it
before anything is produced.

CONE LAYOUT (required by module 1; modules 2 and 3 can also read it)

     A                    B                     C                    D
 1   [any label]          Strategy name         Strategy name        Risk Free
 2   [any label]          Fund name             Fund name            [any label]
 3   [any label]          Expected excess ret.  Expected excess ret. [blank]
 4   [any label]          Expected volatility   Expected volatility  [blank]
 5   31/01/2020           1.20%                 0.80%                0.35%
 6   29/02/2020          -0.40%                 0.25%                0.35%
     ...                  ...                   ...                  ...

PLAIN LAYOUT (modules 2 and 3 only)

     A                    B                     C
 1   [any label]          Strategy name         Strategy name
 2   [any label]          Fund name             Fund name
 3   31/01/2020           1.20%                 0.80%
 4   29/02/2020          -0.40%                 0.25%
     ...                  ...                   ...

Rules that apply to both layouts:

* Column A holds one month-end date per row, one row per calendar month.
* Row 1 is the strategy, repeated above every fund in that strategy.
* Row 2 is the fund name, which must be unique inside its strategy.
* Each fund occupies exactly one column from column B onwards.
* A column whose row 1 begins with "Risk Free" (case-insensitive) is the
  risk-free series. It may sit anywhere from column B onwards. Module 1 uses
  it; modules 2 and 3 exclude it, since it is not a fund.
* Returns must be entered as Excel percentages - format the cells as % so that
  5% is stored underneath as 0.05. A cell may also be text such as "5%".
  A plain number 5 means 500% and will be rejected as a units error.
* Blank cells before a fund launches or after it closes are expected and fine.
  Blanks inside a fund's live history are a data error.

===============================================================================
NORMAL USE
===============================================================================
Double-click the script, or run it with no arguments, to open the settings
window. Tick the modules to run, point it at the workbook, adjust anything you
want and press "Check and run". Every setting is saved to

    hf_analytics_settings.json

in the same folder as this script, and reloaded automatically next time. If
that folder is read-only the file falls back to the per-user configuration
folder, and the window shows which file is in use.

Command-line use skips the window:

    python hedge_fund_analytics.py "C:\\data\\returns.xlsx" --sheet "Raw Data" \
        --modules cone,clustering,performance --output-dir "C:\\data\\output"
    python hedge_fund_analytics.py --check-only "C:\\data\\returns.xlsx"

Required packages:
    pip install pandas numpy scipy matplotlib seaborn openpyxl
Optional (interactive HTML line charts in module 3):
    pip install plotly
"""

from __future__ import annotations

import argparse
import datetime as _dt
import importlib.util
import json
import logging
import math
import os
import re
import sys
import textwrap
from contextlib import ExitStack
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, Literal, Sequence

import matplotlib

import numpy as np
import pandas as pd

LOGGER = logging.getLogger("hedge_fund_analytics")

SETTINGS_FILENAME = "hf_analytics_settings.json"
APP_NAME = "HedgeFundAnalytics"

MODULE_KEYS = ("cone", "clustering", "performance")
MODULE_TITLES = {
    "cone": "1 - Cone charts",
    "clustering": "2 - Correlation & clustering",
    "performance": "3 - Performance & drawdowns",
}
MODULE_FOLDERS = {
    "cone": "1_cone_charts",
    "clustering": "2_correlation_clustering",
    "performance": "3_performance_drawdowns",
}

ROW_STRATEGY = 0
ROW_FUNDNAME = 1
ROW_CONE_EXPECTED_EXCESS = 2
ROW_CONE_EXPECTED_VOL = 3
CONE_DATA_START_ROW = 4
PLAIN_DATA_START_ROW = 2

RISK_FREE_PREFIX = "risk free"

# A single month above this size is reported as unusual; above the hard limit it
# is treated as a units error, because it almost always means the cells hold
# whole numbers (5 for 5%) rather than Excel percentages.
UNUSUAL_MONTHLY_RETURN = 0.50
IMPOSSIBLE_MONTHLY_RETURN = 2.00

INITIAL_PRICE_INDEX_VALUE = 100.0

LAYOUT_DESCRIPTIONS = {
    "cone": (
        "Cone layout - row 1 strategy, row 2 fund, row 3 expected excess return, "
        "row 4 expected volatility, monthly returns from row 5."
    ),
    "plain": (
        "Plain layout - row 1 strategy, row 2 fund, monthly returns from row 3."
    ),
}


class WorkbookFormatError(Exception):
    """Raised when the workbook cannot be used as it stands.

    The message lists every problem found, so that one run of the check tells
    the user everything they need to fix.
    """


class UserCancelled(Exception):
    """Raised when the settings window is closed without running."""


# =============================================================================
# Settings: defaults, persistence and merging
# =============================================================================

def default_settings() -> dict[str, Any]:
    """Every user-facing option, grouped by the part of the run it controls."""
    return {
        "workbook": {
            "path": "",
            "sheet": "Raw Data",
            "output_dir": "hedge_fund_outputs",
            "timestamped_run_folder": True,
            "stop_on_missing_months": True,
            "stop_on_history_gaps": True,
        },
        "modules": {"cone": True, "clustering": True, "performance": True},
        "cone": {
            "use_risk_free": True,
            "use_predetermined_months": False,
            "n_months": 60,
            "rolling_window": 24,
            "sharpe_window": 36,
            "volatility_window": 24,
            "drawdown_threshold_pct": 5.0,
            "common_x_axis": False,
            "create_pdf": True,
            "dpi": 300,
        },
        "clustering": {
            "timeframes": [36, 60, "max"],
            "min_required_months": 12,
            "require_latest_return": True,
            "correlation_method": "pearson",
            "linkage_method": "ward",
            "max_clusters": 5,
            "distance_threshold": None,
            "annotate_heatmaps": True,
            "max_annotated_funds": 45,
            "label_wrap_width": 30,
            "all_strategies_label_wrap_width": 45,
            "create_pdf": True,
            "save_png": True,
            "save_csv": True,
            "dpi": 180,
        },
        "performance": {
            "rolling_window_months": 12,
            "min_history_months": 0,
            "line_chart_window_months": 0,
            "recent_months_heatmap": 0,
            "x_label_wrap_width": 20,
            "require_full_rolling_window": True,
            "line_chart_mode": "html",
            "create_pdf": True,
            "save_png": False,
            "dpi": 300,
        },
    }


def settings_path() -> Path:
    """Prefer the script's own folder; fall back when that folder is read-only."""
    try:
        beside_script = Path(__file__).resolve().parent / SETTINGS_FILENAME
    except NameError:  # pragma: no cover - only when exec'd without a file
        beside_script = Path.cwd() / SETTINGS_FILENAME
    try:
        beside_script.parent.mkdir(parents=True, exist_ok=True)
        probe = beside_script.with_suffix(beside_script.suffix + ".probe")
        probe.write_text("", encoding="utf-8")
        probe.unlink()
        return beside_script
    except OSError:
        if os.name == "nt" and os.environ.get("APPDATA"):
            base = Path(os.environ["APPDATA"])
        else:
            base = Path.home() / ".config"
        return base / APP_NAME / SETTINGS_FILENAME


def merge_settings(defaults: dict[str, Any], saved: Any) -> dict[str, Any]:
    """Copy across known keys only, so an old or damaged file cannot break a run."""
    merged = {section: dict(values) for section, values in defaults.items()}
    if not isinstance(saved, dict):
        return merged
    for section, values in merged.items():
        saved_section = saved.get(section)
        if not isinstance(saved_section, dict):
            continue
        for key in values:
            if key in saved_section:
                values[key] = saved_section[key]
    return merged


def load_settings(path: Path) -> dict[str, Any]:
    settings = default_settings()
    if not path.is_file():
        return settings
    try:
        return merge_settings(settings, json.loads(path.read_text(encoding="utf-8")))
    except (OSError, ValueError, json.JSONDecodeError) as exc:
        LOGGER.warning("Could not read settings from %s (%s). Using defaults.", path, exc)
        return settings


def save_settings(path: Path, settings: dict[str, Any]) -> None:
    """Write atomically, so an interrupted save cannot corrupt the file."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(settings, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    temporary.replace(path)


def validate_settings(settings: dict[str, Any]) -> None:
    """Reject impossible combinations before any file is read or written."""
    workbook = settings["workbook"]
    if not str(workbook["path"]).strip():
        raise ValueError("Select an Excel workbook.")
    if not Path(str(workbook["path"])).expanduser().is_file():
        raise ValueError(f"The workbook does not exist:\n{workbook['path']}")
    if not str(workbook["sheet"]).strip():
        raise ValueError("Enter the worksheet name (or a zero-based sheet number).")
    if not str(workbook["output_dir"]).strip():
        raise ValueError("Select an output folder.")
    if not any(settings["modules"][key] for key in MODULE_KEYS):
        raise ValueError("Select at least one module to run.")

    cone = settings["cone"]
    for key, label in (
        ("n_months", "Fixed number of months"),
        ("rolling_window", "Rolling return window"),
        ("sharpe_window", "Sharpe window"),
        ("volatility_window", "Volatility window"),
    ):
        if int(cone[key]) <= 0:
            raise ValueError(f"Cone charts: {label} must be a positive whole number.")
    if not 0 < float(cone["drawdown_threshold_pct"]) < 100:
        raise ValueError("Cone charts: the drawdown threshold must be between 0 and 100%.")
    if int(cone["dpi"]) < 72:
        raise ValueError("Cone charts: image resolution must be at least 72 DPI.")

    clustering = settings["clustering"]
    normalise_timeframes(clustering["timeframes"])
    if int(clustering["min_required_months"]) < 3:
        raise ValueError("Clustering: the minimum usable history must be at least 3 months.")
    if int(clustering["max_clusters"]) < 2:
        raise ValueError("Clustering: the maximum number of clusters must be at least 2.")
    if clustering["correlation_method"] not in {"pearson", "spearman", "kendall"}:
        raise ValueError("Clustering: the correlation method must be pearson, spearman or kendall.")
    if clustering["linkage_method"] not in {"ward", "average", "complete"}:
        raise ValueError("Clustering: the linkage method must be ward, average or complete.")
    if clustering["linkage_method"] == "ward" and clustering["correlation_method"] == "kendall":
        raise ValueError(
            "Clustering: Ward linkage supports Pearson or Spearman. "
            "Choose average or complete linkage to use Kendall."
        )
    if clustering["distance_threshold"] is not None and float(clustering["distance_threshold"]) <= 0:
        raise ValueError("Clustering: the dendrogram colour threshold must be positive, or left blank.")
    if int(clustering["max_annotated_funds"]) < 2:
        raise ValueError("Clustering: the annotation limit must be at least 2 funds.")
    if int(clustering["label_wrap_width"]) < 5 or int(clustering["all_strategies_label_wrap_width"]) < 5:
        raise ValueError("Clustering: label wrap widths must be at least 5 characters.")
    if int(clustering["dpi"]) < 72:
        raise ValueError("Clustering: image resolution must be at least 72 DPI.")
    if not any((clustering["create_pdf"], clustering["save_png"], clustering["save_csv"])):
        raise ValueError("Clustering: select at least one output type (PDF, PNG or CSV).")

    performance = settings["performance"]
    if int(performance["rolling_window_months"]) < 1:
        raise ValueError("Performance: the rolling window must be at least 1 month.")
    for key, label in (
        ("min_history_months", "Minimum fund history"),
        ("line_chart_window_months", "Line-chart history"),
        ("recent_months_heatmap", "Heatmap history"),
    ):
        if int(performance[key]) < 0:
            raise ValueError(f"Performance: {label} cannot be negative.")
    if int(performance["x_label_wrap_width"]) < 6:
        raise ValueError("Performance: the fund-name wrap width must be at least 6 characters.")
    if performance["line_chart_mode"] not in {"html", "static"}:
        raise ValueError("Performance: the line-chart format must be html or static.")
    if int(performance["dpi"]) < 72:
        raise ValueError("Performance: image resolution must be at least 72 DPI.")
    if not any((performance["create_pdf"], performance["save_png"])) and performance["line_chart_mode"] == "static":
        raise ValueError(
            "Performance: static line charts need PDF or PNG output. "
            "Tick one of those, or choose interactive HTML line charts."
        )


# =============================================================================
# Workbook loading, layout detection and format checking
# =============================================================================

@dataclass
class FundColumn:
    """One fund column, already converted to decimal monthly returns."""

    strategy: str
    fund: str
    column: str
    returns: pd.Series
    expected_excess: float | None = None
    expected_vol: float | None = None

    @property
    def label(self) -> str:
        return f"{self.strategy} - {self.fund}"


@dataclass
class WorkbookData:
    """A checked workbook, ready for any of the three modules."""

    path: Path
    sheet: Any
    layout: Literal["cone", "plain"]
    index: pd.DatetimeIndex
    funds: list[FundColumn]
    risk_free: pd.Series | None = None
    risk_free_column: str | None = None
    risk_free_label: str | None = None
    notes: list[str] = field(default_factory=list)

    @property
    def strategies(self) -> list[str]:
        seen: list[str] = []
        for fund in self.funds:
            if fund.strategy not in seen:
                seen.append(fund.strategy)
        return seen

    def wide_frame(self) -> pd.DataFrame:
        """Fund returns as a (strategy, fund) column MultiIndex frame."""
        frame = pd.DataFrame(
            {(fund.strategy, fund.fund): fund.returns for fund in self.funds},
            index=self.index,
        )
        frame.columns = pd.MultiIndex.from_tuples(frame.columns, names=["strategy", "fund"])
        return frame

    def summary_lines(self) -> list[str]:
        lines = [
            f"Workbook: {self.path}",
            f"Worksheet: {self.sheet}",
            f"Detected layout: {LAYOUT_DESCRIPTIONS[self.layout]}",
            f"Months: {len(self.index)} "
            f"({self.index[0]:%b %Y} to {self.index[-1]:%b %Y})",
            f"Strategies: {len(self.strategies)}   Funds: {len(self.funds)}",
        ]
        if self.risk_free is not None:
            lines.append(
                f"Risk-free column: {self.risk_free_column} ({self.risk_free_label}) - "
                "used by module 1, excluded from modules 2 and 3."
            )
        else:
            lines.append("Risk-free column: not found.")
        if self.layout == "cone":
            missing = [f.label for f in self.funds
                       if f.expected_excess is None or f.expected_vol is None]
            if missing:
                lines.append(
                    f"{len(missing)} fund(s) have no expected return/volatility, so they get "
                    "every module-1 chart except the cone."
                )
        return lines


def list_sheet_names(path: Path | str) -> list[str]:
    """The workbook's tab names, in the order Excel shows them."""
    with pd.ExcelFile(path) as workbook:
        return [str(name) for name in workbook.sheet_names]


def resolve_sheet_choice(names: Sequence[str], current: Any) -> str:
    """Pick the tab to select for a workbook, given whatever was chosen before.

    A tab that is still there is kept, including when it was saved as a sheet
    number or typed with different capitalisation. Otherwise the first tab is
    selected, which is what a new workbook should open on.
    """
    if not names:
        return str(current)
    wanted = str(current).strip()
    if wanted in names:
        return wanted
    if wanted.isdigit() and int(wanted) < len(names):
        return names[int(wanted)]
    for name in names:
        if name.strip().casefold() == wanted.casefold():
            return name
    return names[0]


def column_letter(index: int) -> str:
    """Zero-based column number to its Excel letter, so messages match the sheet."""
    letters = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def parse_date_value(value: Any) -> pd.Timestamp | pd.NaT:
    """Read a date cell however Excel/pandas presents it.

    Real dates arrive as timestamps, a sheet saved with a text column arrives as
    a string (read day-first, matching UK workbooks), and a column stored as a
    number arrives as an Excel serial, which needs the 1900 date system origin
    rather than being read as nanoseconds since 1970.
    """
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return pd.NaT
    if isinstance(value, (pd.Timestamp, _dt.datetime, _dt.date)):
        return pd.Timestamp(value)
    if isinstance(value, (int, float, np.integer, np.floating)):
        serial = float(value)
        if not 1.0 <= serial <= 2958465.0:  # 1900-01-01 to 9999-12-31
            return pd.NaT
        try:
            return pd.to_datetime(serial, unit="D", origin="1899-12-30")
        except (ValueError, OverflowError):
            return pd.NaT
    text = str(value).strip()
    if not text:
        return pd.NaT
    try:
        return pd.to_datetime(text, dayfirst=True)
    except (ValueError, TypeError):
        return pd.NaT


def month_end(value: pd.Timestamp) -> pd.Timestamp:
    """Snap any day in a month to that month's last day, at midnight.

    Workbooks mix 31/01/2020, 2020-01-01 and 2020-01-31 00:00:00 for the same
    observation. Snapping first means every later join, reindex and rolling
    window lines up regardless of which convention a sheet uses.
    """
    return value.to_period("M").to_timestamp(how="end").normalize()


def to_decimal_series(values: pd.Series) -> tuple[pd.Series, list[int]]:
    """Convert a raw column to decimal returns and report unreadable cells.

    Numbers pass straight through, because a percentage-formatted Excel cell is
    already stored as a decimal. Text such as "5%" is converted. Anything else
    that is not blank is reported by position, so the caller can name the Excel
    row it sits on.
    """
    numeric = pd.to_numeric(values, errors="coerce").astype(float)
    unreadable: list[int] = []
    if numeric.notna().all():
        return numeric, unreadable
    for position, raw_value in enumerate(values):
        if not pd.isna(numeric.iloc[position]) or pd.isna(raw_value):
            continue
        text = str(raw_value).strip().replace(",", "")
        if not text:
            continue
        scale = 0.01 if text.endswith("%") else 1.0
        try:
            numeric.iloc[position] = float(text.rstrip("%").strip()) * scale
        except ValueError:
            unreadable.append(position)
    return numeric, unreadable


def detect_layout(raw: pd.DataFrame) -> tuple[Literal["cone", "plain"], int]:
    """Work out where the monthly rows start, and therefore which layout it is."""
    first_date_row: int | None = None
    for row in range(min(len(raw), 12)):
        if pd.notna(parse_date_value(raw.iat[row, 0])):
            first_date_row = row
            break
    if first_date_row is None:
        raise WorkbookFormatError(
            "No month-end dates were found in column A.\n\n"
            + LAYOUT_DESCRIPTIONS["cone"]
            + "\n"
            + LAYOUT_DESCRIPTIONS["plain"]
        )
    if first_date_row == CONE_DATA_START_ROW:
        return "cone", CONE_DATA_START_ROW
    if first_date_row == PLAIN_DATA_START_ROW:
        return "plain", PLAIN_DATA_START_ROW
    raise WorkbookFormatError(
        f"The first date in column A is on Excel row {first_date_row + 1}, which matches "
        "neither supported layout.\n\n"
        f"Expected row 5 for the cone layout, or row 3 for the plain layout.\n\n"
        + LAYOUT_DESCRIPTIONS["cone"]
        + "\n"
        + LAYOUT_DESCRIPTIONS["plain"]
    )


def _header_text(row: pd.Series, column: int) -> str:
    if column >= len(row):
        return ""
    value = row.iloc[column]
    if pd.isna(value):
        return ""
    text = str(value).strip()
    return "" if text.lower().startswith("unnamed:") else text


def _expected_figure(row: pd.Series, column: int) -> tuple[float | None, str | None]:
    """Read one expected return/volatility cell from the cone header rows."""
    if column >= len(row):
        return None, None
    value = row.iloc[column]
    if pd.isna(value) or (isinstance(value, str) and not value.strip()):
        return None, None
    converted, unreadable = to_decimal_series(pd.Series([value]))
    if unreadable or pd.isna(converted.iloc[0]):
        return None, f"'{value}' is not a number or percentage"
    return float(converted.iloc[0]), None


def load_workbook(
    path: Path,
    sheet: Any,
    *,
    stop_on_missing_months: bool = True,
    stop_on_history_gaps: bool = True,
) -> WorkbookData:
    """Read, check and normalise a workbook, or explain everything that is wrong.

    All problems are collected and raised together, so one run of the check
    tells the user the full list rather than one item at a time.
    """
    try:
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
    except FileNotFoundError as exc:
        raise WorkbookFormatError(f"The workbook was not found:\n{path}") from exc
    except ValueError as exc:
        try:
            available = "\n".join(f"  - {name}" for name in list_sheet_names(path))
            tabs = f"\n\nThe tabs in this workbook are:\n{available}"
        except Exception:
            tabs = ""
        raise WorkbookFormatError(
            f"Worksheet '{sheet}' could not be read from:\n{path}\n\n{exc}{tabs}"
        ) from exc

    if raw.empty or raw.shape[1] < 2:
        raise WorkbookFormatError(
            "The worksheet needs column A for dates and at least one fund column from column B."
        )

    layout, data_start = detect_layout(raw)
    errors: list[str] = []
    notes: list[str] = []

    # ---- Dates -------------------------------------------------------------
    raw_dates = raw.iloc[data_start:, 0]
    parsed = raw_dates.map(parse_date_value)
    unreadable_dates = [
        f"row {position + data_start + 1} ('{raw_dates.iloc[position]}')"
        for position in range(len(parsed))
        if pd.isna(parsed.iloc[position]) and pd.notna(raw_dates.iloc[position])
    ]
    if unreadable_dates:
        errors.append(
            "Column A contains "
            f"{len(unreadable_dates)} value(s) that are not dates: "
            + ", ".join(unreadable_dates[:8])
            + ("..." if len(unreadable_dates) > 8 else "")
        )
    keep = parsed.notna().to_numpy()
    if not keep.any():
        raise WorkbookFormatError(
            f"No readable dates were found from Excel row {data_start + 1} onwards."
        )

    body = raw.iloc[data_start:, :].loc[keep].reset_index(drop=True)
    dates = pd.Series([month_end(value) for value in parsed[keep]], name="date")

    if not dates.is_monotonic_increasing:
        order = dates.sort_values(kind="stable").index
        dates = dates.loc[order].reset_index(drop=True)
        body = body.loc[order].reset_index(drop=True)
        notes.append("Dates were not in ascending order, so the rows were sorted oldest first.")

    if dates.duplicated().any():
        duplicated = sorted({value.strftime("%b %Y") for value in dates[dates.duplicated()]})
        errors.append(
            "More than one row was found for the same month: "
            + ", ".join(duplicated[:10])
            + ("..." if len(duplicated) > 10 else "")
            + ". Each calendar month needs exactly one row."
        )

    index = pd.DatetimeIndex(dates, name="date")
    full_index = pd.DatetimeIndex(
        [month_end(stamp) for stamp in pd.period_range(
            index.min().to_period("M"), index.max().to_period("M"), freq="M"
        ).to_timestamp()],
        name="date",
    )
    missing_months = full_index.difference(index)
    if len(missing_months):
        described = ", ".join(stamp.strftime("%b %Y") for stamp in missing_months[:10])
        if stop_on_missing_months:
            errors.append(
                f"{len(missing_months)} calendar month(s) are absent from column A: "
                + described
                + ("..." if len(missing_months) > 10 else "")
                + ". Add the missing rows, or untick 'Stop if a month is missing "
                "from the date column' to treat them as missing data."
            )
        else:
            notes.append(
                f"{len(missing_months)} absent calendar month(s) were inserted as missing data: "
                + described
                + ("..." if len(missing_months) > 10 else "")
                + ". Rolling windows over these months produce no value."
            )

    # ---- Fund columns ------------------------------------------------------
    strategy_row = raw.iloc[ROW_STRATEGY]
    fundname_row = raw.iloc[ROW_FUNDNAME]
    excess_row = raw.iloc[ROW_CONE_EXPECTED_EXCESS] if layout == "cone" else None
    vol_row = raw.iloc[ROW_CONE_EXPECTED_VOL] if layout == "cone" else None

    funds: list[FundColumn] = []
    risk_free: pd.Series | None = None
    risk_free_column: str | None = None
    risk_free_label: str | None = None
    seen_names: dict[tuple[str, str], str] = {}

    for column in range(1, raw.shape[1]):
        letter = column_letter(column)
        strategy = _header_text(strategy_row, column)
        fund_name = _header_text(fundname_row, column)
        cells = body.iloc[:, column] if column < body.shape[1] else pd.Series(dtype=object)
        returns, unreadable = to_decimal_series(cells)
        returns.index = index
        has_data = returns.notna().any()

        if not strategy and not fund_name and not has_data:
            continue  # a genuinely empty spacer column

        is_risk_free = strategy.lower().startswith(RISK_FREE_PREFIX)

        if not is_risk_free:
            if not strategy:
                errors.append(
                    f"Column {letter} holds data but row 1 has no strategy name."
                )
                continue
            if not fund_name:
                if not has_data:
                    notes.append(f"Column {letter} has no fund name and no data, so it was ignored.")
                    continue
                errors.append(
                    f"Column {letter} (strategy '{strategy}') holds data but row 2 has no fund name."
                )
                continue
            if (strategy, fund_name) in seen_names:
                errors.append(
                    f"'{fund_name}' appears twice in strategy '{strategy}' "
                    f"(columns {seen_names[(strategy, fund_name)]} and {letter}). "
                    "Fund names must be unique within a strategy."
                )
                continue
            seen_names[(strategy, fund_name)] = letter
        else:
            if risk_free is not None:
                errors.append(
                    f"More than one risk-free column was found (columns {risk_free_column} and {letter})."
                )
                continue
            risk_free_column = letter
            risk_free_label = fund_name or strategy

        described = f"'{fund_name}' ({strategy})" if not is_risk_free else f"the risk-free column {letter}"

        if unreadable:
            rows = ", ".join(str(position + data_start + 1) for position in unreadable[:8])
            errors.append(
                f"{described} has {len(unreadable)} cell(s) that are neither a number, "
                f"a percentage nor blank, on Excel row(s) {rows}"
                + ("..." if len(unreadable) > 8 else "")
                + "."
            )
            continue

        observed = returns.dropna()
        if observed.empty:
            notes.append(f"{described} has no returns at all, so it was left out of every module.")
            continue

        if (observed <= -1.0).any():
            when = ", ".join(stamp.strftime("%b %Y") for stamp in observed.index[observed <= -1.0][:5])
            errors.append(f"{described} has a monthly return of -100% or worse ({when}).")
            continue

        extreme = observed.abs() > IMPOSSIBLE_MONTHLY_RETURN
        if extreme.any():
            worst = observed[extreme].abs().max()
            errors.append(
                f"{described} contains a monthly return of {worst * 100:.0f}%, which is far outside "
                "any plausible range. Returns must be entered as Excel percentages, so that 5% is "
                "stored as 0.05 - a cell containing 5 means 500%."
            )
            continue

        unusual = int((observed.abs() > UNUSUAL_MONTHLY_RETURN).sum())
        if unusual:
            notes.append(
                f"{described} has {unusual} month(s) beyond +/-{UNUSUAL_MONTHLY_RETURN * 100:.0f}%. "
                "Please confirm the cells are formatted as percentages."
            )

        live = returns.loc[observed.index[0]:observed.index[-1]]
        gaps = live.index[live.isna()]
        if len(gaps):
            described_gaps = ", ".join(stamp.strftime("%b %Y") for stamp in gaps[:8])
            if stop_on_history_gaps:
                errors.append(
                    f"{described} has {len(gaps)} blank month(s) inside its live history "
                    f"({observed.index[0]:%b %Y} to {observed.index[-1]:%b %Y}): "
                    + described_gaps
                    + ("..." if len(gaps) > 8 else "")
                    + ". Fill them in, or untick 'Stop if a fund has blank months inside its "
                    "history' to treat them as missing data."
                )
                continue
            notes.append(
                f"{described} has {len(gaps)} blank month(s) inside its live history: "
                + described_gaps
                + ("..." if len(gaps) > 8 else "")
                + "."
            )

        if not is_risk_free and observed.nunique() == 1:
            errors.append(
                f"{described} reports the same return every month ({observed.iloc[0] * 100:.2f}%), "
                "so its volatility and correlations are undefined."
            )
            continue

        if is_risk_free:
            risk_free = returns
            continue

        expected_excess = expected_vol = None
        if layout == "cone":
            expected_excess, excess_problem = _expected_figure(excess_row, column)
            expected_vol, vol_problem = _expected_figure(vol_row, column)
            if excess_problem:
                errors.append(f"{described}: expected excess return on row 3 - {excess_problem}.")
                continue
            if vol_problem:
                errors.append(f"{described}: expected volatility on row 4 - {vol_problem}.")
                continue
            if expected_excess is not None and expected_excess <= -1.0:
                errors.append(f"{described} has an expected excess return of -100% or worse.")
                continue
            if expected_vol is not None and expected_vol < 0.0:
                errors.append(f"{described} has a negative expected volatility.")
                continue

        funds.append(
            FundColumn(
                strategy=strategy,
                fund=fund_name,
                column=letter,
                returns=returns,
                expected_excess=expected_excess,
                expected_vol=expected_vol,
            )
        )

    if not funds and not errors:
        errors.append("No fund columns with usable returns were found from column B onwards.")

    if errors:
        raise WorkbookFormatError(
            f"{len(errors)} problem(s) were found in {Path(path).name} [{sheet}].\n"
            f"Detected {LAYOUT_DESCRIPTIONS[layout]}\n\n"
            + "\n\n".join(f"{number}. {message}" for number, message in enumerate(errors, start=1))
        )

    # Reindexing onto the complete calendar is what makes every rolling window
    # count real months rather than however many rows happen to be present.
    for fund in funds:
        fund.returns = fund.returns.reindex(full_index)
    if risk_free is not None:
        risk_free = risk_free.reindex(full_index)

    return WorkbookData(
        path=Path(path),
        sheet=sheet,
        layout=layout,
        index=full_index,
        funds=funds,
        risk_free=risk_free,
        risk_free_column=risk_free_column,
        risk_free_label=risk_free_label,
        notes=notes,
    )


def check_modules_against_layout(layout: str, modules: Iterable[str]) -> list[str]:
    """Report any selected module the detected layout cannot support."""
    problems: list[str] = []
    if "cone" in modules and layout != "cone":
        problems.append(
            "Module 1 (cone charts) needs the expected excess return on row 3 and the expected "
            "volatility on row 4, with monthly returns starting on row 5. This sheet uses the "
            "plain layout, where returns start on row 3.\n\n" + LAYOUT_DESCRIPTIONS["cone"]
        )
    return problems


# =============================================================================
# Plotting imports
#
# The Agg backend is selected before pyplot is imported so that charts render
# identically whether the script was started from the settings window, a
# terminal or a scheduled task, and never try to open a display of their own.
# =============================================================================

matplotlib.use("Agg")

import matplotlib.colors as mcolors  # noqa: E402
import matplotlib.pyplot as plt  # noqa: E402
import seaborn as sns  # noqa: E402
from matplotlib.backends.backend_pdf import PdfPages  # noqa: E402
from matplotlib.lines import Line2D  # noqa: E402
from matplotlib.ticker import PercentFormatter  # noqa: E402
from scipy.cluster.hierarchy import dendrogram, fcluster, linkage  # noqa: E402
from scipy.spatial.distance import squareform  # noqa: E402
from scipy.stats import rankdata  # noqa: E402


def safe_file_stem(value: str, limit: int = 120) -> str:
    """A stable file name that is legal on Windows as well as Linux/macOS."""
    cleaned = re.sub(r'[<>:"/\\|?*]+', "_", str(value))
    cleaned = re.sub(r"\s+", "_", cleaned.strip().rstrip("."))
    return cleaned[:limit] or "unnamed"


def unique_stem(stem: str, used: set[str]) -> str:
    """Keep two similarly named funds from overwriting each other's files."""
    candidate = stem
    suffix = 2
    while candidate.casefold() in used:
        candidate = f"{stem}_{suffix}"
        suffix += 1
    used.add(candidate.casefold())
    return candidate


def format_percent(value: float, decimals: int = 1) -> str:
    return f"{value * 100:.{decimals}f}%" if np.isfinite(value) else "n/a"


def format_ratio(numerator: float, denominator: float) -> str:
    if not np.isfinite(numerator) or not np.isfinite(denominator) or denominator == 0:
        return "-"
    return f"{numerator / denominator:.2f}"


def rolling_total_return(series: pd.Series, window: int) -> pd.Series:
    """Compounded return over a complete trailing window, via log returns."""
    log_returns = np.log1p(series.astype(float))
    return np.expm1(log_returns.rolling(window=window, min_periods=window).sum())


# =============================================================================
# Module 1 - cone charts
# =============================================================================

CONE_FIG_SIZE = (9.7, 4.15)
CONE_TITLE_FONT_SIZE = 12


def relative_excess_returns(fund: pd.Series, risk_free: pd.Series | None) -> pd.Series:
    """Growth of 1 in the fund relative to 1 held at the risk-free rate."""
    if risk_free is None:
        return fund.astype(float)
    return (1.0 + fund.astype(float)) / (1.0 + risk_free.astype(float)) - 1.0


def arithmetic_excess_returns(fund: pd.Series, risk_free: pd.Series | None) -> pd.Series:
    """The conventional periodic excess return used for Sharpe-style figures."""
    if risk_free is None:
        return fund.astype(float)
    return fund.astype(float) - risk_free.astype(float)


def annualise_stats(series: pd.Series) -> tuple[float, float, float]:
    clean = series.dropna()
    if clean.empty:
        return np.nan, np.nan, np.nan
    annual_return = clean.mean() * 12.0
    annual_vol = clean.std(ddof=1) * math.sqrt(12.0)
    sharpe = np.nan if not annual_vol else annual_return / annual_vol
    return annual_return, annual_vol, sharpe


def cone_window(
    fund: pd.Series,
    risk_free: pd.Series | None,
    *,
    require_risk_free: bool,
    n_months: int | None,
) -> tuple[pd.Series, pd.Series | None] | None:
    """Trim to the fund's live history, optionally capped at a fixed length."""
    valid = fund.notna()
    if require_risk_free and risk_free is not None:
        valid &= risk_free.notna()
    if not valid.any():
        return None
    live = valid[valid].index
    fund = fund.loc[live[0]:live[-1]]
    trimmed_rf = risk_free.loc[live[0]:live[-1]] if risk_free is not None else None
    if n_months:
        fund = fund.iloc[:n_months]
        trimmed_rf = trimmed_rf.iloc[:n_months] if trimmed_rf is not None else None
    return fund, trimmed_rf


def _apply_xlim(axis: plt.Axes, dates: pd.Index, plot_xlim: tuple | None) -> None:
    if plot_xlim:
        axis.set_xlim(plot_xlim)
    elif len(dates):
        axis.set_xlim(dates[0], dates[-1])


def make_cone_chart(
    fund: FundColumn,
    returns: pd.Series,
    risk_free: pd.Series | None,
    plot_xlim: tuple | None,
    dpi: int,
) -> plt.Figure:
    """Cumulative relative excess return against the expected path and its bands."""
    excess = relative_excess_returns(returns, risk_free)
    # Summing monthly log excess returns is numerically stable and gives the log
    # of fund wealth divided by risk-free wealth.
    cumulative_log = np.log1p(excess).cumsum()

    years = np.arange(1, len(excess) + 1, dtype=float) / 12.0
    expected_path = years * np.log1p(fund.expected_excess)
    one_sigma = np.sqrt(years) * fund.expected_vol
    two_sigma = 2.0 * one_sigma

    actual_return, actual_vol, _ = annualise_stats(arithmetic_excess_returns(returns, risk_free))

    figure, (axis, table_axis) = plt.subplots(
        nrows=2,
        ncols=1,
        figsize=CONE_FIG_SIZE,
        dpi=dpi,
        gridspec_kw={"height_ratios": [3.9, 1.0], "hspace": 0.15},
    )
    dates = returns.index
    axis.plot(dates, expected_path, label="Expected return", linewidth=2, color="black")
    axis.plot(dates, expected_path + one_sigma, label="+/-1 standard deviation",
              linestyle="--", linewidth=1.6, color="darkblue")
    axis.plot(dates, expected_path - one_sigma, linestyle="--", linewidth=1.6, color="darkblue")
    axis.plot(dates, expected_path + two_sigma, label="+/-2 standard deviations",
              linestyle=":", linewidth=1.6, color="darkblue")
    axis.plot(dates, expected_path - two_sigma, linestyle=":", linewidth=1.6, color="darkblue")
    axis.plot(dates, cumulative_log, label="Cumulative relative excess return (actual)",
              linewidth=2.2, color="#006400")

    axis.set_title(
        f"{fund.strategy}\n{fund.fund}\nCumulative Relative Excess Return vs Expectations (log)",
        fontsize=CONE_TITLE_FONT_SIZE,
        pad=18,
    )
    axis.set_ylabel("Log relative wealth")
    axis.grid(True, alpha=0.25)
    _apply_xlim(axis, dates, plot_xlim)
    axis.legend(loc="upper left")

    table_axis.axis("off")
    table = table_axis.table(
        cellText=[
            ["Expected", format_percent(fund.expected_excess), format_percent(fund.expected_vol),
             format_ratio(fund.expected_excess, fund.expected_vol)],
            ["Actual", format_percent(actual_return), format_percent(actual_vol),
             format_ratio(actual_return, actual_vol)],
        ],
        colLabels=["", "Excess Return (pa)", "Std Dev (pa)", "Return-to-Risk"],
        cellLoc="center",
        colLoc="center",
        bbox=[0.02, 0.05, 0.96, 0.9],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1.0, 1.1)
    figure.subplots_adjust(left=0.08, right=0.98, top=0.86, bottom=0.18)
    return figure


def make_rolling_return_chart(
    fund: FundColumn,
    returns: pd.Series,
    risk_free: pd.Series | None,
    window: int,
    volatility_window: int,
    plot_xlim: tuple | None,
    dpi: int,
) -> plt.Figure | None:
    """Rolling total return, with the risk-free line and rolling volatility."""
    fund_rolling = rolling_total_return(returns, window)
    if fund_rolling.dropna().empty:
        return None
    risk_free_rolling = rolling_total_return(risk_free, window) if risk_free is not None else None
    rolling_vol = returns.rolling(
        window=volatility_window, min_periods=volatility_window
    ).std(ddof=1) * np.sqrt(12)

    trailing_12 = rolling_total_return(returns, 12).dropna()
    trailing_24 = rolling_total_return(returns, 24).dropna()
    latest_vol = rolling_vol.dropna()
    metrics = []
    if not trailing_12.empty:
        metrics.append(f"12m: {format_percent(trailing_12.iloc[-1])}")
    if not trailing_24.empty:
        metrics.append(f"24m: {format_percent(trailing_24.iloc[-1])}")
    if not latest_vol.empty:
        metrics.append(f"{volatility_window}m vol: {format_percent(latest_vol.iloc[-1])}")
    metrics_text = f" ({' | '.join(metrics)})" if metrics else ""

    figure, axis = plt.subplots(figsize=CONE_FIG_SIZE, dpi=dpi)
    dates = returns.index
    axis.plot(dates, fund_rolling, label=f"{fund.fund} rolling total return",
              color="#1f77b4", linewidth=2)
    if risk_free_rolling is not None:
        axis.plot(dates, risk_free_rolling, label="Risk-free rolling total return",
                  color="#d62728", linewidth=1.8)

    volatility_axis = axis.twinx()
    volatility_axis.plot(dates, rolling_vol, linestyle="--", linewidth=1.8, color="#ff7f0e",
                         label=f"Rolling {volatility_window}M Vol (ann.)")
    volatility_axis.set_ylabel("Annualised Volatility")
    volatility_axis.yaxis.set_major_formatter(PercentFormatter(1.0, decimals=1))
    volatility_axis.set_ylim(bottom=0)

    axis.axhline(0.0, color="gray", linewidth=1, linestyle="--", alpha=0.7)
    axis.set_title(
        f"{fund.strategy}\n{fund.fund}\nRolling {window}-Month Total Return{metrics_text}",
        fontsize=CONE_TITLE_FONT_SIZE,
        pad=18,
    )
    axis.set_ylabel(f"Total return over trailing {window} months")
    axis.yaxis.set_major_formatter(PercentFormatter(1.0, decimals=1))
    _apply_xlim(axis, dates, plot_xlim)
    axis.margins(x=0)
    axis.grid(True, alpha=0.25)

    handles, labels = axis.get_legend_handles_labels()
    extra_handles, extra_labels = volatility_axis.get_legend_handles_labels()
    axis.legend(handles + extra_handles, labels + extra_labels, loc="upper left")
    figure.subplots_adjust(left=0.1, right=0.9, top=0.86, bottom=0.2)
    return figure


def make_sharpe_chart(
    fund: FundColumn,
    returns: pd.Series,
    risk_free: pd.Series | None,
    window: int,
    plot_xlim: tuple | None,
    dpi: int,
) -> plt.Figure | None:
    excess = arithmetic_excess_returns(returns, risk_free)
    rolling_mean = excess.rolling(window=window, min_periods=window).mean() * 12
    rolling_vol = excess.rolling(window=window, min_periods=window).std(ddof=1) * np.sqrt(12)
    rolling_sharpe = rolling_mean / rolling_vol
    if rolling_sharpe.dropna().empty:
        return None

    window_label = f"{window // 12}-Year" if window % 12 == 0 else f"{window}-Month"
    figure, axis = plt.subplots(figsize=CONE_FIG_SIZE, dpi=dpi)
    axis.plot(returns.index, rolling_sharpe, color="#005f73", linewidth=1.8,
              label=f"Rolling {window_label} Sharpe")
    axis.axhline(0.0, color="gray", linewidth=1, linestyle="--", alpha=0.7)
    axis.set_title(
        f"{fund.strategy}\n{fund.fund}\nRolling {window_label} Sharpe Ratio",
        fontsize=CONE_TITLE_FONT_SIZE,
        pad=18,
    )
    axis.set_ylabel("Sharpe ratio")
    _apply_xlim(axis, returns.index, plot_xlim)
    axis.margins(x=0)
    axis.grid(True, alpha=0.25)
    axis.legend(loc="upper left")
    figure.subplots_adjust(left=0.1, right=0.98, top=0.86, bottom=0.2)
    return figure


def make_drawdown_chart(
    fund: FundColumn,
    returns: pd.Series,
    threshold: float,
    plot_xlim: tuple | None,
    dpi: int,
) -> plt.Figure | None:
    """Drawdown from peak, with a table of the maximum and time spent below it."""
    observed = returns.dropna()
    if observed.empty:
        return None
    cumulative = (1.0 + observed).cumprod()
    drawdowns = cumulative / cumulative.cummax() - 1.0
    max_drawdown = drawdowns.min()
    below_threshold = float((drawdowns <= -threshold).mean() * 100.0)

    figure, (axis, table_axis) = plt.subplots(
        nrows=2,
        ncols=1,
        figsize=CONE_FIG_SIZE,
        dpi=dpi,
        gridspec_kw={"height_ratios": [3.9, 1.0], "hspace": 0.15},
    )
    axis.fill_between(drawdowns.index, drawdowns, 0.0, where=drawdowns < 0.0,
                      color="#ffa0a0", alpha=0.7, interpolate=True)
    axis.plot(drawdowns.index, drawdowns, color="#b22222", linewidth=1.6)
    axis.axhline(0.0, color="black", linewidth=1, linestyle="--", alpha=0.6)
    axis.set_title(
        f"{fund.strategy}\n{fund.fund}\nTotal Drawdown (Max: {format_percent(max_drawdown)})",
        fontsize=CONE_TITLE_FONT_SIZE,
        pad=18,
    )
    axis.set_ylabel("Drawdown from peak")
    axis.yaxis.set_major_formatter(PercentFormatter(1.0, decimals=1))
    _apply_xlim(axis, drawdowns.index, plot_xlim)
    axis.margins(x=0)
    axis.grid(True, alpha=0.25)

    table_axis.axis("off")
    table = table_axis.table(
        cellText=[
            ["Max drawdown (from peak)", format_percent(max_drawdown)],
            [f"% of track record <= -{threshold * 100:g}% drawdown", f"{below_threshold:.1f}%"],
        ],
        colLabels=["Metric", "Value"],
        cellLoc="center",
        colLoc="center",
        bbox=[0.05, 0.15, 0.90, 0.75],
    )
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1.0, 1.12)
    figure.subplots_adjust(left=0.1, right=0.98, top=0.86, bottom=0.18)
    return figure


def run_cone_module(data: WorkbookData, options: dict[str, Any], output_dir: Path) -> list[str]:
    """Create every module-1 chart, plus the compiled PDF when requested."""
    output_dir.mkdir(parents=True, exist_ok=True)
    log: list[str] = []
    dpi = int(options["dpi"])
    window = int(options["rolling_window"])
    sharpe_window = int(options["sharpe_window"])
    volatility_window = int(options["volatility_window"])
    threshold = float(options["drawdown_threshold_pct"]) / 100.0
    n_months = int(options["n_months"]) if options["use_predetermined_months"] else None
    use_risk_free = bool(options["use_risk_free"]) and data.risk_free is not None
    risk_free = data.risk_free if use_risk_free else None

    if bool(options["use_risk_free"]) and data.risk_free is None:
        log.append("No risk-free column was found, so module 1 ran on total returns.")

    plot_xlim = None
    if options["common_x_axis"]:
        axis_dates = data.index[:n_months] if n_months else data.index
        if len(axis_dates):
            plot_xlim = (axis_dates[0], axis_dates[-1])

    charts_by_strategy: dict[str, list[dict[str, Any]]] = {}
    used_stems: set[str] = set()

    for fund in data.funds:
        window_data = cone_window(
            fund.returns, risk_free, require_risk_free=use_risk_free, n_months=n_months
        )
        if window_data is None:
            log.append(f"{fund.label}: no months overlap the risk-free series, so it was skipped.")
            continue
        returns, fund_risk_free = window_data
        stem = unique_stem(safe_file_stem(fund.fund), used_stems)
        images: dict[str, str | None] = {}

        if fund.expected_excess is None or fund.expected_vol is None:
            log.append(
                f"{fund.label}: no cone chart, because row 3 or row 4 has no expected figure."
            )
            images["cone"] = None
        else:
            figure = make_cone_chart(fund, returns, fund_risk_free, plot_xlim, dpi)
            images["cone"] = str(output_dir / f"{stem}_cone.png")
            figure.savefig(images["cone"], dpi=dpi, bbox_inches="tight")
            plt.close(figure)

        builders: list[tuple[str, str, Callable[[], plt.Figure | None]]] = [
            ("rolling", f"{stem}_rolling_{window}m.png",
             lambda: make_rolling_return_chart(fund, returns, fund_risk_free, window,
                                               volatility_window, plot_xlim, dpi)),
            ("sharpe", f"{stem}_rolling_sharpe_{sharpe_window}m.png",
             lambda: make_sharpe_chart(fund, returns, fund_risk_free, sharpe_window, plot_xlim, dpi)),
            ("drawdown", f"{stem}_drawdown.png",
             lambda: make_drawdown_chart(fund, returns, threshold, plot_xlim, dpi)),
        ]
        required_months = {"rolling": window, "sharpe": sharpe_window, "drawdown": 1}
        for key, filename, builder in builders:
            figure = builder()
            if figure is None:
                images[key] = None
                log.append(
                    f"{fund.label}: no {key} chart, because it needs at least "
                    f"{required_months[key]} month(s) of returns."
                )
                continue
            images[key] = str(output_dir / filename)
            figure.savefig(images[key], dpi=dpi, bbox_inches="tight")
            plt.close(figure)

        charts_by_strategy.setdefault(fund.strategy, []).append(
            {"fund": fund.fund, "images": images}
        )

    saved = sum(1 for charts in charts_by_strategy.values()
                for entry in charts for path in entry["images"].values() if path)
    log.append(f"Saved {saved} chart image(s) to {output_dir}.")

    if options["create_pdf"] and charts_by_strategy:
        pdf_path = output_dir / "All_Funds_Cones.pdf"
        try:
            _write_cone_pdf(pdf_path, charts_by_strategy)
            log.append(f"Compiled PDF: {pdf_path}")
        except PermissionError:
            log.append(
                f"The compiled PDF could not be written to {pdf_path}. "
                "Close the file if it is open in another application."
            )
    return log


def _write_cone_pdf(pdf_path: Path, charts_by_strategy: dict[str, list[dict[str, Any]]]) -> None:
    """One title page per strategy, then one portrait A4 page per fund."""
    a4_width, a4_height = 8.27, 11.69
    page_dpi = 180
    with PdfPages(pdf_path) as pdf:
        for strategy, fund_charts in charts_by_strategy.items():
            title_figure = plt.figure(figsize=(a4_width, a4_height), dpi=page_dpi)
            title_figure.patch.set_facecolor("white")
            title_figure.text(0.5, 0.6, strategy, ha="center", va="center",
                              fontsize=36, fontweight="bold")
            title_figure.text(0.5, 0.4, "Performance Analytics", ha="center", va="center", fontsize=18)
            pdf.savefig(title_figure, dpi=page_dpi)
            plt.close(title_figure)

            for entry in fund_charts:
                images = [entry["images"].get(key) for key in ("cone", "rolling", "sharpe", "drawdown")]
                images = [path for path in images if path]
                if not images:
                    continue
                page = plt.figure(figsize=(a4_width, a4_height), dpi=page_dpi)
                grid = page.add_gridspec(nrows=len(images), ncols=1, hspace=0.02,
                                         top=0.99, bottom=0.01, left=0.02, right=0.98)
                for position, image_path in enumerate(images):
                    axis = page.add_subplot(grid[position, 0])
                    axis.axis("off")
                    axis.imshow(plt.imread(image_path))
                pdf.savefig(page, dpi=page_dpi)
                plt.close(page)


# =============================================================================
# Module 2 - correlation and clustering
#
# Note on "exclusions" here: by the time this module runs, the workbook has
# already passed the format check, so nothing in this section is a data error.
# A fund is excluded from a particular timeframe only because it is too young
# for the requested window, has stopped reporting, or has no variation inside
# that window. Each exclusion is written to the audit CSV with its reason.
# =============================================================================

Timeframe = int | Literal["max"]


def normalise_timeframes(values: Iterable[Any]) -> tuple[Timeframe, ...]:
    parsed: list[Timeframe] = []
    for value in values:
        if isinstance(value, str) and value.strip().lower() == "max":
            item: Timeframe = "max"
        else:
            try:
                item = int(value)
            except (TypeError, ValueError) as exc:
                raise ValueError(
                    f"'{value}' is not a lookback. Use whole numbers of months and/or 'max', "
                    "for example 36,60,max."
                ) from exc
            if item < 2:
                raise ValueError("Each numeric lookback must be at least 2 months.")
        if item not in parsed:
            parsed.append(item)
    if not parsed:
        raise ValueError("Enter at least one lookback period.")
    return tuple(parsed)


def parse_timeframes(text: str) -> tuple[Timeframe, ...]:
    return normalise_timeframes(part.strip() for part in str(text).split(",") if part.strip())


def timeframe_label(timeframe: Timeframe) -> str:
    return "MaxAvailable" if timeframe == "max" else f"{timeframe}mths"


@dataclass
class ClusterAnalysis:
    name: str
    timeframe_label: str
    returns: pd.DataFrame
    correlation: pd.DataFrame
    linkage_matrix: np.ndarray
    leaf_order: list[int]
    assignments: pd.DataFrame


def latest_contiguous_complete_block(frame: pd.DataFrame) -> pd.DataFrame:
    """The uninterrupted run of months, ending at the latest one, with no gaps."""
    complete = frame.notna().all(axis=1).to_numpy()
    count = 0
    for is_complete in complete[::-1]:
        if not is_complete:
            break
        count += 1
    return frame.tail(count) if count else frame.iloc[0:0]


def select_cluster_returns(
    frame: pd.DataFrame,
    timeframe: Timeframe,
    options: dict[str, Any],
    strategy: str,
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    exclusions: list[dict[str, Any]] = []
    minimum_months = int(options["min_required_months"])
    counts = frame.notna().sum()
    keep = counts[counts >= minimum_months].index.tolist()
    for fund in counts.index.difference(keep):
        exclusions.append({
            "strategy": strategy, "fund": fund, "timeframe": timeframe_label(timeframe),
            "reason": f"Track record shorter than the {minimum_months}-month minimum",
            "available_months": int(counts[fund]),
        })
    eligible = frame[keep]
    if eligible.empty:
        return eligible, exclusions

    # Use the latest month this strategy actually reports, rather than the
    # workbook's final month, so one late-reporting strategy cannot empty another.
    reported = eligible.dropna(axis=0, how="all")
    if reported.empty:
        return eligible.iloc[0:0], exclusions
    latest_month = reported.index.max()
    eligible = eligible.loc[:latest_month]

    if options["require_latest_return"]:
        current = eligible.loc[latest_month].notna()
        for fund in current.index[~current]:
            exclusions.append({
                "strategy": strategy, "fund": fund, "timeframe": timeframe_label(timeframe),
                "reason": f"No return for the latest month ({latest_month})",
                "available_months": int(eligible[fund].notna().sum()),
            })
        eligible = eligible.loc[:, current]
    if eligible.shape[1] < 2:
        return eligible.iloc[0:0], exclusions

    if timeframe == "max":
        selected = latest_contiguous_complete_block(eligible)
    else:
        requested = pd.period_range(end=latest_month, periods=timeframe, freq="M")
        window = eligible.reindex(requested)
        complete = window.notna().all(axis=0)
        for fund in complete.index[~complete]:
            trailing = latest_contiguous_complete_block(window[[fund]]).shape[0]
            exclusions.append({
                "strategy": strategy, "fund": fund, "timeframe": timeframe_label(timeframe),
                "reason": f"Incomplete {timeframe}-month window",
                "available_months": int(trailing),
            })
        selected = window.loc[:, complete]

    if not selected.empty:
        varying = selected.nunique(dropna=True) > 1
        for fund in varying.index[~varying]:
            exclusions.append({
                "strategy": strategy, "fund": fund, "timeframe": timeframe_label(timeframe),
                "reason": "Returns do not vary within this window, so correlation is undefined",
                "available_months": int(selected[fund].notna().sum()),
            })
        selected = selected.loc[:, varying]

    if selected.shape[1] < 2 or len(selected) < minimum_months:
        return selected.iloc[0:0], exclusions
    return selected, exclusions


def standardised_profiles(returns: pd.DataFrame, method: str) -> np.ndarray:
    values = returns.T.to_numpy(dtype=float)
    if method == "spearman":
        values = np.apply_along_axis(rankdata, 1, values)
    elif method == "kendall":
        raise ValueError("Ward linkage cannot be used with Kendall; choose average or complete.")
    deviations = values.std(axis=1, ddof=0, keepdims=True)
    if np.any(deviations == 0):
        raise ValueError("A fund has no variation in this window, so its correlation is undefined.")
    return (values - values.mean(axis=1, keepdims=True)) / deviations


def build_cluster_analysis(
    name: str, label: str, returns: pd.DataFrame, options: dict[str, Any]
) -> ClusterAnalysis:
    method = options["correlation_method"]
    correlation = returns.corr(method=method)
    if correlation.isna().any().any():
        undefined = correlation.columns[correlation.isna().any()].tolist()
        raise ValueError(f"Undefined correlations for: {', '.join(map(str, undefined))}")
    if options["linkage_method"] == "ward":
        linked = linkage(standardised_profiles(returns, method), method="ward",
                         metric="euclidean", optimal_ordering=True)
    else:
        distance = (1.0 - correlation).clip(lower=0.0, upper=2.0)
        np.fill_diagonal(distance.values, 0.0)
        distance = (distance + distance.T) / 2.0
        linked = linkage(squareform(distance.values, checks=True),
                         method=options["linkage_method"], optimal_ordering=True)
    leaf_order = list(dendrogram(linked, no_plot=True)["leaves"])
    cluster_count = min(int(options["max_clusters"]), returns.shape[1])
    assignments = pd.DataFrame({
        "fund": returns.columns,
        "cluster": fcluster(linked, cluster_count, criterion="maxclust"),
    }).sort_values(["cluster", "fund"], kind="stable").reset_index(drop=True)
    return ClusterAnalysis(name, label, returns, correlation, linked, leaf_order, assignments)


def _cluster_footer(figure: plt.Figure, returns: pd.DataFrame, exclusions: int,
                    options: dict[str, Any]) -> None:
    note = (
        f"Common monthly period: {returns.index.min().strftime('%b %Y')}-"
        f"{returns.index.max().strftime('%b %Y')} ({len(returns)} months); "
        f"{returns.shape[1]} funds; correlation: {options['correlation_method']}; "
        f"excluded records: {exclusions}."
    )
    figure.text(0.99, 0.006, note, ha="right", va="bottom", fontsize=7, color="0.25")


def plot_cluster_heatmap(result: ClusterAnalysis, exclusions: int,
                         options: dict[str, Any]) -> plt.Figure:
    ordered_names = result.returns.columns[result.leaf_order]
    ordered = result.correlation.loc[ordered_names, ordered_names]
    count = len(ordered)
    size = max(7.0, min(24.0, count * 0.48))
    figure, axis = plt.subplots(figsize=(size + 2.5, size + 1.5))
    wrapped = [textwrap.fill(str(name), int(options["label_wrap_width"])) for name in ordered_names]
    sns.heatmap(
        ordered, cmap="vlag", center=0, vmin=-1, vmax=1, square=True,
        linewidths=0.35, linecolor="0.88",
        annot=bool(options["annotate_heatmaps"]) and count <= int(options["max_annotated_funds"]),
        fmt=".2f", annot_kws={"size": max(5, 9 - count // 5)},
        xticklabels=wrapped, yticklabels=wrapped,
        cbar_kws={"label": "Correlation", "shrink": 0.8}, ax=axis,
    )
    axis.set_title(
        textwrap.fill(f"{result.name} - clustered fund correlation ({result.timeframe_label})", 85),
        fontsize=13, pad=14,
    )
    axis.tick_params(axis="x", labelrotation=55, labelsize=max(5, 9 - count // 8))
    axis.tick_params(axis="y", labelrotation=0, labelsize=max(5, 9 - count // 8))
    _cluster_footer(figure, result.returns, exclusions, options)
    figure.tight_layout(rect=(0.03, 0.035, 0.98, 0.96))
    return figure


def plot_cluster_dendrogram(result: ClusterAnalysis, exclusions: int, options: dict[str, Any],
                            label_width: int) -> plt.Figure:
    count = result.returns.shape[1]
    figure, axis = plt.subplots(figsize=(max(10, 7 + count * 0.18), max(7, 4 + count * 0.35)))
    threshold = options["distance_threshold"]
    dendrogram(
        result.linkage_matrix, orientation="right",
        labels=[textwrap.fill(str(name), label_width) for name in result.returns.columns],
        leaf_font_size=max(5, 9 - count // 12), distance_sort=False,
        color_threshold=float(threshold) if threshold is not None else None, ax=axis,
    )
    if threshold is not None:
        axis.axvline(float(threshold), color="crimson", linestyle="--", linewidth=1,
                     label=f"Distance threshold: {float(threshold):g}")
        axis.legend(fontsize=8)
    axis.set_xlabel(
        "Ward distance on standardised return profiles"
        if options["linkage_method"] == "ward"
        else f"{options['linkage_method'].title()} linkage distance (1 - correlation)"
    )
    axis.set_ylabel("Fund / Strategy::Fund")
    axis.grid(axis="x", linestyle="--", alpha=0.35)
    axis.set_title(
        textwrap.fill(
            f"{result.name} - hierarchical clustering "
            f"({result.timeframe_label}; {options['linkage_method']})", 85),
        fontsize=13, pad=14,
    )
    _cluster_footer(figure, result.returns, exclusions, options)
    figure.tight_layout(rect=(0.08, 0.035, 0.98, 0.96))
    return figure


def _save_cluster_figure(figure: plt.Figure, stem: Path, pdf: PdfPages | None,
                         options: dict[str, Any]) -> None:
    if options["save_png"]:
        figure.savefig(stem.with_suffix(".png"), dpi=int(options["dpi"]), bbox_inches="tight")
    if pdf is not None:
        pdf.savefig(figure, dpi=int(options["dpi"]), bbox_inches="tight")
    plt.close(figure)


def _save_cluster_outputs(result: ClusterAnalysis, exclusions: int, options: dict[str, Any],
                          pdf: PdfPages | None, output_dir: Path,
                          all_strategies: bool = False) -> None:
    stem = output_dir / f"{safe_file_stem(result.name)}_{result.timeframe_label}"
    if options["save_csv"]:
        result.correlation.to_csv(stem.with_name(stem.name + "_correlations.csv"))
        result.assignments.to_csv(stem.with_name(stem.name + "_clusters.csv"), index=False)
    _save_cluster_figure(plot_cluster_heatmap(result, exclusions, options),
                         stem.with_name(stem.name + "_heatmap"), pdf, options)
    width = int(options["all_strategies_label_wrap_width"] if all_strategies
                else options["label_wrap_width"])
    _save_cluster_figure(plot_cluster_dendrogram(result, exclusions, options, width),
                         stem.with_name(stem.name + "_dendrogram"), pdf, options)


def run_clustering_module(data: WorkbookData, options: dict[str, Any],
                          output_dir: Path) -> list[str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    log: list[str] = []
    timeframes = normalise_timeframes(options["timeframes"])

    frame = data.wide_frame()
    frame.index = frame.index.to_period("M")

    exclusions: list[dict[str, Any]] = []
    run_rows: list[dict[str, Any]] = []
    combined: dict[Timeframe, list[pd.DataFrame]] = {period: [] for period in timeframes}

    with ExitStack() as stack:
        pdfs: dict[Timeframe, PdfPages | None] = {}
        for period in timeframes:
            pdfs[period] = (
                stack.enter_context(
                    PdfPages(output_dir / f"Clustering_Report_{timeframe_label(period)}.pdf"))
                if options["create_pdf"] else None
            )

        for strategy in data.strategies:
            strategy_frame = frame.xs(strategy, axis=1, level=0, drop_level=True)
            for period in timeframes:
                selected, excluded = select_cluster_returns(strategy_frame, period, options, strategy)
                exclusions.extend(excluded)
                if selected.empty:
                    log.append(
                        f"{strategy} [{timeframe_label(period)}]: skipped - fewer than two funds "
                        "share a complete window."
                    )
                    continue
                try:
                    result = build_cluster_analysis(strategy, timeframe_label(period), selected, options)
                except ValueError as exc:
                    log.append(f"{strategy} [{timeframe_label(period)}]: skipped - {exc}")
                    continue
                _save_cluster_outputs(result, len(excluded), options, pdfs[period], output_dir)
                renamed = selected.copy()
                renamed.columns = [f"{strategy}::{fund}" for fund in selected.columns]
                combined[period].append(renamed)
                run_rows.append({
                    "strategy": strategy, "timeframe": timeframe_label(period),
                    "start": str(selected.index.min()), "end": str(selected.index.max()),
                    "months": len(selected), "funds": selected.shape[1],
                    "excluded_records": len(excluded),
                })
                log.append(
                    f"{strategy} [{timeframe_label(period)}]: {selected.shape[1]} funds "
                    f"over {len(selected)} months."
                )

        for period, frames in combined.items():
            if not frames:
                continue
            merged = latest_contiguous_complete_block(
                pd.concat(frames, axis=1, join="outer").sort_index()
            )
            if len(merged) < int(options["min_required_months"]) or merged.shape[1] < 2:
                log.append(
                    f"All strategies [{timeframe_label(period)}]: skipped - too little common history."
                )
                continue
            try:
                result = build_cluster_analysis(
                    "All Strategies Combined", timeframe_label(period), merged, options)
            except ValueError as exc:
                log.append(f"All strategies [{timeframe_label(period)}]: skipped - {exc}")
                continue
            _save_cluster_outputs(result, 0, options, pdfs[period], output_dir, all_strategies=True)
            run_rows.append({
                "strategy": "All Strategies Combined", "timeframe": timeframe_label(period),
                "start": str(merged.index.min()), "end": str(merged.index.max()),
                "months": len(merged), "funds": merged.shape[1], "excluded_records": 0,
            })
            log.append(
                f"All strategies [{timeframe_label(period)}]: {merged.shape[1]} funds "
                f"over {len(merged)} months."
            )

    if options["save_csv"]:
        audit_columns = ["strategy", "fund", "timeframe", "reason", "available_months"]
        audit = (pd.DataFrame(exclusions).drop_duplicates() if exclusions
                 else pd.DataFrame(columns=audit_columns))
        audit.to_csv(output_dir / "excluded_funds_audit.csv", index=False)
        pd.DataFrame(run_rows).to_csv(output_dir / "analysis_run_summary.csv", index=False)
        log.append("Wrote excluded_funds_audit.csv and analysis_run_summary.csv.")
    return log


# =============================================================================
# Module 3 - performance, NAV and drawdown analysis
# =============================================================================

MAX_LEGEND_ITEMS_ON_CHART = 10


def _normalise_for_cmap(value: float, low: float, high: float) -> float:
    if high == low:
        return 0.5
    return float(np.clip((value - low) / (high - low), 0, 1))


def _build_colormaps() -> dict[str, Any]:
    """Fixed colour scales, so the same number always looks the same."""
    white = mcolors.to_rgb("white")
    blue = mcolors.to_rgb("cornflowerblue")
    light_red = (1.0, 0.75, 0.75)
    dark_red = mcolors.to_rgb("firebrick")
    black = mcolors.to_rgb("black")

    drawdown_min, drawdown_max = -30.0, 0.0
    drawdown_nodes = [
        (_normalise_for_cmap(-30.0, drawdown_min, drawdown_max), black),
        (_normalise_for_cmap(-20.0, drawdown_min, drawdown_max), dark_red),
        (_normalise_for_cmap(-5.0, drawdown_min, drawdown_max), light_red),
        (_normalise_for_cmap(-5.0, drawdown_min, drawdown_max), blue),
        (_normalise_for_cmap(0.0, drawdown_min, drawdown_max), white),
    ]
    performance_min, performance_max = -35.0, 35.0
    performance_nodes = [
        (_normalise_for_cmap(value, performance_min, performance_max), colour)
        for value, colour in (
            (performance_min, dark_red), (-20.0, dark_red), (0.0, white),
            (30.0, mcolors.to_rgb("forestgreen")), (performance_max, mcolors.to_rgb("forestgreen")),
        )
    ]
    ratio_min, ratio_max = -2.0, 2.0
    ratio_nodes = [
        (_normalise_for_cmap(value, ratio_min, ratio_max), colour)
        for value, colour in (
            (ratio_min, mcolors.to_rgb("orangered")), (0.0, mcolors.to_rgb("ivory")),
            (ratio_max, mcolors.to_rgb("seagreen")),
        )
    ]
    return {
        "drawdown": mcolors.LinearSegmentedColormap.from_list(
            "drawdown", sorted(drawdown_nodes, key=lambda node: node[0])),
        "drawdown_limits": (drawdown_min, drawdown_max),
        "performance": mcolors.LinearSegmentedColormap.from_list(
            "performance", sorted(performance_nodes, key=lambda node: node[0])),
        "performance_limits": (performance_min, performance_max),
        "ratio": mcolors.LinearSegmentedColormap.from_list(
            "ratio", sorted(ratio_nodes, key=lambda node: node[0])),
        "ratio_limits": (ratio_min, ratio_max),
    }


def calculate_price_index(returns: pd.Series,
                          initial_value: float = INITIAL_PRICE_INDEX_VALUE) -> pd.Series:
    """Compound monthly returns into a NAV series starting at initial_value."""
    numeric = pd.to_numeric(returns, errors="coerce").astype(float)
    if numeric.empty or numeric.isna().all():
        return pd.Series(np.nan, index=returns.index, name=returns.name, dtype=float)
    return ((1.0 + numeric).cumprod(skipna=True) * initial_value).rename(returns.name)


def calculate_actual_drawdowns(price_index: pd.Series) -> pd.Series:
    """Drawdown from the all-time high, in percent.

    The inception value is part of the high-water mark, so a loss in the first
    reported month shows as a drawdown rather than as zero.
    """
    if price_index.empty or price_index.isna().all():
        return pd.Series(np.nan, index=price_index.index, name=price_index.name)
    peak = price_index.expanding(min_periods=1).max().clip(lower=INITIAL_PRICE_INDEX_VALUE)
    return ((price_index / peak) - 1).replace([np.inf, -np.inf], np.nan) * 100


def calculate_rolling_drawdowns(price_index: pd.Series, window: int,
                                require_full_window: bool) -> pd.Series:
    if price_index.empty or price_index.isna().all() or window <= 0:
        return pd.Series(np.nan, index=price_index.index, name=price_index.name)
    peak = price_index.rolling(window=window,
                               min_periods=window if require_full_window else 1).max()
    return ((price_index / peak) - 1).replace([np.inf, -np.inf], np.nan) * 100


def _complete_window_mask(price_index: pd.Series, window: int) -> pd.Series:
    return (price_index.notna().rolling(window + 1, min_periods=window + 1).sum() == window + 1)


def calculate_rolling_performance(price_index: pd.Series, window: int,
                                  require_full_window: bool) -> pd.Series:
    if price_index.empty or price_index.isna().all() or window <= 0:
        return pd.Series(np.nan, index=price_index.index, name=price_index.name)
    period_return = price_index.pct_change(periods=window, fill_method=None)
    if require_full_window:
        period_return = period_return.where(_complete_window_mask(price_index, window))
    return ((1 + period_return) ** (12.0 / window) - 1) * 100


def calculate_rolling_sharpe_like_ratio(returns: pd.Series, price_index: pd.Series, window: int,
                                        require_full_window: bool) -> pd.Series:
    """Annualised rolling return divided by annualised rolling volatility.

    No risk-free rate is deducted, so this is Sharpe-like rather than a Sharpe
    ratio. Module 1 produces the true excess-return Sharpe.
    """
    if (price_index.empty or price_index.isna().all()
            or returns.empty or returns.isna().all() or window <= 0):
        return pd.Series(np.nan, index=price_index.index, name=price_index.name)
    period_return = price_index.pct_change(periods=window, fill_method=None)
    if require_full_window:
        period_return = period_return.where(_complete_window_mask(price_index, window))
    annualised_return = (1 + period_return) ** (12.0 / window) - 1
    annualised_vol = returns.rolling(
        window=window, min_periods=window if require_full_window else 2
    ).std() * np.sqrt(12)
    return (annualised_return / annualised_vol).replace([np.inf, -np.inf], np.nan)


def generate_heatmap_figure(frame: pd.DataFrame, base_title: str, chart_title: str,
                            colour_label: str, colormap: Any, low: float | None, high: float | None,
                            centre: float | None, rolling_window: int | None,
                            wrap_width: int, annotation_format: str = ".1f") -> plt.Figure | None:
    if frame.empty:
        return None
    fund_count, month_count = frame.shape[1], frame.shape[0]
    figure, axis = plt.subplots(
        figsize=(max(10, 2 + fund_count * 0.5), max(7, min(30, 3 + month_count * 0.25)))
    )
    annotation_size = max(4, 8 - fund_count // 6)
    if month_count > 40:
        annotation_size = max(3, annotation_size - 1)
    if fund_count > 20:
        annotation_size = max(3, annotation_size - 1)
    if fund_count > 40:
        annotation_size = max(2, annotation_size - 1)
    sns.heatmap(
        frame, annot=True, fmt=annotation_format, linewidths=0.5, linecolor="lightgray",
        cmap=colormap, vmin=low, vmax=high, center=centre,
        cbar_kws={"label": colour_label, "shrink": 0.8}, annot_kws={"size": annotation_size},
        ax=axis, xticklabels=[textwrap.fill(str(name), width=wrap_width) for name in frame.columns],
    )
    axis.set_yticks(np.arange(month_count) + 0.5)
    axis.set_yticklabels([stamp.strftime("%Y-%m") for stamp in frame.index], rotation=0,
                         fontsize=max(7, 10 - month_count // 20))
    plt.setp(axis.get_xticklabels(), rotation=90, ha="center", va="top",
             fontsize=max(6, 9 - fund_count // 8))
    title = f"{base_title} - {chart_title}"
    if "Rolling" in chart_title and rolling_window is not None:
        title += f" ({rolling_window}-Month Window)"
    if "%" in colour_label and "%" not in title:
        title += " (%)"
    axis.set_title(title, fontsize=16, pad=25)
    axis.set_ylabel("Month End Date", fontsize=12)
    axis.set_xlabel("Fund", fontsize=12, labelpad=20)
    try:
        figure.tight_layout(pad=2.0)
    except ValueError:
        LOGGER.debug("tight_layout could not adjust the heatmap '%s'.", title)
    return figure


def create_legend_only_figure(handles: list[Line2D], labels: list[str],
                              chart_title: str) -> plt.Figure | None:
    """A separate legend page, used when too many funds to label on the chart."""
    if not handles or not labels:
        return None
    columns = 1 if len(labels) <= 15 else (2 if len(labels) <= 30 else 3)
    figure, axis = plt.subplots(figsize=(
        max(8, 4 + columns * 3),
        max(4, 2 + (len(labels) / columns) * 0.3 + (columns - 1) * 0.5),
    ))
    axis.axis("off")
    axis.legend(handles, labels, loc="center", fontsize="medium", ncol=columns, frameon=True,
                title=f"Legend for: {chart_title}", title_fontsize="large")
    try:
        figure.tight_layout(pad=1.0)
    except ValueError:
        LOGGER.debug("tight_layout could not adjust the legend page for '%s'.", chart_title)
    return figure


def plot_line_chart(frame: pd.DataFrame, title: str, ylabel: str,
                    shade_negative: bool = False) -> tuple[plt.Figure | None, plt.Figure | None]:
    if frame is None or frame.empty or frame.shape[1] == 0:
        return None, None
    series_count = frame.shape[1]
    extra_width = (series_count * 0.03
                   if 0 < series_count <= MAX_LEGEND_ITEMS_ON_CHART
                   and series_count > MAX_LEGEND_ITEMS_ON_CHART / 2 else 0)
    figure, axis = plt.subplots(figsize=(max(12, 12 + extra_width), 7))
    handles: list[Line2D] = []
    labels: list[str] = []
    for column in frame.columns:
        if frame[column].notna().any():
            line, = axis.plot(frame.index, frame[column], label=str(column), linewidth=1.5)
            handles.append(line)
            labels.append(str(column))
    if not handles:
        plt.close(figure)
        return None, None

    axis.set_title(title, fontsize=16, pad=20)
    axis.set_xlabel("Date", fontsize=12)
    axis.set_ylabel(ylabel, fontsize=12)
    if pd.api.types.is_datetime64_any_dtype(frame.index):
        figure.autofmt_xdate()
    if shade_negative:
        axis.axhline(0, color="black", linestyle="--", linewidth=0.8)
        for column in frame.columns:
            values = pd.to_numeric(frame[column], errors="coerce")
            if values.notna().any():
                axis.fill_between(frame.index, values, 0, where=(values < 0),
                                  alpha=0.1, interpolate=True)

    legend_figure = None
    if series_count <= MAX_LEGEND_ITEMS_ON_CHART:
        axis.legend(handles, labels, loc="best", fontsize="small")
    else:
        legend_figure = create_legend_only_figure(handles, labels, title)
    axis.grid(True, linestyle=":", alpha=0.7)
    try:
        figure.tight_layout()
    except ValueError:
        LOGGER.debug("tight_layout could not adjust the line chart '%s'.", title)
    return figure, legend_figure


def plot_line_chart_plotly(frame: pd.DataFrame, title: str, ylabel: str) -> Any | None:
    try:
        import plotly.graph_objects as go
    except ImportError:
        return None
    if frame is None or frame.empty or frame.shape[1] == 0:
        return None
    figure = go.Figure()
    for column in frame.columns:
        if frame[column].notna().any():
            figure.add_trace(go.Scatter(
                x=frame.index, y=frame[column], mode="lines", name=str(column),
                hovertemplate=f"<b>{column}</b><br>%{{x|%Y-%m-%d}}<br>{ylabel}: %{{y:.2f}}<extra></extra>",
            ))
    if not figure.data:
        return None
    figure.update_layout(title_text=title, xaxis_title="Date", yaxis_title=ylabel,
                         legend_title_text="Funds", hovermode="closest", height=600)
    if "Drawdown" in title:
        figure.add_hline(y=0, line_dash="dash", line_color="black", line_width=1)
    return figure


def create_excluded_funds_page(excluded: list[tuple[str, str, int, str]], min_history: int,
                               plot_window: int) -> plt.Figure:
    """A PDF page listing every fund left off the line charts, and why."""
    if not excluded:
        figure, axis = plt.subplots(figsize=(10, 2))
        axis.text(0.5, 0.5, "No funds were left out of the line charts.",
                  ha="center", va="center", fontsize=12)
        axis.axis("off")
        return figure

    table_frame = pd.DataFrame(
        excluded,
        columns=["Fund Name", "Strategy Context", "Months of Data", "Reason for Exclusion"],
    ).sort_values(by=["Strategy Context", "Fund Name"])
    height = max(5, min(25, 2 + len(table_frame) * 0.30 + 1.5))
    figure, axis = plt.subplots(figsize=(16, height))
    axis.axis("tight")
    axis.axis("off")
    table = axis.table(cellText=table_frame.values, colLabels=table_frame.columns, loc="center",
                       cellLoc="left", colColours=["#E0E0E0"] * len(table_frame.columns))
    table.set_fontsize(8)
    table.scale(1.0, 1.5)
    for column, width in enumerate([0.20, 0.15, 0.10, 0.55]):
        for row in range(len(table_frame) + 1):
            try:
                table[(row, column)].set_width(width)
            except KeyError:
                pass
    axis.set_title("Funds Not Included in Line Charts", fontsize=16, pad=20)

    history_text = (f"Funds need at least {min_history} months of data."
                    if min_history > 0 else "There is no minimum track-record filter.")
    window_text = (f" Charts show the most recent {plot_window} months."
                   if plot_window > 0 else " Charts show all available history.")
    figure.text(
        0.5, 0.03,
        history_text + window_text
        + "\n'Months of Data' counts the fund's reported months in the workbook."
        + "\nEvery fund with usable returns appears in the heatmaps, whatever its history length.",
        ha="center", va="bottom", fontsize=9, wrap=True,
        bbox=dict(boxstyle="round,pad=0.5", fc="aliceblue", ec="lightgrey", lw=0.5),
    )
    try:
        figure.tight_layout(rect=(0, 0.08, 1, 0.92))
    except ValueError:
        LOGGER.debug("tight_layout could not adjust the excluded-funds page.")
    return figure


def prepare_line_chart_data(series_by_name: dict[str, pd.Series], min_history: int,
                            plot_window: int, excluded: list[tuple[str, str, int, str]],
                            context: str, transform: str) -> pd.DataFrame | None:
    """Filter fund histories and build the frame behind one line chart.

    NAV is deliberately rebased at the start of the displayed window, so the
    funds on a chart are comparable. Drawdown is calculated from each fund's
    whole history before the window is applied, so the high-water mark is real.
    """
    if not series_by_name:
        return None
    prepared: dict[str, pd.Series] = {}
    for name, returns in series_by_name.items():
        observed = pd.to_numeric(returns, errors="coerce").dropna()
        months = len(observed)
        if min_history > 0 and months < min_history:
            excluded.append((name, context, months,
                             f"Only {months} months of data; {min_history} required"))
            continue
        if plot_window > 0 and months < plot_window:
            excluded.append((name, context, months,
                             f"Only {months} months of data; the {plot_window}-month "
                             "chart window requires more"))
            continue
        if transform == "drawdown":
            values = calculate_actual_drawdowns(calculate_price_index(observed))
            values = values.tail(plot_window) if plot_window > 0 else values
        else:
            window = observed.tail(plot_window) if plot_window > 0 else observed
            values = calculate_price_index(window)
        if values is None or values.dropna().empty:
            excluded.append((name, context, months,
                             "No values could be calculated for the selected window"))
            continue
        prepared[name] = values
    return pd.DataFrame(prepared) if prepared else None


def _tail_months(frame: pd.DataFrame, months: int) -> pd.DataFrame:
    """The most recent months of a frame; 0 means keep everything."""
    if frame.empty:
        return frame
    ordered = frame.sort_index()
    return ordered if months <= 0 or len(ordered) <= months else ordered.tail(months)


def _performance_heatmaps(returns_by_name: dict[str, pd.Series], options: dict[str, Any]
                          ) -> dict[str, pd.DataFrame]:
    """The four heatmap frames, all built from each fund's complete history."""
    window = int(options["rolling_window_months"])
    require_full = bool(options["require_full_rolling_window"])
    results = {key: {} for key in ("actual_dd", "rolling_dd", "rolling_performance", "ratio")}
    for name, returns in returns_by_name.items():
        price_index = calculate_price_index(returns)
        if price_index.dropna().empty:
            continue
        results["actual_dd"][name] = calculate_actual_drawdowns(price_index)
        results["rolling_dd"][name] = calculate_rolling_drawdowns(price_index, window, require_full)
        results["rolling_performance"][name] = calculate_rolling_performance(
            price_index, window, require_full)
        results["ratio"][name] = calculate_rolling_sharpe_like_ratio(
            returns, price_index, window, require_full)
    frames = {}
    for key, columns in results.items():
        frame = pd.DataFrame(columns)
        if not frame.empty:
            frame = frame.dropna(axis=0, how="all").dropna(axis=1, how="all")
        frames[key] = frame
    return frames


def run_performance_module(data: WorkbookData, options: dict[str, Any],
                           output_dir: Path) -> list[str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    log: list[str] = []
    colormaps = _build_colormaps()
    # Every figure is held open until the end, because the PDF pages have to be
    # written in a fixed order that puts the excluded-funds page first, and that
    # page can only be built once every chart has reported its exclusions.
    plt.rcParams["figure.max_open_warning"] = 0
    window = int(options["rolling_window_months"])
    min_history = int(options["min_history_months"])
    plot_window = int(options["line_chart_window_months"])
    heatmap_months = int(options["recent_months_heatmap"])
    wrap_width = int(options["x_label_wrap_width"])
    dpi = int(options["dpi"])
    want_html = options["line_chart_mode"] == "html"
    want_static = (not want_html) or options["create_pdf"] or options["save_png"]

    if want_html and importlib.util.find_spec("plotly") is None:
        want_html = False
        want_static = True
        log.append("Plotly is not installed, so static line charts were produced instead.")

    figures: list[tuple[plt.Figure, int, str]] = []
    html_files: list[str] = []
    excluded: list[tuple[str, str, int, str]] = []
    all_returns: dict[str, pd.Series] = {}
    combined_heatmaps: dict[str, list[pd.DataFrame]] = {
        key: [] for key in ("actual_dd", "rolling_dd", "rolling_performance", "ratio")}
    sort_key = 10

    history_title = (f"Min {min_history}m History" if min_history > 0 else "No Minimum History")
    window_title = (f"Last {plot_window}m" if plot_window > 0 else "All Available History")
    title_suffix = f"({history_title}; {window_title})"

    heatmap_specs = [
        ("rolling_dd", "Rolling Drawdowns", "Rolling Drawdown (%)", "drawdown", window, ".1f"),
        ("actual_dd", "Actual Drawdowns", "Actual Drawdown (%)", "drawdown", None, ".1f"),
        ("rolling_performance", "Annualized Rolling Performance",
         "Annualized Rolling Performance (%)", "performance", window, ".1f"),
        ("ratio", "Annualized Rolling Risk-Adjusted Return",
         "Risk-Adjusted Return (Ann.Ret/Ann.Vol)", "ratio", window, ".2f"),
    ]

    def add_line_charts(returns_by_name: dict[str, pd.Series], context: str,
                        chart_base: str, start_key: int) -> int:
        """NAV and drawdown line charts for one strategy, or for every fund."""
        key = start_key
        for transform, ylabel, shade, title in (
            ("nav", f"NAV (Initial {INITIAL_PRICE_INDEX_VALUE:g})", False,
             f"{chart_base} - NAV Comparison (Growth of {INITIAL_PRICE_INDEX_VALUE:g}) {title_suffix}"),
            ("drawdown", "Drawdown (%)", True, f"{chart_base} - Actual Drawdowns {title_suffix}"),
        ):
            frame = prepare_line_chart_data(returns_by_name, min_history, plot_window,
                                            excluded, context, transform)
            if frame is None or frame.empty:
                log.append(f"{chart_base}: no {transform} line chart - no fund met the criteria.")
                continue
            if want_html:
                interactive = plot_line_chart_plotly(frame, title, ylabel)
                if interactive is not None:
                    filename = output_dir / (
                        f"{'NAV' if transform == 'nav' else 'Drawdown'}_"
                        f"{safe_file_stem(chart_base)}.html")
                    interactive.write_html(str(filename))
                    html_files.append(str(filename))
            if want_static:
                chart, legend = plot_line_chart(frame, title, ylabel, shade_negative=shade)
                if chart is not None:
                    figures.append((chart, key, f"{transform.upper()} Line Plot - {chart_base}"))
                    key += 1
                if legend is not None:
                    figures.append((legend, key, f"Legend for {transform.upper()} Line Plot - {chart_base}"))
                    key += 1
        return key

    for strategy in data.strategies:
        returns_by_name = {fund.fund: fund.returns for fund in data.funds
                           if fund.strategy == strategy}
        for name, returns in returns_by_name.items():
            all_returns[f"{strategy} - {name}"] = returns

        sort_key = add_line_charts(returns_by_name, strategy, strategy, sort_key) + 10

        frames = _performance_heatmaps(returns_by_name, options)
        for key, chart_title, colour_label, palette, rolling, annotation in heatmap_specs:
            frame = _tail_months(frames[key], heatmap_months)
            if frame.empty:
                continue
            low, high = colormaps[f"{palette}_limits"]
            figure = generate_heatmap_figure(
                frame, strategy, chart_title, colour_label, colormaps[palette], low, high,
                0.0 if palette != "drawdown" else None, rolling, wrap_width, annotation)
            if figure is not None:
                figures.append((figure, sort_key, f"{chart_title} Heatmap - {strategy}"))
            sort_key += 1
        for key in combined_heatmaps:
            if not frames[key].empty:
                combined_heatmaps[key].append(
                    frames[key].rename(columns=lambda name: f"{strategy} - {name}"))
        sort_key = ((sort_key // 100) + 1) * 100

    combined_title = "All Funds (Combined Strategies)"
    if all_returns:
        sort_key = add_line_charts(all_returns, "All Funds (Combined)", combined_title, sort_key) + 10
        for key, chart_title, colour_label, palette, rolling, annotation in heatmap_specs:
            if not combined_heatmaps[key]:
                continue
            merged = pd.concat(combined_heatmaps[key], axis=1).sort_index().dropna(axis=0, how="all")
            frame = _tail_months(merged, heatmap_months)
            if frame.empty:
                continue
            low, high = colormaps[f"{palette}_limits"]
            figure = generate_heatmap_figure(
                frame, combined_title, chart_title, colour_label, colormaps[palette], low, high,
                0.0 if palette != "drawdown" else None, rolling, wrap_width, annotation)
            if figure is not None:
                figures.append((figure, sort_key, f"{chart_title} Heatmap - {combined_title}"))
            sort_key += 1

    # One row per fund/reason, so a fund excluded from both the strategy chart
    # and the combined chart is only listed once.
    seen: set[tuple[str, str]] = set()
    unique_excluded: list[tuple[str, str, int, str]] = []
    for entry in excluded:
        name, context, months, reason = entry
        identity = (name.split(" - ", 1)[-1] if context == "All Funds (Combined)" else name, reason)
        if identity not in seen:
            seen.add(identity)
            unique_excluded.append(entry)
    figures.insert(0, (create_excluded_funds_page(unique_excluded, min_history, plot_window),
                       0, "Funds Excluded From Line Charts"))

    if html_files:
        log.append(f"Saved {len(html_files)} interactive HTML line chart(s).")

    figures.sort(key=lambda item: item[1])
    if options["create_pdf"] and figures:
        pdf_name = (
            f"{_dt.date.today():%Y-%m-%d} - Hedge Fund Analysis - "
            f"{'all' if heatmap_months <= 0 else heatmap_months}m_heatmaps - "
            f"{window}m_rolling - {min_history}m_min_history - {plot_window}m_line_window.pdf"
        )
        pdf_path = output_dir / pdf_name
        try:
            with PdfPages(pdf_path) as pdf:
                for figure, _, _ in figures:
                    pdf.savefig(figure, bbox_inches="tight", pad_inches=0.1)
            log.append(f"Compiled PDF: {pdf_path}")
        except PermissionError:
            log.append(f"The PDF could not be written to {pdf_path}. Close it and run again.")

    if options["save_png"] and figures:
        for figure, key, description in figures:
            filename = output_dir / f"{key:04d}_{safe_file_stem(description)}.png"
            figure.savefig(filename, dpi=dpi, bbox_inches="tight")
        log.append(f"Saved {len(figures)} PNG chart(s).")

    for figure, _, _ in figures:
        plt.close(figure)
    log.append(f"Module 3 output folder: {output_dir}")
    return log


# =============================================================================
# Orchestration
# =============================================================================

def selected_modules(settings: dict[str, Any]) -> list[str]:
    return [key for key in MODULE_KEYS if settings["modules"].get(key)]


def check_workbook(settings: dict[str, Any]) -> WorkbookData:
    """Load and check the workbook, then confirm it suits the chosen modules.

    Anything wrong stops the run before a single chart is produced, and the
    message lists every problem at once.
    """
    workbook = settings["workbook"]
    sheet_text = str(workbook["sheet"]).strip()
    data = load_workbook(
        Path(str(workbook["path"])).expanduser(),
        int(sheet_text) if sheet_text.isdigit() else sheet_text,
        stop_on_missing_months=bool(workbook["stop_on_missing_months"]),
        stop_on_history_gaps=bool(workbook["stop_on_history_gaps"]),
    )
    problems = check_modules_against_layout(data.layout, selected_modules(settings))
    if problems:
        raise WorkbookFormatError("\n\n".join(problems))
    return data


def run_analysis(settings: dict[str, Any], data: WorkbookData | None = None) -> list[str]:
    """Run every selected module and return the lines of the run report."""
    validate_settings(settings)
    if data is None:
        data = check_workbook(settings)

    output_root = Path(str(settings["workbook"]["output_dir"])).expanduser()
    if not output_root.is_absolute():
        output_root = Path.cwd() / output_root
    if settings["workbook"]["timestamped_run_folder"]:
        output_root = output_root / _dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
    output_root.mkdir(parents=True, exist_ok=True)

    report: list[str] = [
        f"Hedge fund analytics run - {_dt.datetime.now():%Y-%m-%d %H:%M:%S}",
        "",
        *data.summary_lines(),
        f"Output folder: {output_root}",
        "",
    ]
    if data.notes:
        report.append("Data notes")
        report.extend(f"  - {note}" for note in data.notes)
        report.append("")

    runners: dict[str, Callable[[WorkbookData, dict[str, Any], Path], list[str]]] = {
        "cone": run_cone_module,
        "clustering": run_clustering_module,
        "performance": run_performance_module,
    }
    for key in selected_modules(settings):
        report.append(f"--- {MODULE_TITLES[key]} ---")
        LOGGER.info("Running %s", MODULE_TITLES[key])
        try:
            report.extend(f"  {line}" for line in
                          runners[key](data, settings[key], output_root / MODULE_FOLDERS[key]))
        except Exception as exc:  # keep the other modules going, and say what failed
            LOGGER.exception("%s failed", MODULE_TITLES[key])
            report.append(f"  FAILED: {exc}")
        report.append("")

    report.append("Settings used")
    report.extend(f"  {line}" for line in
                  json.dumps(settings, indent=2, sort_keys=True).splitlines())
    (output_root / "run_report.txt").write_text("\n".join(report) + "\n", encoding="utf-8")
    report.insert(0, "")
    report.append("")
    report.append(f"Run report saved to {output_root / 'run_report.txt'}")
    return report


# =============================================================================
# Settings window
# =============================================================================

def show_settings_window(settings: dict[str, Any], path: Path) -> dict[str, Any] | None:
    """Show every option in one window; return the chosen settings and workbook.

    The result is {"settings": ..., "data": ...} - the checked workbook is
    handed back so the run does not have to read and check it a second time.
    None means the user cancelled. The settings are saved to the JSON file
    before the window closes, so the next run starts from the same place.
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("Hedge Fund Analytics - Settings")
    root.minsize(940, 720)
    try:
        root.option_add("*Font", ("Segoe UI", 10))
    except tk.TclError:
        pass

    chosen: list[dict[str, Any]] = []
    detected_layout: list[str] = []

    variables: dict[str, Any] = {}

    def string_var(section: str, key: str, transform: Callable[[Any], str] = str) -> tk.StringVar:
        variable = tk.StringVar(value=transform(settings[section][key]))
        variables[f"{section}.{key}"] = variable
        return variable

    def boolean_var(section: str, key: str) -> tk.BooleanVar:
        variable = tk.BooleanVar(value=bool(settings[section][key]))
        variables[f"{section}.{key}"] = variable
        return variable

    outer = ttk.Frame(root, padding=12)
    outer.pack(fill="both", expand=True)
    ttk.Label(outer, text="Hedge Fund Analytics", font=("Segoe UI", 15, "bold")).pack(anchor="w")
    ttk.Label(
        outer,
        text=("Tick the modules to run, point at the workbook and adjust anything you need. "
              "Every setting is saved and restored automatically."),
        wraplength=900,
    ).pack(anchor="w", pady=(2, 10))

    notebook = ttk.Notebook(outer)
    notebook.pack(fill="both", expand=True)
    run_tab = ttk.Frame(notebook, padding=12)
    cone_tab = ttk.Frame(notebook, padding=12)
    cluster_tab = ttk.Frame(notebook, padding=12)
    performance_tab = ttk.Frame(notebook, padding=12)
    notebook.add(run_tab, text="Workbook & modules")
    notebook.add(cone_tab, text=MODULE_TITLES["cone"])
    notebook.add(cluster_tab, text=MODULE_TITLES["clustering"])
    notebook.add(performance_tab, text=MODULE_TITLES["performance"])

    def labelled_entry(parent: ttk.Frame, row: int, label: str, variable: tk.StringVar,
                       hint: str = "", width: int | None = None) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)
        entry = ttk.Entry(parent, textvariable=variable, width=width or 24)
        entry.grid(row=row, column=1, sticky="ew" if width is None else "w", pady=4)
        if hint:
            ttk.Label(parent, text=hint, foreground="#555555").grid(
                row=row, column=2, sticky="w", padx=(10, 0), pady=4)

    def labelled_combo(parent: ttk.Frame, row: int, label: str, variable: tk.StringVar,
                       values: Sequence[str], hint: str = "") -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)
        ttk.Combobox(parent, textvariable=variable, values=list(values), state="readonly",
                     width=22).grid(row=row, column=1, sticky="w", pady=4)
        if hint:
            ttk.Label(parent, text=hint, foreground="#555555").grid(
                row=row, column=2, sticky="w", padx=(10, 0), pady=4)

    # ---- Workbook & modules tab -------------------------------------------
    run_tab.columnconfigure(1, weight=1)
    excel_var = string_var("workbook", "path")
    sheet_var = string_var("workbook", "sheet")
    output_var = string_var("workbook", "output_dir")

    def browse_workbook() -> None:
        selected = filedialog.askopenfilename(
            parent=root, title="Select the returns workbook",
            initialdir=str(Path(excel_var.get()).parent) if excel_var.get().strip() else None,
            filetypes=[("Excel workbooks", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")],
        )
        if selected:
            excel_var.set(selected)

    def browse_output() -> None:
        selected = filedialog.askdirectory(parent=root, title="Select the output folder",
                                           initialdir=output_var.get() or None)
        if selected:
            output_var.set(selected)

    ttk.Label(run_tab, text="Excel workbook").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=4)
    excel_entry = ttk.Entry(run_tab, textvariable=excel_var)
    excel_entry.grid(row=0, column=1, sticky="ew", pady=4)
    ttk.Button(run_tab, text="Browse...", command=browse_workbook).grid(
        row=0, column=2, padx=(8, 0), pady=4)
    ttk.Label(run_tab, text="Worksheet").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=4)
    sheet_combo = ttk.Combobox(run_tab, textvariable=sheet_var, width=34)
    sheet_combo.grid(row=1, column=1, sticky="w", pady=4)
    sheet_status = ttk.Label(run_tab, text="", foreground="#555555")
    sheet_status.grid(row=1, column=2, sticky="w", padx=(10, 0), pady=4)

    listed_workbook: list[str] = []

    def refresh_sheet_list(force: bool = False) -> None:
        """Fill the worksheet dropdown with the tabs in the chosen workbook.

        The list is rebuilt whenever the workbook changes, so the tab is picked
        rather than typed. When the workbook cannot be read the dropdown stays
        editable, so a name can still be entered by hand.
        """
        path_text = excel_var.get().strip().strip('"')
        if not force and listed_workbook and listed_workbook[0] == path_text:
            return
        listed_workbook.clear()

        def editable(message: str) -> None:
            sheet_combo.configure(values=(), state="normal")
            sheet_status.configure(text=message)

        if not path_text:
            editable("Choose a workbook to list its tabs.")
            return
        workbook = Path(path_text).expanduser()
        if not workbook.is_file():
            editable("Workbook not found - the tab name can be typed instead.")
            return
        try:
            names = list_sheet_names(workbook)
        except Exception as exc:
            editable(f"Tabs could not be listed ({exc.__class__.__name__}) - type the name instead.")
            return
        if not names:
            editable("This workbook has no worksheets.")
            return
        listed_workbook.append(path_text)
        sheet_combo.configure(values=names, state="readonly")
        sheet_var.set(resolve_sheet_choice(names, sheet_var.get()))
        sheet_status.configure(
            text=f"{len(names)} tab{'' if len(names) == 1 else 's'} in this workbook.")

    # Watching the variable catches every way the path can change - the Browse
    # dialog, typing, pasting - where a focus or Return binding would miss some.
    excel_var.trace_add("write", lambda *_args: refresh_sheet_list())
    ttk.Label(run_tab, text="Output folder").grid(row=2, column=0, sticky="w", padx=(0, 10), pady=4)
    ttk.Entry(run_tab, textvariable=output_var).grid(row=2, column=1, sticky="ew", pady=4)
    ttk.Button(run_tab, text="Browse...", command=browse_output).grid(
        row=2, column=2, padx=(8, 0), pady=4)

    modules_box = ttk.LabelFrame(run_tab, text="Modules to run", padding=10)
    modules_box.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 6))
    module_vars = {key: boolean_var("modules", key) for key in MODULE_KEYS}
    cone_check = ttk.Checkbutton(modules_box, text=MODULE_TITLES["cone"] +
                                 "  (needs the cone layout: expected figures on rows 3-4)",
                                 variable=module_vars["cone"])
    cone_check.grid(row=0, column=0, sticky="w", pady=2)
    ttk.Checkbutton(modules_box, text=MODULE_TITLES["clustering"] +
                    "  (either layout)", variable=module_vars["clustering"]).grid(
        row=1, column=0, sticky="w", pady=2)
    ttk.Checkbutton(modules_box, text=MODULE_TITLES["performance"] +
                    "  (either layout)", variable=module_vars["performance"]).grid(
        row=2, column=0, sticky="w", pady=2)

    checks_box = ttk.LabelFrame(run_tab, text="Data checks and output", padding=10)
    checks_box.grid(row=4, column=0, columnspan=3, sticky="ew", pady=6)
    ttk.Checkbutton(checks_box, text="Stop if a month is missing from the date column",
                    variable=boolean_var("workbook", "stop_on_missing_months")).grid(
        row=0, column=0, sticky="w", pady=2)
    ttk.Checkbutton(checks_box, text="Stop if a fund has blank months inside its live history",
                    variable=boolean_var("workbook", "stop_on_history_gaps")).grid(
        row=1, column=0, sticky="w", pady=2)
    ttk.Checkbutton(checks_box, text="Put each run in its own date-and-time sub-folder",
                    variable=boolean_var("workbook", "timestamped_run_folder")).grid(
        row=2, column=0, sticky="w", pady=2)

    check_box = ttk.LabelFrame(run_tab, text="Workbook format check", padding=10)
    check_box.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(6, 0))
    run_tab.rowconfigure(5, weight=1)
    check_box.columnconfigure(0, weight=1)
    check_box.rowconfigure(1, weight=1)
    ttk.Button(check_box, text="Check the workbook now",
               command=lambda: perform_check()).grid(row=0, column=0, sticky="w", pady=(0, 6))
    check_output = tk.Text(check_box, height=9, wrap="word")
    check_output.grid(row=1, column=0, sticky="nsew")
    check_scroll = ttk.Scrollbar(check_box, orient="vertical", command=check_output.yview)
    check_scroll.grid(row=1, column=1, sticky="ns")
    check_output.configure(yscrollcommand=check_scroll.set, state="disabled")

    def write_check_output(text: str) -> None:
        check_output.configure(state="normal")
        check_output.delete("1.0", "end")
        check_output.insert("1.0", text)
        check_output.configure(state="disabled")

    # ---- Module 1 tab ------------------------------------------------------
    cone_tab.columnconfigure(1, weight=1)
    ttk.Label(cone_tab, text=LAYOUT_DESCRIPTIONS["cone"], foreground="#555555",
              wraplength=860).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
    ttk.Checkbutton(cone_tab, text="Subtract the risk-free return (uses the Risk Free column)",
                    variable=boolean_var("cone", "use_risk_free")).grid(
        row=1, column=0, columnspan=3, sticky="w", pady=3)

    period_box = ttk.LabelFrame(cone_tab, text="History to analyse", padding=8)
    period_box.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(6, 8))
    period_mode = tk.StringVar(value="fixed" if settings["cone"]["use_predetermined_months"] else "full")
    variables["cone.period_mode"] = period_mode
    ttk.Radiobutton(period_box, text="Each fund's full history", variable=period_mode,
                    value="full").grid(row=0, column=0, sticky="w", padx=(0, 18))
    ttk.Radiobutton(period_box, text="First", variable=period_mode, value="fixed").grid(
        row=0, column=1, sticky="w")
    ttk.Spinbox(period_box, from_=1, to=1200, width=7,
                textvariable=string_var("cone", "n_months")).grid(row=0, column=2, padx=5)
    ttk.Label(period_box, text="months from each fund's start").grid(row=0, column=3, sticky="w")

    labelled_entry(cone_tab, 3, "Rolling return window", string_var("cone", "rolling_window"),
                   "months; the middle chart", width=10)
    labelled_entry(cone_tab, 4, "Rolling Sharpe window", string_var("cone", "sharpe_window"),
                   "months; 36 = 3 years", width=10)
    labelled_entry(cone_tab, 5, "Rolling volatility window", string_var("cone", "volatility_window"),
                   "months; second axis of the rolling chart", width=10)
    labelled_entry(cone_tab, 6, "Drawdown threshold",
                   string_var("cone", "drawdown_threshold_pct"),
                   "percent; the drawdown table reports time spent below this", width=10)
    labelled_entry(cone_tab, 7, "Image resolution", string_var("cone", "dpi"), "DPI", width=10)
    ttk.Checkbutton(cone_tab, text="Use one common date axis for every fund",
                    variable=boolean_var("cone", "common_x_axis")).grid(
        row=8, column=0, columnspan=3, sticky="w", pady=3)
    ttk.Checkbutton(cone_tab, text="Create the compiled PDF (all funds, four charts a page)",
                    variable=boolean_var("cone", "create_pdf")).grid(
        row=9, column=0, columnspan=3, sticky="w", pady=3)

    # ---- Module 2 tab ------------------------------------------------------
    cluster_tab.columnconfigure(1, weight=1)
    ttk.Label(cluster_tab, text="Runs on either layout. The risk-free column is excluded, "
                                "because it is not a fund.",
              foreground="#555555", wraplength=860).grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
    labelled_entry(cluster_tab, 1, "Lookback periods",
                   string_var("clustering", "timeframes",
                              lambda values: ",".join(str(value) for value in values)),
                   'months and/or "max", e.g. 36,60,max')
    labelled_entry(cluster_tab, 2, "Minimum usable months",
                   string_var("clustering", "min_required_months"), width=10)
    labelled_combo(cluster_tab, 3, "Correlation method",
                   string_var("clustering", "correlation_method"),
                   ("pearson", "spearman", "kendall"))
    labelled_combo(cluster_tab, 4, "Linkage method", string_var("clustering", "linkage_method"),
                   ("ward", "average", "complete"),
                   "Ward needs Pearson or Spearman")
    labelled_entry(cluster_tab, 5, "Maximum clusters", string_var("clustering", "max_clusters"),
                   width=10)
    labelled_entry(cluster_tab, 6, "Dendrogram colour threshold",
                   string_var("clustering", "distance_threshold",
                              lambda value: "" if value is None else str(value)),
                   "leave blank for automatic colouring", width=10)
    labelled_entry(cluster_tab, 7, "Annotate heatmaps up to",
                   string_var("clustering", "max_annotated_funds"),
                   "funds; larger grids omit the numbers", width=10)
    labelled_entry(cluster_tab, 8, "Fund-label wrap width",
                   string_var("clustering", "label_wrap_width"), "characters", width=10)
    labelled_entry(cluster_tab, 9, "Combined-label wrap width",
                   string_var("clustering", "all_strategies_label_wrap_width"),
                   "characters; the all-strategies charts", width=10)
    labelled_entry(cluster_tab, 10, "Image resolution", string_var("clustering", "dpi"),
                   "DPI", width=10)
    ttk.Checkbutton(cluster_tab, text="Show correlation values inside the heatmap cells",
                    variable=boolean_var("clustering", "annotate_heatmaps")).grid(
        row=11, column=0, columnspan=3, sticky="w", pady=3)
    ttk.Checkbutton(cluster_tab,
                    text="Only include funds that reported in the latest month",
                    variable=boolean_var("clustering", "require_latest_return")).grid(
        row=12, column=0, columnspan=3, sticky="w", pady=3)
    cluster_outputs = ttk.LabelFrame(cluster_tab, text="Files to create", padding=8)
    cluster_outputs.grid(row=13, column=0, columnspan=3, sticky="ew", pady=(8, 0))
    ttk.Checkbutton(cluster_outputs, text="PDF report per lookback",
                    variable=boolean_var("clustering", "create_pdf")).grid(row=0, column=0, sticky="w")
    ttk.Checkbutton(cluster_outputs, text="PNG charts",
                    variable=boolean_var("clustering", "save_png")).grid(
        row=0, column=1, sticky="w", padx=(20, 0))
    ttk.Checkbutton(cluster_outputs, text="CSV correlations, clusters and audit",
                    variable=boolean_var("clustering", "save_csv")).grid(
        row=0, column=2, sticky="w", padx=(20, 0))

    # ---- Module 3 tab ------------------------------------------------------
    performance_tab.columnconfigure(1, weight=1)
    ttk.Label(performance_tab,
              text="Runs on either layout. The risk-free column is excluded, so the "
                   "risk-adjusted return here is return over volatility, with no risk-free "
                   "deduction. Module 1 produces the true excess-return Sharpe.",
              foreground="#555555", wraplength=860).grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
    labelled_entry(performance_tab, 1, "Rolling window",
                   string_var("performance", "rolling_window_months"),
                   "months; normally 12", width=10)
    labelled_entry(performance_tab, 2, "Minimum fund history",
                   string_var("performance", "min_history_months"),
                   "months; 0 = no filter", width=10)
    labelled_entry(performance_tab, 3, "Line-chart history",
                   string_var("performance", "line_chart_window_months"),
                   "months; 0 = all available", width=10)
    labelled_entry(performance_tab, 4, "Heatmap history",
                   string_var("performance", "recent_months_heatmap"),
                   "months; 0 = all available", width=10)
    labelled_entry(performance_tab, 5, "Fund-name wrap width",
                   string_var("performance", "x_label_wrap_width"), "characters", width=10)
    labelled_combo(performance_tab, 6, "Line-chart format",
                   string_var("performance", "line_chart_mode"), ("html", "static"),
                   "html = interactive, needs plotly")
    labelled_entry(performance_tab, 7, "Image resolution", string_var("performance", "dpi"),
                   "DPI", width=10)
    ttk.Checkbutton(performance_tab,
                    text="Require complete rolling periods (recommended for fair comparisons)",
                    variable=boolean_var("performance", "require_full_rolling_window")).grid(
        row=8, column=0, columnspan=3, sticky="w", pady=3)
    performance_outputs = ttk.LabelFrame(performance_tab, text="Files to create", padding=8)
    performance_outputs.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(8, 0))
    ttk.Checkbutton(performance_outputs, text="Combined PDF report",
                    variable=boolean_var("performance", "create_pdf")).grid(row=0, column=0, sticky="w")
    ttk.Checkbutton(performance_outputs, text="PNG charts",
                    variable=boolean_var("performance", "save_png")).grid(
        row=0, column=1, sticky="w", padx=(20, 0))

    # ---- Collecting and validating ----------------------------------------
    def whole_number(section: str, key: str, label: str) -> int:
        text = variables[f"{section}.{key}"].get().strip()
        try:
            return int(text)
        except ValueError as exc:
            raise ValueError(f"{label} must be a whole number (found '{text}').") from exc

    def collect() -> dict[str, Any]:
        threshold_text = variables["clustering.distance_threshold"].get().strip()
        try:
            drawdown_threshold = float(variables["cone.drawdown_threshold_pct"].get().strip())
        except ValueError as exc:
            raise ValueError("The cone drawdown threshold must be a number, such as 5.") from exc
        collected = {
            "workbook": {
                "path": excel_var.get().strip().strip('"'),
                "sheet": sheet_var.get().strip(),
                "output_dir": output_var.get().strip().strip('"'),
                "timestamped_run_folder": bool(variables["workbook.timestamped_run_folder"].get()),
                "stop_on_missing_months": bool(variables["workbook.stop_on_missing_months"].get()),
                "stop_on_history_gaps": bool(variables["workbook.stop_on_history_gaps"].get()),
            },
            "modules": {key: bool(module_vars[key].get()) for key in MODULE_KEYS},
            "cone": {
                "use_risk_free": bool(variables["cone.use_risk_free"].get()),
                "use_predetermined_months": period_mode.get() == "fixed",
                "n_months": whole_number("cone", "n_months", "The fixed number of months"),
                "rolling_window": whole_number("cone", "rolling_window", "The rolling return window"),
                "sharpe_window": whole_number("cone", "sharpe_window", "The rolling Sharpe window"),
                "volatility_window": whole_number("cone", "volatility_window",
                                                  "The rolling volatility window"),
                "drawdown_threshold_pct": drawdown_threshold,
                "common_x_axis": bool(variables["cone.common_x_axis"].get()),
                "create_pdf": bool(variables["cone.create_pdf"].get()),
                "dpi": whole_number("cone", "dpi", "The cone image resolution"),
            },
            "clustering": {
                "timeframes": list(parse_timeframes(variables["clustering.timeframes"].get())),
                "min_required_months": whole_number("clustering", "min_required_months",
                                                    "The minimum usable months"),
                "require_latest_return": bool(variables["clustering.require_latest_return"].get()),
                "correlation_method": variables["clustering.correlation_method"].get(),
                "linkage_method": variables["clustering.linkage_method"].get(),
                "max_clusters": whole_number("clustering", "max_clusters", "The maximum clusters"),
                "distance_threshold": float(threshold_text) if threshold_text else None,
                "annotate_heatmaps": bool(variables["clustering.annotate_heatmaps"].get()),
                "max_annotated_funds": whole_number("clustering", "max_annotated_funds",
                                                    "The annotation limit"),
                "label_wrap_width": whole_number("clustering", "label_wrap_width",
                                                 "The fund-label wrap width"),
                "all_strategies_label_wrap_width": whole_number(
                    "clustering", "all_strategies_label_wrap_width",
                    "The combined-label wrap width"),
                "create_pdf": bool(variables["clustering.create_pdf"].get()),
                "save_png": bool(variables["clustering.save_png"].get()),
                "save_csv": bool(variables["clustering.save_csv"].get()),
                "dpi": whole_number("clustering", "dpi", "The clustering image resolution"),
            },
            "performance": {
                "rolling_window_months": whole_number("performance", "rolling_window_months",
                                                      "The rolling window"),
                "min_history_months": whole_number("performance", "min_history_months",
                                                   "The minimum fund history"),
                "line_chart_window_months": whole_number("performance", "line_chart_window_months",
                                                         "The line-chart history"),
                "recent_months_heatmap": whole_number("performance", "recent_months_heatmap",
                                                      "The heatmap history"),
                "x_label_wrap_width": whole_number("performance", "x_label_wrap_width",
                                                   "The fund-name wrap width"),
                "require_full_rolling_window": bool(
                    variables["performance.require_full_rolling_window"].get()),
                "line_chart_mode": variables["performance.line_chart_mode"].get(),
                "create_pdf": bool(variables["performance.create_pdf"].get()),
                "save_png": bool(variables["performance.save_png"].get()),
                "dpi": whole_number("performance", "dpi", "The performance image resolution"),
            },
        }
        validate_settings(collected)
        return collected

    def perform_check(quiet: bool = False) -> WorkbookData | None:
        try:
            candidate = collect()
        except ValueError as exc:
            write_check_output(str(exc))
            if not quiet:
                messagebox.showerror("Check the settings", str(exc), parent=root)
            return None
        try:
            data = check_workbook(candidate)
        except WorkbookFormatError as exc:
            detected_layout.clear()
            write_check_output(f"The workbook cannot be used as it stands.\n\n{exc}")
            if not quiet:
                messagebox.showerror("Workbook format", str(exc), parent=root)
            return None
        except Exception as exc:
            write_check_output(f"The workbook could not be read.\n\n{exc}")
            if not quiet:
                messagebox.showerror("Workbook", str(exc), parent=root)
            return None
        detected_layout.clear()
        detected_layout.append(data.layout)
        lines = ["The workbook passed every check.", ""] + data.summary_lines()
        if data.notes:
            lines += ["", "Notes:"] + [f"  - {note}" for note in data.notes]
        write_check_output("\n".join(lines))
        return data

    def reset_defaults() -> None:
        defaults = default_settings()
        for name, variable in variables.items():
            section, key = name.split(".", 1)
            if key == "period_mode":
                variable.set("fixed" if defaults["cone"]["use_predetermined_months"] else "full")
            elif key == "timeframes":
                variable.set(",".join(str(value) for value in defaults[section][key]))
            elif defaults[section][key] is None:
                variable.set("")
            else:
                variable.set(defaults[section][key])
        for key in MODULE_KEYS:
            module_vars[key].set(defaults["modules"][key])
        refresh_sheet_list(force=True)
        write_check_output("")

    def check_and_run() -> None:
        data = perform_check()
        if data is None:
            return
        candidate = collect()
        try:
            save_settings(path, candidate)
        except OSError as exc:
            if not messagebox.askyesno(
                "Settings could not be saved",
                f"{exc}\n\nRun the analysis anyway?", parent=root
            ):
                return
        chosen.append({"settings": candidate, "data": data})
        root.destroy()

    def cancel() -> None:
        root.destroy()

    ttk.Label(outer, text=f"Settings file: {path}", foreground="#555555").pack(
        anchor="w", pady=(10, 0))
    buttons = ttk.Frame(outer)
    buttons.pack(fill="x", pady=(8, 0))
    ttk.Button(buttons, text="Reset to defaults", command=reset_defaults).pack(side="left")
    ttk.Button(buttons, text="Check and run", command=check_and_run).pack(side="right")
    ttk.Button(buttons, text="Cancel", command=cancel).pack(side="right", padx=(0, 8))
    ttk.Button(buttons, text="Check workbook only",
               command=lambda: perform_check()).pack(side="right", padx=(0, 8))

    root.protocol("WM_DELETE_WINDOW", cancel)
    root.bind("<Escape>", lambda _event: cancel())
    refresh_sheet_list(force=True)
    root.update_idletasks()
    root.geometry(
        f"+{max(0, (root.winfo_screenwidth() - root.winfo_reqwidth()) // 2)}"
        f"+{max(0, (root.winfo_screenheight() - root.winfo_reqheight()) // 6)}"
    )
    root.mainloop()
    return chosen[0] if chosen else None


# =============================================================================
# Command line
# =============================================================================

def parse_arguments(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Hedge fund analytics: cone charts, correlation clustering and "
            "performance/drawdown analysis. Run with no arguments to open the settings window."
        ),
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    parser.add_argument("excel_path", nargs="?", help="Excel workbook to analyse")
    parser.add_argument("--sheet", help="Worksheet name, or a zero-based sheet number")
    parser.add_argument("--output-dir", help="Folder to write the results into")
    parser.add_argument(
        "--modules",
        help=("Comma-separated list from " + ", ".join(MODULE_KEYS)
              + ". Module 1 (cone) needs the cone layout."),
    )
    parser.add_argument("--check-only", action="store_true",
                        help="Check the workbook format and stop, without producing anything")
    parser.add_argument("--gui", action=argparse.BooleanOptionalAction, default=None,
                        help="Force the settings window on or off")
    parser.add_argument("--settings-file", help="Use a different settings JSON file")
    return parser.parse_args(list(argv))


def apply_arguments(settings: dict[str, Any], arguments: argparse.Namespace) -> dict[str, Any]:
    if arguments.excel_path:
        settings["workbook"]["path"] = arguments.excel_path
    if arguments.sheet is not None:
        settings["workbook"]["sheet"] = arguments.sheet
    if arguments.output_dir is not None:
        settings["workbook"]["output_dir"] = arguments.output_dir
    if arguments.modules:
        requested = {part.strip().lower() for part in arguments.modules.split(",") if part.strip()}
        unknown = requested - set(MODULE_KEYS)
        if unknown:
            raise ValueError(
                f"Unknown module(s): {', '.join(sorted(unknown))}. "
                f"Choose from {', '.join(MODULE_KEYS)}."
            )
        settings["modules"] = {key: key in requested for key in MODULE_KEYS}
    return settings


def main(argv: Sequence[str] | None = None) -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    argv = list(sys.argv[1:] if argv is None else argv)
    try:
        arguments = parse_arguments(argv)
    except SystemExit as exc:  # argparse already printed the reason
        return int(exc.code or 0)

    path = Path(arguments.settings_file).expanduser() if arguments.settings_file else settings_path()
    try:
        settings = apply_arguments(load_settings(path), arguments)
    except ValueError as exc:
        print(f"Error: {exc}")
        return 2

    try:
        if arguments.check_only:
            validate_settings(settings)
            data = check_workbook(settings)
            print("\n".join(["The workbook passed every check.", ""] + data.summary_lines()))
            for note in data.notes:
                print(f"  Note: {note}")
            return 0

        use_gui = arguments.gui is True or (arguments.gui is None and not argv)
        if use_gui:
            try:
                result = show_settings_window(settings, path)
            except ImportError:
                print(
                    "The settings window needs tkinter, which is not installed.\n"
                    "Install the Python tkinter component, or pass the workbook on the command "
                    "line with --no-gui."
                )
                return 2
            except Exception as exc:
                print(f"The settings window could not open ({exc}). Use --no-gui instead.")
                return 2
            if result is None:
                print("Cancelled.")
                return 0
            report = run_analysis(result["settings"], result["data"])
        else:
            validate_settings(settings)
            report = run_analysis(settings)
        print("\n".join(report))
        return 0
    except WorkbookFormatError as exc:
        print(f"\nThe workbook cannot be used as it stands.\n\n{exc}\n")
        return 1
    except ValueError as exc:
        print(f"Error: {exc}")
        return 2
    except KeyboardInterrupt:
        print("Cancelled.")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
