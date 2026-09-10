"""Hedge fund analytics suite - cone charts, correlation clustering,
performance/drawdown analysis and portfolio risk allocation, driven from a
single settings window.

    Module 1  Cone charts            expected-vs-actual excess return cones,
                                     rolling total return, rolling Sharpe,
                                     drawdown.
    Module 2  Correlation clustering correlation heatmaps, Ward dendrograms,
                                     cluster assignments, exclusion audit.
    Module 3  Performance/drawdowns  NAV and drawdown lines plus rolling
                                     drawdown, rolling performance and
                                     risk-adjusted-return heatmaps.
    Module 4  Portfolio risk         the portfolio's volatility split across
                                     module 2's clusters, the dendrogram with
                                     weights beside it, and a marginal
                                     add-or-trim chart. Needs a weights row.

Every run also writes a short commentary document beside the charts, explaining
what each output is and how to read it, and then what this particular workbook
shows - where the risk sits, where funds could be added or trimmed, and
anything else notable, with the funds named. It is a Word file where python-docx
is installed and Markdown where it is not.

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

PLAIN LAYOUT (modules 2, 3 and 4)

     A                    B                     C
 1   [any label]          Strategy name         Strategy name
 2   [any label]          Fund name             Fund name
 3   31/01/2020           1.20%                 0.80%
 4   29/02/2020          -0.40%                 0.25%
     ...                  ...                   ...

OPTIONAL WEIGHTS ROW (required by module 4, ignored by the rest)

A row labelled "Weight" in column A, directly above the first month, holding
each fund's share of the portfolio. "Current Position", "Position",
"Allocation", "Holding", "% of Portfolio" and "% of NAV" are accepted as well,
since a book's weights go by several names. It works with either layout and
pushes everything below it down one row, so a plain sheet then starts its
months on row 4 and a cone sheet on row 6:

     A                    B                     C
 1   [any label]          Strategy name         Strategy name
 2   [any label]          Fund name             Fund name
 3   Weight               6.00%                 4.00%
 4   31/01/2020           1.20%                 0.80%
     ...                  ...                   ...

A blank weight means the fund is not held, and it is then measured as a
candidate rather than left out. Weights may be percentage-formatted cells or
whole numbers, because their total tells the two apart. A total below 100% is
cash and above it is gearing; the risk-free column takes no weight.

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
Optional (the commentary as a Word document rather than Markdown):
    pip install python-docx
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

MODULE_KEYS = ("cone", "clustering", "performance", "allocation")
MODULE_TITLES = {
    "cone": "1 - Cone charts",
    "clustering": "2 - Correlation & clustering",
    "performance": "3 - Performance & drawdowns",
    "allocation": "4 - Portfolio risk allocation",
}
MODULE_FOLDERS = {
    "cone": "1_cone_charts",
    "clustering": "2_correlation_clustering",
    "performance": "3_performance_drawdowns",
    "allocation": "4_portfolio_risk_allocation",
}

ROW_STRATEGY = 0
ROW_FUNDNAME = 1
ROW_CONE_EXPECTED_EXCESS = 2
ROW_CONE_EXPECTED_VOL = 3
CONE_HEADER_ROWS = 4
PLAIN_HEADER_ROWS = 2
CONE_DATA_START_ROW = CONE_HEADER_ROWS
PLAIN_DATA_START_ROW = PLAIN_HEADER_ROWS

RISK_FREE_PREFIX = "risk free"

# The optional portfolio-weights row sits directly above the first month, and
# is recognised by its own label in column A rather than by position alone, so
# a row put in the wrong place is reported instead of being read as something
# else. Several names for the same row are accepted, because a book's weights
# are as often called its current positions or its allocation. Anything else
# there is still refused and named in the message, rather than being guessed
# at: a row of numbers above the first month could as easily be assets under
# management or a fee.
WEIGHTS_ROW_PREFIXES = (
    "weight", "current position", "position", "current allocation", "allocation",
    "holding", "% of portfolio", "portfolio %", "% of nav", "nav %",
)

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

WEIGHTS_ROW_DESCRIPTION = (
    'Optional weights row - a row labelled "Weight" in column A, directly above the first '
    "month, holding each fund's share of the portfolio. "
    '"Current Position", "Allocation", "Holding" and "% of Portfolio" are accepted too. '
    "Everything below it moves down one row, so the plain layout then starts its months on "
    "row 4 and the cone layout on row 6. A blank weight means the fund is not held, which is "
    "how a candidate is measured before it is bought. Module 4 needs this row; the other "
    "modules ignore it."
)


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
        "modules": {"cone": True, "clustering": True, "performance": True,
                    "allocation": False},
        # A gap smaller than material_gap_pct is noise in an estimate made from
        # sixty-odd monthly observations, and naming it would hand the reader a
        # decision the data cannot support. It is a setting because a
        # concentrated book of eight funds and a diversified one of forty do
        # not have the same idea of what counts as material.
        "commentary": {
            "create_document": True,
            "file_format": "auto",
            "embed_charts": True,
            "chart_dpi": 200,
            "max_named_funds": 6,
            "material_gap_pct": 3.0,
        },
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
        "allocation": {
            "covariance_method": "ledoit_wolf",
            "unsmooth_returns": False,
            "return_basis": "auto",
            "risk_multiplier_flag": 1.25,
            "label_wrap_width": 38,
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

    allocation = settings["allocation"]
    if allocation["covariance_method"] not in {"ledoit_wolf", "sample"}:
        raise ValueError(
            "Portfolio risk: the covariance estimator must be ledoit_wolf or sample."
        )
    if allocation["return_basis"] not in {"auto", "expected", "realised"}:
        raise ValueError(
            "Portfolio risk: the expected-return basis must be auto, expected or realised."
        )
    if float(allocation["risk_multiplier_flag"]) <= 1.0:
        raise ValueError(
            "Portfolio risk: the crowding flag must be above 1, since 1 is a cluster carrying "
            "exactly its share of the risk."
        )
    if int(allocation["label_wrap_width"]) < 5:
        raise ValueError("Portfolio risk: the label wrap width must be at least 5 characters.")
    if int(allocation["dpi"]) < 72:
        raise ValueError("Portfolio risk: image resolution must be at least 72 DPI.")
    if not any((allocation["create_pdf"], allocation["save_png"], allocation["save_csv"])):
        raise ValueError("Portfolio risk: select at least one output type (PDF, PNG or CSV).")

    commentary = settings["commentary"]
    if commentary["file_format"] not in {"auto", "docx", "markdown"}:
        raise ValueError(
            "Commentary: the file format must be auto, docx or markdown."
        )
    if int(commentary["chart_dpi"]) < 72:
        raise ValueError("Commentary: embedded chart resolution must be at least 72 DPI.")
    if int(commentary["max_named_funds"]) < 1:
        raise ValueError("Commentary: at least one fund must be named in each list.")
    if not 0 < float(commentary["material_gap_pct"]) < 100:
        raise ValueError(
            "Commentary: the materiality threshold must be between 0 and 100 percentage points."
        )

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
    weight: float | None = None

    @property
    def label(self) -> str:
        return f"{self.strategy} - {self.fund}"

    @property
    def combined_label(self) -> str:
        """The name used once every strategy is put on one chart."""
        return f"{self.strategy}::{self.fund}"


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
    weights_row: int | None = None
    notes: list[str] = field(default_factory=list)

    @property
    def has_weights(self) -> bool:
        return self.weights_row is not None

    def weights(self) -> pd.Series:
        """Every fund's weight on the combined labels, unheld funds included."""
        return pd.Series(
            {fund.combined_label: float(fund.weight or 0.0) for fund in self.funds},
            dtype=float,
        )

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
        if self.has_weights:
            weights = self.weights()
            held = weights[weights != 0.0]
            lines.append(
                f"Portfolio weights: row {self.weights_row + 1} - {len(held)} fund(s) held, "
                f"totalling {weights.sum() * 100:.1f}%; "
                f"{len(weights) - len(held)} unheld fund(s) measured as candidates."
            )
        else:
            lines.append(
                "Portfolio weights: no weights row, so module 4 cannot run."
            )
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


def detect_layout(raw: pd.DataFrame) -> tuple[Literal["cone", "plain"], int, int | None]:
    """Work out where the monthly rows start, and therefore which layout it is.

    Returns the layout, the first row of monthly data and the row holding
    portfolio weights, or None when the sheet has no weights row. The weights
    row is optional and adds one row to either layout, so the number of header
    rows above it - not the position of the first date on its own - is what
    identifies the layout.
    """
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
            + "\n\n"
            + WEIGHTS_ROW_DESCRIPTION
        )

    weights_row: int | None = None
    if first_date_row > 0:
        candidate = _header_text(raw.iloc[first_date_row - 1], 0).strip().lower()
        if candidate.startswith(WEIGHTS_ROW_PREFIXES):
            weights_row = first_date_row - 1

    header_rows = first_date_row if weights_row is None else weights_row
    if header_rows == CONE_HEADER_ROWS:
        return "cone", first_date_row, weights_row
    if header_rows == PLAIN_HEADER_ROWS:
        return "plain", first_date_row, weights_row

    expected = (
        "Expected row 3 (plain layout) or row 5 (cone layout), each one lower "
        "if a weights row is used."
    )
    stray = ""
    if weights_row is None and first_date_row > 0:
        label = _header_text(raw.iloc[first_date_row - 1], 0)
        if label:
            stray = (
                f"\n\nThe row above the first date has '{label}' in column A, which is not one "
                "of the names a weights row is recognised by. If that row holds portfolio "
                'weights, change column A to "Weight" and it will be read as one.'
            )
    raise WorkbookFormatError(
        f"The first date in column A is on Excel row {first_date_row + 1}, which matches "
        "neither supported layout.\n\n"
        + expected
        + "\n\n"
        + LAYOUT_DESCRIPTIONS["cone"]
        + "\n"
        + LAYOUT_DESCRIPTIONS["plain"]
        + "\n\n"
        + WEIGHTS_ROW_DESCRIPTION
        + stray
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


def _weight_figure(row: pd.Series | None, column: int) -> tuple[float | None, str | None]:
    """Read one portfolio weight cell, or explain why it cannot be read.

    A blank weight is not an error: it means the fund is not currently held,
    which is exactly how a candidate is measured before it is bought.
    """
    if row is None or column >= len(row):
        return None, None
    value = row.iloc[column]
    if pd.isna(value) or (isinstance(value, str) and not str(value).strip()):
        return None, None
    converted, unreadable = to_decimal_series(pd.Series([value]))
    if unreadable or pd.isna(converted.iloc[0]):
        return None, f"'{value}' is not a number or percentage"
    return float(converted.iloc[0]), None


def normalise_weights(weights: dict[str, float]) -> tuple[dict[str, float], list[str]]:
    """Put weights on a 0-1 scale, whichever way they were typed.

    Unlike returns, weights carry their own units check: a book adds up to
    about 100%, so a total near 1 means percentage-formatted cells and a total
    near 100 means whole numbers such as 5 for 5%. Anything else is reported
    rather than guessed at. A total below 100% is taken as cash, which holds no
    risk and therefore no risk contribution.
    """
    notes: list[str] = []
    total = float(sum(weights.values()))
    if not weights or total <= 0:
        raise WorkbookFormatError(
            "The weights row has no positive weights, so there is no portfolio to analyse."
        )
    if 0.5 <= total <= 1.5:
        scaled = dict(weights)
    elif 50.0 <= total <= 150.0:
        scaled = {name: value / 100.0 for name, value in weights.items()}
        notes.append(
            f"The weights row totals {total:.1f}, so it was read as whole-number percentages "
            "(5 meaning 5%). Format those cells as % to remove any doubt."
        )
    else:
        raise WorkbookFormatError(
            f"The weights row totals {total:g}, which is neither about 1 (percentage-formatted "
            "cells) nor about 100 (whole numbers). Portfolio weights must add up to 100%, or "
            "less when part of the book is held in cash."
        )

    invested = float(sum(scaled.values()))
    if invested < 0.999:
        notes.append(
            f"The weights add up to {invested * 100:.1f}%, so the remaining "
            f"{(1.0 - invested) * 100:.1f}% is treated as cash: it holds none of the "
            "portfolio's risk."
        )
    elif invested > 1.001:
        notes.append(
            f"The weights add up to {invested * 100:.1f}%, so the book is geared. "
            "Risk figures are reported on that geared basis."
        )
    shorts = [_split_label(name)[1] for name, value in scaled.items() if value < 0]
    if shorts:
        notes.append(
            f"{_plural(len(shorts), 'negative weight was', 'negative weights were')} found and "
            f"kept as short positions: {_names(shorts)}.")
    return scaled, notes


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

    layout, data_start, weights_row_index = detect_layout(raw)
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
    weights_row = raw.iloc[weights_row_index] if weights_row_index is not None else None
    raw_weights: dict[str, float] = {}

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
            stray_weight, _ = _weight_figure(weights_row, column)
            if stray_weight is not None:
                errors.append(
                    f"The risk-free column {letter} has a portfolio weight on row "
                    f"{weights_row_index + 1}. The risk-free series is not a fund, so leave its "
                    "weight blank; cash is whatever the fund weights leave over."
                )
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

        weight, weight_problem = _weight_figure(weights_row, column)
        if weight_problem:
            errors.append(
                f"{described}: portfolio weight on row {weights_row_index + 1} - {weight_problem}."
            )
            continue

        funds.append(
            FundColumn(
                strategy=strategy,
                fund=fund_name,
                column=letter,
                returns=returns,
                expected_excess=expected_excess,
                expected_vol=expected_vol,
                weight=weight,
            )
        )
        if weight is not None:
            raw_weights[f"{strategy}::{fund_name}"] = weight

    if not funds and not errors:
        errors.append("No fund columns with usable returns were found from column B onwards.")

    if errors:
        raise WorkbookFormatError(
            f"{len(errors)} problem(s) were found in {Path(path).name} [{sheet}].\n"
            f"Detected {LAYOUT_DESCRIPTIONS[layout]}\n\n"
            + "\n\n".join(f"{number}. {message}" for number, message in enumerate(errors, start=1))
        )

    # The weights are only put on a 0-1 scale once every column has been read,
    # because their total is what tells percentages from whole numbers.
    if weights_row_index is not None:
        scaled, weight_notes = normalise_weights(raw_weights)
        notes.extend(weight_notes)
        for fund in funds:
            fund.weight = scaled.get(fund.combined_label, 0.0)
        unheld = [fund.fund for fund in funds if not fund.weight]
        if unheld:
            notes.append(
                f"{_plural(len(unheld), 'fund has', 'funds have')} no weight, so they are "
                "treated as candidates the portfolio does not hold: "
                + _names(unheld) + "."
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
        weights_row=weights_row_index,
        notes=notes,
    )


def check_modules_against_layout(layout: str, modules: Iterable[str],
                                 has_weights: bool = False) -> list[str]:
    """Report any selected module the detected layout cannot support."""
    modules = list(modules)
    problems: list[str] = []
    if "cone" in modules and layout != "cone":
        problems.append(
            "Module 1 (cone charts) needs the expected excess return on row 3 and the expected "
            "volatility on row 4, with monthly returns starting on row 5. This sheet uses the "
            "plain layout, where returns start on row 3.\n\n" + LAYOUT_DESCRIPTIONS["cone"]
        )
    if "allocation" in modules and not has_weights:
        problems.append(
            "Module 4 (portfolio risk allocation) needs to know what the portfolio holds, and "
            "this sheet has no weights row.\n\n" + WEIGHTS_ROW_DESCRIPTION
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
from matplotlib.patches import Patch  # noqa: E402
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


def _names(values: Sequence[str], limit: int = 6) -> str:
    """A readable list, truncated once it stops being readable."""
    items = [str(value) for value in values]
    if not items:
        return "none"
    if len(items) > limit:
        shown, rest = items[:limit], len(items) - limit
        return ", ".join(shown) + f" and {rest} other{'s' if rest > 1 else ''}"
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + f" and {items[-1]}"


def _trimmed(parts: Sequence[str], limit: int, separator: str = "; ") -> str:
    """A list of phrases, with an honest tail when it is longer than the reader wants.

    A sentence that counts nine funds and then names six reads as a mistake, so
    the count of what was left out is part of the sentence rather than
    something the reader has to notice.
    """
    items = list(parts)
    if len(items) <= limit:
        return separator.join(items)
    rest = len(items) - limit
    return separator.join(items[:limit]) + f"{separator}and {rest} other{'s' if rest > 1 else ''}"


def _plural(count: int, singular: str, plural: str | None = None) -> str:
    return f"{count} {singular if count == 1 else (plural or singular + 's')}"


def _fund_name(row: Any) -> str:
    """The fund on its own where that is unambiguous, with its strategy where it is not."""
    return str(getattr(row, "fund", getattr(row, "label", "")))


def _split_label(label: str) -> tuple[str, str]:
    strategy, _, fund = str(label).partition("::")
    return (strategy, fund) if fund else ("", strategy)


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
# Findings collected while the modules run
#
# The commentary document is written after every module has finished, and it
# describes what those modules actually produced rather than recomputing the
# analysis a second time. Each runner is handed an optional collector and drops
# into it the few objects the write-up needs: the cluster analyses, the risk
# decompositions, per-fund realised figures, and copies of the charts worth
# embedding.
#
# The collector is optional everywhere, so a run with the commentary switched
# off behaves exactly as it did before, and a module that fails part way
# through still leaves whatever it had already recorded.
# =============================================================================

class RunFindings:
    """What the modules recorded on the way past, for the written summary.

    Charts are copied into an assets folder at the collector's own resolution
    rather than read back from the module folders, because a module may be set
    to write no PNGs at all and the document still needs a picture.
    """

    def __init__(self, assets_dir: Path, dpi: int = 200, embed_charts: bool = True) -> None:
        self.assets_dir = assets_dir
        self.dpi = int(dpi)
        self.embed_charts = bool(embed_charts)
        self.cone: list[dict[str, Any]] = []
        self.cone_notes: list[str] = []
        self.clusters: dict[str, ClusterAnalysis] = {}
        self.cluster_exclusions: list[dict[str, Any]] = []
        self.performance: list[dict[str, Any]] = []
        self.performance_excluded: list[tuple[str, str, int, str]] = []
        self.risk: list[PortfolioRisk] = []
        self.charts: dict[str, Path] = {}

    def capture(self, key: str, figure: plt.Figure) -> None:
        """Save a copy of a figure for the document, if it wants one."""
        if not self.embed_charts or key in self.charts:
            return
        try:
            self.assets_dir.mkdir(parents=True, exist_ok=True)
            path = self.assets_dir / f"{safe_file_stem(key)}.png"
            figure.savefig(path, dpi=self.dpi, bbox_inches="tight")
        except Exception:  # a missing picture must never fail a run
            LOGGER.exception("Could not save the commentary chart %s", key)
            return
        self.charts[key] = path

    def adopt(self, key: str, path: str | Path | None) -> None:
        """Point the document at a chart a module has already written."""
        if not self.embed_charts or key in self.charts or path is None:
            return
        candidate = Path(path)
        if candidate.is_file():
            self.charts[key] = candidate


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


def run_cone_module(data: WorkbookData, options: dict[str, Any], output_dir: Path,
                    findings: RunFindings | None = None) -> list[str]:
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

        if findings is not None:
            # The same excess-return basis the cone itself is drawn on, so the
            # write-up cannot quote a number the chart disagrees with.
            excess = arithmetic_excess_returns(returns, fund_risk_free)
            realised_excess, realised_vol, realised_sharpe = annualise_stats(excess)
            findings.cone.append({
                "strategy": fund.strategy,
                "fund": fund.fund,
                "label": fund.label,
                "months": int(len(returns.dropna())),
                "start": returns.index.min(),
                "end": returns.index.max(),
                "expected_excess": fund.expected_excess,
                "expected_vol": fund.expected_vol,
                "realised_excess": realised_excess,
                "realised_vol": realised_vol,
                "realised_sharpe": realised_sharpe,
            })

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
            if findings is not None:
                findings.adopt(f"cone::{fund.label}", images["cone"])

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
    if findings is not None:
        findings.cone_notes.extend(log[:-1])

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
    minimum_funds: int = 2,
) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    """The funds a strategy can contribute to one lookback, and why the rest cannot.

    Two funds are needed to correlate a strategy with itself, which is what
    minimum_funds defaults to. The combined, all-strategies view asks for one
    instead: a fund that is the only one in its strategy still correlates
    perfectly well with every fund outside it, and leaving it out would drop a
    real holding from the picture.
    """
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
    if eligible.shape[1] < minimum_funds:
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

    if selected.shape[1] < minimum_funds or len(selected) < minimum_months:
        return selected.iloc[0:0], exclusions
    return selected, exclusions


@dataclass
class ClusterUniverse:
    """Every strategy's usable funds for one lookback, plus the combined frame.

    Module 2 and module 4 both work from this, so the clusters on a risk chart
    are always the clusters on the dendrogram beside it.
    """

    timeframe: Timeframe
    selected: dict[str, pd.DataFrame]
    exclusions: list[dict[str, Any]]
    exclusion_counts: dict[str, int]
    combined: pd.DataFrame


def build_cluster_universe(data: WorkbookData, options: dict[str, Any],
                           period: Timeframe) -> ClusterUniverse:
    """Choose the funds each strategy contributes to one lookback, and merge them."""
    frame = data.wide_frame()
    frame.index = frame.index.to_period("M")

    selected: dict[str, pd.DataFrame] = {}
    exclusions: list[dict[str, Any]] = []
    counts: dict[str, int] = {}
    renamed_frames: list[pd.DataFrame] = []

    for strategy in data.strategies:
        strategy_frame = frame.xs(strategy, axis=1, level=0, drop_level=True)
        chosen, excluded = select_cluster_returns(strategy_frame, period, options, strategy)
        exclusions.extend(excluded)
        counts[strategy] = len(excluded)
        selected[strategy] = chosen

        # The combined view takes one fund where a strategy's own chart needs
        # two, so a strategy holding a single fund still reaches the
        # all-strategies dendrogram and the portfolio's risk. The exclusions
        # are the ones recorded above, which describe the same funds.
        contributed, _ = select_cluster_returns(
            strategy_frame, period, options, strategy, minimum_funds=1)
        if contributed.empty:
            continue
        renamed = contributed.copy()
        renamed.columns = [f"{strategy}::{fund}" for fund in contributed.columns]
        renamed_frames.append(renamed)

    combined = (
        latest_contiguous_complete_block(
            pd.concat(renamed_frames, axis=1, join="outer").sort_index())
        if renamed_frames else pd.DataFrame()
    )
    return ClusterUniverse(period, selected, exclusions, counts, combined)


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
                          output_dir: Path, findings: RunFindings | None = None) -> list[str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    log: list[str] = []
    timeframes = normalise_timeframes(options["timeframes"])

    universes = {period: build_cluster_universe(data, options, period) for period in timeframes}
    exclusions: list[dict[str, Any]] = []
    run_rows: list[dict[str, Any]] = []

    with ExitStack() as stack:
        pdfs: dict[Timeframe, PdfPages | None] = {}
        for period in timeframes:
            pdfs[period] = (
                stack.enter_context(
                    PdfPages(output_dir / f"Clustering_Report_{timeframe_label(period)}.pdf"))
                if options["create_pdf"] else None
            )

        for strategy in data.strategies:
            for period in timeframes:
                universe = universes[period]
                selected = universe.selected[strategy]
                excluded_count = universe.exclusion_counts[strategy]
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
                _save_cluster_outputs(result, excluded_count, options, pdfs[period], output_dir)
                run_rows.append({
                    "strategy": strategy, "timeframe": timeframe_label(period),
                    "start": str(selected.index.min()), "end": str(selected.index.max()),
                    "months": len(selected), "funds": selected.shape[1],
                    "excluded_records": excluded_count,
                })
                log.append(
                    f"{strategy} [{timeframe_label(period)}]: {selected.shape[1]} funds "
                    f"over {len(selected)} months."
                )

        for period in timeframes:
            merged = universes[period].combined
            if merged.empty:
                continue
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
            if findings is not None:
                findings.clusters[timeframe_label(period)] = result
                if findings.embed_charts:
                    figure = plot_cluster_dendrogram(
                        result, 0, options, int(options["all_strategies_label_wrap_width"]))
                    findings.capture(f"dendrogram::{timeframe_label(period)}", figure)
                    plt.close(figure)
            run_rows.append({
                "strategy": "All Strategies Combined", "timeframe": timeframe_label(period),
                "start": str(merged.index.min()), "end": str(merged.index.max()),
                "months": len(merged), "funds": merged.shape[1], "excluded_records": 0,
            })
            log.append(
                f"All strategies [{timeframe_label(period)}]: {merged.shape[1]} funds "
                f"over {len(merged)} months."
            )

    # Kept in strategy-then-lookback order, so the audit CSV reads the way the
    # analysis is laid out rather than the order the frames happened to be built.
    for strategy in data.strategies:
        for period in timeframes:
            exclusions.extend(row for row in universes[period].exclusions
                              if row["strategy"] == strategy)

    if findings is not None:
        findings.cluster_exclusions.extend(exclusions)

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


def drawdown_records(returns_by_name: dict[str, pd.Series]) -> list[dict[str, Any]]:
    """Each fund's worst drawdown, and where it stands against its high-water mark.

    The figures are the ones the module's own drawdown charts are drawn from,
    so the write-up and the chart cannot disagree. A fund still below its peak
    is reported with the number of months it has been there, because a shallow
    drawdown that has lasted three years is a different problem from a deep one
    that is already recovering.
    """
    # A drawdown this shallow is rounding rather than a position below water.
    at_peak_tolerance = 0.0001

    records: list[dict[str, Any]] = []
    for name, returns in returns_by_name.items():
        series = pd.to_numeric(returns, errors="coerce").dropna()
        if series.empty:
            continue
        # The chart's own drawdowns are in percent; everything here is decimal,
        # so that one threshold cannot mean two different things.
        drawdowns = calculate_actual_drawdowns(calculate_price_index(series)).dropna() / 100.0
        if drawdowns.empty:
            continue
        current = float(drawdowns.iloc[-1])

        # Months since the fund was last at, or above, its high-water mark. A
        # fund that has never regained its opening value has been under water
        # for its whole life, which is not the same as nought months.
        at_peak = drawdowns[drawdowns >= -at_peak_tolerance]
        if current >= -at_peak_tolerance:
            months_under = 0
        elif at_peak.empty:
            months_under = int(len(drawdowns))
        else:
            months_under = int((drawdowns.index > at_peak.index[-1]).sum())

        annual_return, annual_vol, _ = annualise_stats(series)
        strategy, _, fund = str(name).partition(" - ")
        records.append({
            "label": name,
            "strategy": strategy if fund else "",
            "fund": fund or str(name),
            "months": int(len(series)),
            "start": series.index.min(),
            "end": series.index.max(),
            "max_drawdown": float(drawdowns.min()),
            "trough": drawdowns.idxmin(),
            "current_drawdown": current,
            "months_underwater": months_under,
            "at_worst_now": bool(np.isclose(current, float(drawdowns.min()))),
            "annual_return": annual_return,
            "annual_vol": annual_vol,
        })
    return records


def run_performance_module(data: WorkbookData, options: dict[str, Any],
                           output_dir: Path, findings: RunFindings | None = None) -> list[str]:
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

    if findings is not None:
        findings.performance.extend(drawdown_records(all_returns))
        findings.performance_excluded.extend(unique_excluded)
        # Interactive line charts leave no picture behind and the PNGs are off
        # by default, so the document takes its own copy of the two combined
        # charts while the figures are still open.
        wanted = {
            f"NAV Line Plot - {combined_title}": "performance_nav",
            f"Actual Drawdowns Heatmap - {combined_title}": "performance_drawdown_heatmap",
        }
        for figure, _, description in figures:
            if description in wanted:
                findings.capture(wanted[description], figure)

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
# Module 4 - portfolio risk allocation
#
# Weight is not risk. Three managed-futures funds at 5% each are a smaller
# block of risk than their 15% of the book suggests, because they diversify
# one another, while three multi-strategy funds at 5% each are a larger one.
# This module answers the question the dendrogram raises but cannot settle:
# how much of the portfolio's risk sits in each of those clusters, and which
# fund is the next one to add to or trim.
#
# The decomposition is the standard one. With weights w and covariance S the
# portfolio's volatility is sqrt(w'Sw), fund i's marginal contribution is
# (Sw)i / vol and its component contribution is wi times that. Those
# components add up to the portfolio's volatility exactly, so they can be
# grouped by cluster with nothing left over.
#
# The clusters come from module 2's own linkage, so a cluster on these charts
# is the same cluster on the dendrogram beside it.
# =============================================================================

RISK_PALETTE = "tab10"
UNHELD_COLOUR = "0.55"


def unsmooth_frame(returns: pd.DataFrame, max_rho: float = 0.6,
                   min_rho: float = 0.05) -> tuple[pd.DataFrame, dict[str, float]]:
    """Undo return smoothing, one fund at a time (Geltner).

    Monthly hedge fund marks are stale, especially in credit and event, which
    shows up as positive autocorrelation and makes reported volatility and
    correlation too low. Reversing it with r*(t) = (r(t) - p r(t-1)) / (1 - p)
    puts those funds back on a comparable footing with the daily-marked ones.

    Only positive autocorrelation is reversed, and it is capped, because the
    estimate is noisy on a short window and the correction divides by 1 - p.
    The first month goes, since it has no predecessor; every fund loses the
    same one, so the window stays rectangular.
    """
    applied: dict[str, float] = {}
    adjusted = returns.copy()
    for name in returns.columns:
        series = returns[name].astype(float)
        rho = series.autocorr(lag=1)
        if not np.isfinite(rho) or rho < min_rho:
            continue
        rho = float(min(rho, max_rho))
        adjusted[name] = (series - rho * series.shift(1)) / (1.0 - rho)
        applied[str(name)] = rho
    return adjusted.iloc[1:], applied


def ledoit_wolf_covariance(returns: pd.DataFrame) -> tuple[np.ndarray, float]:
    """Sample covariance shrunk towards a scaled identity (Ledoit-Wolf 2004).

    Sixty months of twenty-odd funds is enough to estimate a covariance but
    not enough to trust every entry of it, and risk contributions are more
    sensitive to that noise than volatility is. Shrinking towards equal,
    uncorrelated variances pulls the least reliable entries in by an amount
    the data itself chooses.
    """
    values = returns.to_numpy(dtype=float)
    months, count = values.shape
    centred = values - values.mean(axis=0, keepdims=True)
    sample = (centred.T @ centred) / months

    mean_variance = float(np.trace(sample) / count)
    target = mean_variance * np.eye(count)
    dispersion = float(np.sum((sample - target) ** 2) / count)
    if dispersion <= 0:
        return sample, 0.0

    # sum_t ||x_t x_t' - S||^2 reduces to sum_t ||x_t||^4 - T ||S||^2, which
    # avoids building one outer product per month.
    squared_norms = np.einsum("ij,ij->i", centred, centred)
    noise = float((np.sum(squared_norms ** 2) - months * np.sum(sample ** 2))
                  / (count * months * months))
    shrinkage = float(np.clip(min(noise, dispersion) / dispersion, 0.0, 1.0))
    return shrinkage * target + (1.0 - shrinkage) * sample, shrinkage


def estimate_covariance(returns: pd.DataFrame, method: str) -> tuple[pd.DataFrame, float]:
    """The monthly covariance of the window, by the chosen estimator."""
    if method == "ledoit_wolf":
        matrix, shrinkage = ledoit_wolf_covariance(returns)
    else:
        matrix, shrinkage = returns.cov().to_numpy(dtype=float), 0.0
    return pd.DataFrame(matrix, index=returns.columns, columns=returns.columns), shrinkage


def annualised_return(series: pd.Series) -> float:
    """The compound annual rate implied by a run of monthly returns."""
    clean = series.dropna().astype(float)
    if clean.empty:
        return float("nan")
    growth = float(np.expm1(np.log1p(clean).sum()))
    return float((1.0 + growth) ** (12.0 / len(clean)) - 1.0)


def describe_cluster(members: Sequence[str]) -> str:
    """Name a cluster after the strategies inside it, not just its number."""
    strategies: list[str] = []
    for name in members:
        strategy = str(name).split("::", 1)[0]
        if strategy not in strategies:
            strategies.append(strategy)
    counts = {name: sum(1 for member in members if str(member).startswith(f"{name}::"))
              for name in strategies}
    ranked = sorted(strategies, key=lambda name: (-counts[name], name))
    if len(ranked) == 1:
        return ranked[0]
    if len(ranked) == 2:
        return " / ".join(ranked)
    return f"{ranked[0]} +{len(ranked) - 1} more"


@dataclass
class PortfolioRisk:
    """One lookback's risk decomposition, ready to chart or write out."""

    timeframe_label: str
    returns: pd.DataFrame
    weights: pd.Series
    covariance: pd.DataFrame
    analysis: ClusterAnalysis
    funds: pd.DataFrame
    clusters: pd.DataFrame
    portfolio_vol: float
    portfolio_return: float
    diversification_ratio: float
    effective_bets: float
    weight_concentration: float
    risk_concentration: float
    shrinkage: float
    return_basis: str
    unsmoothed: dict[str, float]
    uncovered: pd.Series
    cash_weight: float
    notes: list[str]

    @property
    def cluster_colours(self) -> dict[int, Any]:
        palette = plt.get_cmap(RISK_PALETTE)
        return {int(cluster): palette(position % palette.N)
                for position, cluster in enumerate(sorted(self.clusters["cluster"]))}


def expected_returns_for(universe: Sequence[str], data: WorkbookData, returns: pd.DataFrame,
                         basis: str) -> tuple[pd.Series, str, list[str]]:
    """The return each fund is expected to earn, on one consistent basis.

    The workbook's own expectations are used when the cone layout supplies one
    for every fund in the window, because those are the numbers the portfolio
    was built on. Otherwise every fund falls back to what it actually
    delivered, rather than mixing a stated view for some funds with a realised
    number for others on the same axis.
    """
    notes: list[str] = []
    stated = {fund.combined_label: fund.expected_excess for fund in data.funds}
    complete = all(stated.get(name) is not None for name in universe)

    if basis in {"expected", "auto"} and complete:
        return (pd.Series({name: float(stated[name]) for name in universe}),
                "expected excess return from the workbook", notes)
    if basis == "expected" and not complete:
        missing = [name for name in universe if stated.get(name) is None]
        notes.append(
            f"{len(missing)} fund(s) in the window have no expected excess return on row 3, so "
            "every fund is shown on realised returns instead: "
            + ", ".join(missing[:4])
            + ("..." if len(missing) > 4 else "")
            + "."
        )

    realised = pd.Series({name: annualised_return(returns[name]) for name in universe})
    label = "realised annualised return over the window"
    if data.risk_free is not None:
        window = data.risk_free.copy()
        window.index = window.index.to_period("M")
        cash = annualised_return(window.reindex(returns.index))
        if np.isfinite(cash):
            realised = realised - cash
            label = (f"realised annualised excess return over the window "
                     f"({data.risk_free_label}: {format_percent(cash)} a year)")
    return realised, label, notes


def build_portfolio_risk(data: WorkbookData, returns: pd.DataFrame, analysis: ClusterAnalysis,
                         options: dict[str, Any], label: str) -> PortfolioRisk:
    """Decompose the portfolio's volatility across its funds and clusters."""
    notes: list[str] = []
    universe = [str(name) for name in returns.columns]

    all_weights = data.weights()
    held = all_weights[all_weights != 0.0]
    uncovered = held.drop(labels=[name for name in universe if name in held.index])
    if not uncovered.empty:
        notes.append(
            f"{_plural(len(uncovered), 'held fund is', 'held funds are')} not in this window, "
            "so every share of risk and capital here is of the other "
            f"{(1.0 - uncovered.sum()) * 100:.1f}% of the book rather than all of it. Left out: "
            + _names([f"{_split_label(str(name))[1]} ({value * 100:.1f}%)"
                      for name, value in uncovered.items()])
            + ". Module 2's excluded_funds_audit.csv says why."
        )

    working = returns.astype(float)
    unsmoothed: dict[str, float] = {}
    if options["unsmooth_returns"]:
        working, unsmoothed = unsmooth_frame(working)
        if unsmoothed:
            worst = sorted(unsmoothed.items(), key=lambda item: -item[1])[:3]
            notes.append(
                f"Return smoothing was reversed for {len(unsmoothed)} of {len(universe)} funds, "
                "raising their volatility and correlations. Most autocorrelated: "
                + ", ".join(f"{name} (rho {rho:.2f})" for name, rho in worst)
                + f". One month is lost to the adjustment, leaving {len(working)}."
            )
        else:
            notes.append("No fund showed enough positive autocorrelation to unsmooth.")

    covariance, shrinkage = estimate_covariance(working, options["covariance_method"])
    annual = covariance * 12.0

    weights = pd.Series({name: float(all_weights.get(name, 0.0)) for name in universe})
    invested = float(weights.sum())
    if invested <= 0:
        raise ValueError(
            "None of the funds in this window is held, so the portfolio has no risk to allocate."
        )
    cash_weight = float(1.0 - all_weights.sum())

    vector = weights.to_numpy(dtype=float)
    matrix = annual.to_numpy(dtype=float)
    variance = float(vector @ matrix @ vector)
    if variance <= 0:
        raise ValueError("The portfolio's estimated variance is not positive.")
    portfolio_vol = float(np.sqrt(variance))

    marginal = (matrix @ vector) / portfolio_vol          # d(vol) / d(weight)
    contribution = vector * marginal                      # adds up to the portfolio vol
    standalone = np.sqrt(np.diag(matrix))
    risk_share = contribution / portfolio_vol
    weight_share = vector / invested

    expected, basis, basis_notes = expected_returns_for(universe, data, working,
                                                        options["return_basis"])
    notes.extend(basis_notes)

    assignment = analysis.assignments.set_index("fund")["cluster"]
    table = pd.DataFrame({
        "strategy": [name.split("::", 1)[0] for name in universe],
        "fund": [name.split("::", 1)[-1] for name in universe],
        "label": universe,
        "cluster": [int(assignment[name]) for name in universe],
        "weight": vector,
        "weight_share": weight_share,
        "expected_return": expected.reindex(universe).to_numpy(dtype=float),
        "standalone_vol": standalone,
        "marginal_risk": marginal,
        "risk_contribution": contribution,
        "risk_share": risk_share,
        "correlation_to_portfolio": marginal / standalone,
    })
    # A multiplier above 1 means the fund carries more of the risk than of the
    # money. It is undefined for a fund that is not held, which is the point of
    # the marginal figure beside it.
    table["risk_multiplier"] = np.where(
        table["weight_share"] > 0, table["risk_share"] / table["weight_share"], np.nan)
    table["risk_share_gap"] = table["risk_share"] - table["weight_share"]

    grouped = table.groupby("cluster", sort=True)
    clusters = pd.DataFrame({
        "cluster": [int(key) for key in grouped.groups],
        "members": grouped.size().to_numpy(),
        "held": grouped["weight"].apply(lambda column: int((column != 0).sum())).to_numpy(),
        "weight": grouped["weight"].sum().to_numpy(),
        "weight_share": grouped["weight_share"].sum().to_numpy(),
        "risk_share": grouped["risk_share"].sum().to_numpy(),
        "risk_contribution": grouped["risk_contribution"].sum().to_numpy(),
    })
    clusters["label"] = [describe_cluster(list(grouped.get_group(key)["label"]))
                         for key in clusters["cluster"]]
    clusters["risk_multiplier"] = np.where(
        clusters["weight_share"] > 0, clusters["risk_share"] / clusters["weight_share"], np.nan)
    clusters = clusters.sort_values("risk_share", ascending=False).reset_index(drop=True)

    weighted_vol = float(np.abs(vector) @ standalone)
    diversification = weighted_vol / portfolio_vol if portfolio_vol > 0 else float("nan")
    # Squared shares are taken on absolute values so that a short position adds
    # to concentration rather than quietly cancelling part of it.
    absolute_risk = np.abs(risk_share)
    absolute_weight = np.abs(weight_share)
    risk_concentration = float(np.sum((absolute_risk / absolute_risk.sum()) ** 2))
    weight_concentration = float(np.sum((absolute_weight / absolute_weight.sum()) ** 2))

    return PortfolioRisk(
        timeframe_label=label,
        returns=working,
        weights=weights,
        covariance=annual,
        analysis=analysis,
        funds=table,
        clusters=clusters,
        portfolio_vol=portfolio_vol,
        portfolio_return=float(vector @ table["expected_return"].to_numpy(dtype=float)),
        diversification_ratio=diversification,
        effective_bets=1.0 / risk_concentration if risk_concentration > 0 else float("nan"),
        weight_concentration=weight_concentration,
        risk_concentration=risk_concentration,
        shrinkage=shrinkage,
        return_basis=basis,
        unsmoothed=unsmoothed,
        uncovered=uncovered,
        cash_weight=cash_weight,
        notes=notes,
    )


def _allocation_footer(figure: plt.Figure, risk: PortfolioRisk, options: dict[str, Any],
                       y: float = 0.006) -> None:
    estimator = ("Ledoit-Wolf shrinkage" if options["covariance_method"] == "ledoit_wolf"
                 else "sample covariance")
    if options["covariance_method"] == "ledoit_wolf":
        estimator += f" ({risk.shrinkage * 100:.0f}% towards the target)"
    note = (
        f"{risk.returns.index.min().strftime('%b %Y')}-"
        f"{risk.returns.index.max().strftime('%b %Y')} ({len(risk.returns)} months); "
        f"{risk.returns.shape[1]} funds; {estimator}; "
        f"returns {'unsmoothed' if risk.unsmoothed else 'as reported'}; "
        f"portfolio volatility {format_percent(risk.portfolio_vol)} a year."
    )
    figure.text(0.99, y, note, ha="right", va="bottom", fontsize=7, color="0.25")


# Legal form carries no information on a chart, and on a fund called something
# like "... Global Strategies Fund, Ltd." it is most of the room a label has.
LEGAL_SUFFIX = re.compile(
    r"[\s,]+(?:Ltd|Limited|L\.?P|LLC|LLP|plc|Inc|SPC|N\.?V|S\.?A|Co)\.?$", re.IGNORECASE)
TRAILING_FUND = re.compile(r"[\s,]+Funds?$", re.IGNORECASE)


def short_fund_name(name: str) -> str:
    """A fund's name with its legal form trimmed off, for use on a chart."""
    text = str(name).strip()
    for _ in range(3):
        trimmed = TRAILING_FUND.sub("", LEGAL_SUFFIX.sub("", text)).strip(" ,-")
        if trimmed == text or not trimmed:
            break
        text = trimmed
    return text or str(name).strip()


def _data_width(axis: plt.Axes, text: Any, renderer: Any) -> float:
    """How wide a drawn label is, measured in the axis's own units."""
    box = text.get_window_extent(renderer=renderer)
    inverse = axis.transData.inverted()
    left, right = inverse.transform([(box.x0, box.y0), (box.x1, box.y0)])
    return abs(float(right[0]) - float(left[0]))


def _fit_label(axis: plt.Axes, renderer: Any, x: float, y: float, name: str,
               available: float, minimum: int = 7, **kwargs: Any) -> Any | None:
    """Write a label inside a bar segment, trimmed to what will actually fit.

    Guessing the width from a character count is what puts one fund's name on
    top of the next one's, because it depends on the font, the figure size and
    the final layout. Measuring the drawn text instead is exact, so this runs
    after the layout is settled.
    """
    text = axis.text(x, y, name, **kwargs)
    width = _data_width(axis, text, renderer)
    if width <= available:
        return text
    if width <= 0:
        text.remove()
        return None
    fits = int(len(name) * available / width) - 1
    if fits < minimum:
        text.remove()
        return None
    text.set_text(name[:fits].rstrip(" ,-"))
    if _data_width(axis, text, renderer) > available:
        text.remove()
        return None
    return text


def _is_dark(colour: Any) -> bool:
    """Whether white text reads better than black on this background."""
    red, green, blue = mcolors.to_rgb(colour)
    return (0.299 * red + 0.587 * green + 0.114 * blue) < 0.55


def _shade(colour: Any, position: int, total: int) -> tuple[float, float, float]:
    """One cluster's colour, lightened a step per fund, so members stay apart."""
    base = np.array(mcolors.to_rgb(colour))
    if total <= 1:
        return tuple(base)
    return tuple(base + (1.0 - base) * (0.62 * position / (total - 1)))


def plot_cluster_risk_budget(risk: PortfolioRisk, options: dict[str, Any]) -> plt.Figure:
    """Share of money against share of risk, cluster by cluster.

    The left panel is the answer to "where is the risk": a cluster whose risk
    bar overshoots its weight bar is carrying more than it is being paid for in
    capital. The right panel says which fund inside the cluster is doing it.
    """
    clusters = risk.clusters
    colours = risk.cluster_colours
    height = max(4.5, 1.15 * len(clusters) + 2.6)
    figure, (budget, breakdown) = plt.subplots(
        1, 2, figsize=(15.5, height), gridspec_kw={"width_ratios": [1.15, 1.0]})

    positions = np.arange(len(clusters))
    bar_height = 0.36
    budget.barh(positions - bar_height / 2, clusters["weight_share"], height=bar_height,
                color="0.72", edgecolor="0.35", linewidth=0.6, label="Share of capital")
    budget.barh(positions + bar_height / 2, clusters["risk_share"], height=bar_height,
                color=[colours[int(key)] for key in clusters["cluster"]],
                edgecolor="0.25", linewidth=0.6)

    limit = float(max(clusters["weight_share"].max(), clusters["risk_share"].max()))
    for position, row in enumerate(clusters.itertuples()):
        budget.text(row.weight_share + limit * 0.015, position - bar_height / 2,
                    format_percent(row.weight_share, 1), va="center", fontsize=8, color="0.35")
        budget.text(row.risk_share + limit * 0.015, position + bar_height / 2,
                    format_percent(row.risk_share, 1), va="center", fontsize=8.5,
                    fontweight="bold", color="0.15")

    labels = [
        textwrap.fill(f"C{row.cluster}  {row.label}", 30)
        + f"\n{row.members} funds, {row.held} held"
        for row in clusters.itertuples()
    ]
    budget.set_yticks(positions)
    budget.set_yticklabels(labels, fontsize=8.5)
    budget.invert_yaxis()
    budget.set_xlim(0, limit * 1.22)
    budget.xaxis.set_major_formatter(PercentFormatter(xmax=1.0, decimals=0))
    budget.set_xlabel("Share of the invested book")
    budget.grid(axis="x", linestyle="--", alpha=0.35)
    budget.set_axisbelow(True)
    # The risk bars take each cluster's own colour, so the legend cannot show a
    # single swatch for them without implying one of the clusters is "the" one.
    budget.legend(
        handles=[
            Patch(facecolor="0.72", edgecolor="0.35", label="Share of capital"),
            Patch(facecolor="0.32", edgecolor="0.25", label="Share of risk (coloured by cluster)"),
        ],
        fontsize=8.5, loc="lower right", framealpha=0.92)
    budget.set_title("Capital against risk, by cluster", fontsize=11, pad=10)

    # The multiplier column earns its place: it is the one number that says
    # whether a cluster is punching above its weight, and by how much.
    for position, row in enumerate(clusters.itertuples()):
        multiplier = row.risk_multiplier
        if not np.isfinite(multiplier):
            text, colour = "no weight", "0.5"
        else:
            text = f"{multiplier:.2f}x"
            colour = ("#b2182b" if multiplier >= float(options["risk_multiplier_flag"])
                      else "#2166ac" if multiplier <= 1.0 / float(options["risk_multiplier_flag"])
                      else "0.3")
        budget.annotate(text, xy=(1.0, position), xycoords=("axes fraction", "data"),
                        xytext=(6, 0), textcoords="offset points", va="center", ha="left",
                        fontsize=8.5, color=colour, fontweight="bold", annotation_clip=False)
    budget.annotate("risk\nper unit\nof capital", xy=(1.0, -0.7),
                    xycoords=("axes fraction", "data"), xytext=(6, 0),
                    textcoords="offset points", va="center", ha="left", fontsize=7.5,
                    color="0.4", annotation_clip=False)

    # A segment is only labelled when the name will fit inside it, measured
    # against this panel's own scale rather than the one beside it.
    segments: list[tuple[float, float, str, float, Any]] = []
    for position, cluster in enumerate(clusters["cluster"]):
        members = (risk.funds[risk.funds["cluster"] == cluster]
                   .sort_values("risk_share", ascending=False))
        left = 0.0
        for order, row in enumerate(members.itertuples()):
            colour = _shade(colours[int(cluster)], order, len(members))
            breakdown.barh(position, row.risk_share, left=left, height=0.62,
                           color=colour, edgecolor="white", linewidth=0.7)
            segments.append((left + row.risk_share / 2, float(position),
                             short_fund_name(row.fund), float(row.risk_share), colour))
            left += row.risk_share

    breakdown.set_yticks(positions)
    breakdown.set_yticklabels([f"C{key}" for key in clusters["cluster"]], fontsize=9)
    breakdown.invert_yaxis()
    breakdown.xaxis.set_major_formatter(PercentFormatter(xmax=1.0, decimals=0))
    breakdown.set_xlabel("Share of portfolio risk")
    breakdown.grid(axis="x", linestyle="--", alpha=0.35)
    breakdown.set_axisbelow(True)
    breakdown.set_title("Which funds make up each cluster's risk", fontsize=11, pad=10)

    figure.suptitle(
        f"Portfolio risk by cluster ({risk.timeframe_label})", fontsize=14, y=0.985)
    _allocation_footer(figure, risk, options)
    figure.tight_layout(rect=(0.01, 0.035, 0.955, 0.945))
    figure.subplots_adjust(wspace=0.32)

    try:
        figure.canvas.draw()
        renderer = figure.canvas.get_renderer()
    except Exception:  # a backend with no measurable renderer: leave them off
        return figure
    for x, y, name, width, colour in segments:
        _fit_label(breakdown, renderer, x, y, name, width * 0.92,
                   va="center", ha="center", fontsize=6.4,
                   color="white" if _is_dark(colour) else "0.12")
    return figure


def cluster_link_colours(linkage_matrix: np.ndarray, clusters: Sequence[int],
                         colours: dict[int, Any]) -> Callable[[int], Any]:
    """Colour each branch by the cluster it belongs to, grey where it joins two.

    scipy's own colouring works from a distance threshold, which need not agree
    with the cluster assignment used everywhere else on the page. Driving it
    from the assignment instead keeps one cluster one colour across the report.
    """
    leaves = len(clusters)
    node_cluster: dict[int, int | None] = {index: int(clusters[index]) for index in range(leaves)}
    for step, row in enumerate(linkage_matrix):
        left, right = int(row[0]), int(row[1])
        first, second = node_cluster.get(left), node_cluster.get(right)
        node_cluster[leaves + step] = first if first is not None and first == second else None

    def colour_of(node: int) -> Any:
        cluster = node_cluster.get(int(node))
        return mcolors.to_hex(colours[cluster]) if cluster is not None else UNHELD_COLOUR

    return colour_of


def plot_weighted_dendrogram(risk: PortfolioRisk, options: dict[str, Any]) -> plt.Figure:
    """The clustering picture, with what the portfolio actually owns beside it.

    The dendrogram on its own shows which funds behave alike. Putting the
    weight and the risk contribution in the margin shows how much has been
    staked on each of those behaviours, which is what the tree cannot say.
    """
    order = risk.analysis.leaf_order
    names = [str(name) for name in risk.analysis.returns.columns]
    table = risk.funds.set_index("label")
    colours = risk.cluster_colours
    count = len(names)

    figure = plt.figure(figsize=(max(13.0, 9.0 + count * 0.2), max(7.5, 4.2 + count * 0.36)))
    grid = figure.add_gridspec(1, 3, width_ratios=[3.1, 1.0, 1.0])
    tree = figure.add_subplot(grid[0, 0])
    weight_axis = figure.add_subplot(grid[0, 1])
    risk_axis = figure.add_subplot(grid[0, 2])

    assignment = risk.analysis.assignments.set_index("fund")["cluster"]
    dendrogram(
        risk.analysis.linkage_matrix, orientation="right",
        labels=[textwrap.fill(name, int(options["label_wrap_width"])) for name in names],
        leaf_font_size=max(5, 9 - count // 12), distance_sort=False,
        link_color_func=cluster_link_colours(
            risk.analysis.linkage_matrix, [int(assignment[name]) for name in names], colours),
        ax=tree,
    )
    tree.set_xlabel(
        "Ward distance on standardised return profiles"
        if options["linkage_method"] == "ward"
        else f"{options['linkage_method'].title()} linkage distance (1 - correlation)"
    )
    tree.set_ylabel("Fund / Strategy::Fund")
    tree.grid(axis="x", linestyle="--", alpha=0.35)
    tree.set_title("Hierarchical clustering", fontsize=11, pad=10)

    # scipy lays leaves out at 5, 15, 25 ... in leaf order, so the two margin
    # panels can share those positions exactly rather than approximating them.
    positions = np.array([5.0 + 10.0 * step for step in range(count)])
    ordered = [names[index] for index in order]
    limits = tree.get_ylim()

    for axis, column, title, formatter in (
        (weight_axis, "weight", "Weight", lambda value: format_percent(value, 1)),
        (risk_axis, "risk_share", "Share of risk", lambda value: format_percent(value, 1)),
    ):
        values = np.array([float(table.loc[name, column]) for name in ordered])
        bar_colours = [
            colours[int(table.loc[name, "cluster"])] if values[step] != 0 else UNHELD_COLOUR
            for step, name in enumerate(ordered)
        ]
        axis.barh(positions, values, height=6.4, color=bar_colours,
                  edgecolor="0.3", linewidth=0.5)
        span = float(np.max(np.abs(values))) or 1.0
        for position, value in zip(positions, values, strict=True):
            axis.text(value + span * 0.03 if value >= 0 else value - span * 0.03, position,
                      formatter(value) if value else "-", va="center",
                      ha="left" if value >= 0 else "right", fontsize=6.8,
                      color="0.2" if value else "0.55")
        axis.set_ylim(limits)
        axis.set_yticks([])
        axis.set_xlim(min(0.0, float(values.min()) * 1.35), span * 1.34)
        axis.xaxis.set_major_formatter(PercentFormatter(xmax=1.0))
        axis.grid(axis="x", linestyle="--", alpha=0.3)
        axis.set_axisbelow(True)
        axis.tick_params(axis="x", labelsize=7.5)
        axis.set_title(title, fontsize=10, pad=10)

    weight_axis.set_xlabel("of the book")
    risk_axis.set_xlabel("of portfolio volatility")
    handles = [
        Line2D([0], [0], marker="s", linestyle="none", markersize=8,
               markerfacecolor=colours[int(row.cluster)], markeredgecolor="0.3",
               label=f"C{int(row.cluster)}  {row.label}")
        for row in risk.clusters.sort_values("cluster").itertuples()
    ]
    handles.append(Line2D([0], [0], marker="s", linestyle="none", markersize=8,
                          markerfacecolor=UNHELD_COLOUR, markeredgecolor="0.3",
                          label="not held"))
    figure.legend(handles=handles, loc="lower left", bbox_to_anchor=(0.02, 0.0),
                  ncol=min(4, len(handles)), fontsize=8, frameon=False)

    figure.suptitle(
        f"Where the portfolio's risk sits in the cluster tree ({risk.timeframe_label})",
        fontsize=14, y=0.985)
    _allocation_footer(figure, risk, options, y=0.052)
    figure.tight_layout(rect=(0.01, 0.095, 0.99, 0.945))
    figure.subplots_adjust(wspace=0.06)
    return figure


def _spread_labels(values: Sequence[float], low: float, high: float,
                   gap: float) -> list[float]:
    """Push point labels apart until none overlaps, keeping their order.

    One pass upwards opens the gaps; a second pass downwards pulls the column
    back inside the axis when the first pass has run past the top. With a gap
    of at most the span divided by the number of labels, the two passes always
    leave a layout that fits.
    """
    order = sorted(range(len(values)), key=lambda index: values[index])
    placed = [float(value) for value in values]

    previous = low - gap
    for index in order:
        placed[index] = max(placed[index], previous + gap)
        previous = placed[index]

    previous = high + gap
    for index in reversed(order):
        placed[index] = min(placed[index], previous - gap)
        previous = placed[index]
    return placed


def plot_add_trim_scatter(risk: PortfolioRisk, options: dict[str, Any]) -> plt.Figure:
    """What the next pound into each fund costs in risk, against what it earns.

    The diagonal is the portfolio's own return per unit of risk. A fund above
    it improves the portfolio's return-to-risk at the margin, so the next pound
    is better spent there; a fund below it is being paid less than the risk it
    adds. The portfolio itself sits on the line by construction, because the
    marginal contributions add up to its volatility.
    """
    table = risk.funds.sort_values("marginal_risk").reset_index(drop=True)
    colours = risk.cluster_colours
    count = len(table)
    figure, axis = plt.subplots(figsize=(14.5, max(9.0, 5.4 + count * 0.24)))

    slope = risk.portfolio_return / risk.portfolio_vol if risk.portfolio_vol > 0 else np.nan
    x_values = table["marginal_risk"].to_numpy(dtype=float)
    y_values = table["expected_return"].to_numpy(dtype=float)

    x_low = min(0.0, float(np.nanmin(x_values)) * 1.15)
    x_high = max(float(np.nanmax(x_values)), risk.portfolio_vol) * 1.12
    y_low = min(float(np.nanmin(y_values)), risk.portfolio_return)
    y_high = max(float(np.nanmax(y_values)), risk.portfolio_return)
    y_pad = (y_high - y_low) * 0.12 or 0.01
    y_low, y_high = y_low - y_pad, y_high + y_pad
    axis.set_xlim(x_low, x_high)
    axis.set_ylim(y_low, y_high)

    if np.isfinite(slope):
        line = np.array([x_low, x_high])
        axis.fill_between(line, slope * line, y_high, color="#2166ac", alpha=0.055, zorder=0)
        axis.fill_between(line, y_low, slope * line, color="#b2182b", alpha=0.055, zorder=0)
        axis.plot(line, slope * line, color="0.3", linestyle="--", linewidth=1.3, zorder=2,
                  label=f"Portfolio return per unit of risk ({slope:.2f})")

    held = (table["weight"] != 0).to_numpy()
    sizes = (60.0 + 2600.0 * table["weight"].abs()).to_numpy()
    axis.scatter(x_values[held], y_values[held], s=sizes[held],
                 c=[colours[int(key)] for key in table.loc[held, "cluster"]],
                 edgecolor="0.2", linewidth=0.8, alpha=0.85, zorder=3)
    if (~held).any():
        axis.scatter(x_values[~held], y_values[~held], s=130, facecolor="none",
                     edgecolor=[colours[int(key)] for key in table.loc[~held, "cluster"]],
                     linewidth=1.9, linestyle="--", zorder=3)
    axis.scatter([risk.portfolio_vol], [risk.portfolio_return], marker="*", s=520,
                 color="#111111", edgecolor="white", linewidth=1.0, zorder=5,
                 label="The portfolio as it stands")

    # Labels go in the margin at the right, spread far enough apart to stay
    # legible however many funds share a return, with a leader back to the dot.
    positions = _spread_labels(list(y_values), y_low, y_high,
                               (y_high - y_low) / (count + 1.4))
    for row, position in zip(table.itertuples(), positions, strict=True):
        name = short_fund_name(row.fund)[:38]
        suffix = "" if row.weight else "  (not held)"
        axis.annotate(
            f"{name}{suffix}", xy=(row.marginal_risk, row.expected_return),
            xytext=(x_high * 1.015, position), textcoords="data",
            fontsize=7.4, va="center", ha="left",
            color="0.2" if row.weight else "0.45",
            arrowprops={"arrowstyle": "-", "color": "0.78", "linewidth": 0.6,
                        "shrinkA": 1, "shrinkB": 3,
                        "connectionstyle": "arc3,rad=0.0"},
            annotation_clip=False,
        )

    if y_low < 0 < y_high:
        axis.axhline(0.0, color="0.75", linewidth=0.8, zorder=1)
    axis.xaxis.set_major_formatter(PercentFormatter(xmax=1.0, decimals=1))
    axis.yaxis.set_major_formatter(PercentFormatter(xmax=1.0, decimals=1))
    axis.set_xlabel("Marginal contribution to portfolio volatility - what the next pound adds")
    axis.set_ylabel(f"Expected return\n({risk.return_basis})")
    axis.grid(linestyle="--", alpha=0.3)
    axis.set_axisbelow(True)
    axis.text(0.0, 1.012,
              "Above the line, the next pound improves the portfolio's return per unit of risk. "
              "Below it, trimming does. Bubble size is the current weight.",
              transform=axis.transAxes, ha="left", va="bottom", fontsize=8.8, color="0.35")

    handles = [
        Line2D([0], [0], marker="o", linestyle="none", markersize=9,
               markerfacecolor=colours[int(row.cluster)], markeredgecolor="0.2",
               label=f"C{int(row.cluster)}  {row.label}")
        for row in risk.clusters.sort_values("cluster").itertuples()
    ]
    handles.append(Line2D([0], [0], marker="o", linestyle="none", markersize=9,
                          markerfacecolor="none", markeredgecolor="0.4", markeredgewidth=1.7,
                          label="Not held - a candidate"))
    handles.extend(axis.get_legend_handles_labels()[0])
    figure.legend(handles=handles, loc="lower left", bbox_to_anchor=(0.015, 0.0),
                  ncol=min(4, len(handles)), fontsize=8.2, frameon=False)

    figure.suptitle(f"Where the next pound belongs ({risk.timeframe_label})",
                    fontsize=14, y=0.985)
    _allocation_footer(figure, risk, options, y=0.062)
    figure.tight_layout(rect=(0.01, 0.115, 0.795, 0.945))
    return figure


def concentration_lines(risk: PortfolioRisk) -> list[tuple[str, str]]:
    """The headline numbers, as label-and-value pairs."""
    funds = len(risk.funds)
    held = int((risk.funds["weight"] != 0).sum())
    top = risk.funds.nlargest(min(5, funds), "risk_share")["risk_share"].sum()
    return [
        ("Portfolio volatility, a year", format_percent(risk.portfolio_vol)),
        ("Funds held, of those in the window", f"{held} of {funds}"),
        ("Effective number of bets", f"{risk.effective_bets:.1f}"),
        ("Diversification ratio", f"{risk.diversification_ratio:.2f}"),
        ("Risk in the top 5 funds", format_percent(top)),
        ("Concentration of capital (Herfindahl)", f"{risk.weight_concentration:.3f}"),
        ("Concentration of risk (Herfindahl)", f"{risk.risk_concentration:.3f}"),
        ("Cash", format_percent(risk.cash_weight) if risk.cash_weight > 0.0005 else "none"),
    ]


def plot_concentration_summary(risk: PortfolioRisk, options: dict[str, Any]) -> plt.Figure:
    """The headline numbers, and every fund's gap between risk and capital.

    A fund's bar is how much more, or less, of the portfolio's risk it carries
    than of its money. The long bars are where a decision is worth making: a
    large positive bar is a position bigger in risk than it looks on a
    weights sheet.
    """
    # A fund the portfolio does not hold has no share of either the risk or the
    # capital, so its gap is zero by construction and says nothing. The place
    # to judge those is the marginal chart, not this one.
    table = (risk.funds[risk.funds["weight"] != 0]
             .sort_values("risk_share_gap", ascending=False).reset_index(drop=True))
    unheld = int((risk.funds["weight"] == 0).sum())
    colours = risk.cluster_colours
    height = max(7.0, 3.4 + 0.34 * len(table))
    figure = plt.figure(figsize=(13.5, height))
    grid = figure.add_gridspec(2, 1, height_ratios=[1.35, 2.2 + 0.055 * len(table)])
    headline = figure.add_subplot(grid[0, 0])
    gaps = figure.add_subplot(grid[1, 0])

    headline.axis("off")
    pairs = concentration_lines(risk)
    columns = 4
    for index, (label, value) in enumerate(pairs):
        column, row = index % columns, index // columns
        x = 0.015 + column * (0.985 / columns)
        y = 0.90 - row * 0.42
        headline.text(x, y, value, fontsize=17, fontweight="bold", color="#1a1a1a",
                      transform=headline.transAxes, va="center")
        headline.text(x, y - 0.13, textwrap.fill(label, 24), fontsize=8, color="0.42",
                      transform=headline.transAxes, va="top")
    trailer = (
        f"{len(risk.funds)} funds behaving like {risk.effective_bets:.1f} independent ones. "
        "The effective number of bets counts positions by how much distinct risk they carry, "
        "so two funds that move together count close to one."
    )
    if unheld:
        trailer += (f" {unheld} fund(s) in the window are not held, so they are left off the "
                    "chart below.")
    headline.text(0.015, -0.10, textwrap.fill(trailer, 132), fontsize=8.5, color="0.42",
                  transform=headline.transAxes, va="top")

    positions = np.arange(len(table))
    values = table["risk_share_gap"].to_numpy(dtype=float)
    gaps.barh(positions, values, color=[colours[int(key)] for key in table["cluster"]],
              edgecolor="0.25", linewidth=0.6, height=0.68)
    gaps.axvline(0.0, color="0.35", linewidth=1.0)

    span = float(np.max(np.abs(values))) or 1.0
    for position, row in enumerate(table.itertuples()):
        offset = span * 0.02
        gaps.text(row.risk_share_gap + (offset if row.risk_share_gap >= 0 else -offset),
                  position,
                  f"{format_percent(row.risk_share, 1)} risk vs "
                  f"{format_percent(row.weight_share, 1)} capital",
                  va="center", ha="left" if row.risk_share_gap >= 0 else "right",
                  fontsize=6.8, color="0.35")

    gaps.set_yticks(positions)
    gaps.set_yticklabels(
        [textwrap.fill(f"C{int(row.cluster)}  {row.fund}", 46) for row in table.itertuples()],
        fontsize=7.4)
    gaps.invert_yaxis()
    gaps.set_xlim(-span * 1.55, span * 1.55)
    gaps.xaxis.set_major_formatter(PercentFormatter(xmax=1.0))
    gaps.set_xlabel("Share of portfolio risk minus share of capital")
    gaps.grid(axis="x", linestyle="--", alpha=0.3)
    gaps.set_axisbelow(True)
    gaps.set_title("Funds carrying more, or less, risk than their weight suggests",
                   fontsize=11, pad=10)

    figure.suptitle(f"Portfolio concentration ({risk.timeframe_label})", fontsize=14, y=0.985)
    _allocation_footer(figure, risk, options)
    figure.tight_layout(rect=(0.01, 0.035, 0.99, 0.95))
    figure.subplots_adjust(hspace=0.42)
    return figure


def _save_allocation_figure(figure: plt.Figure, stem: Path, pdf: PdfPages | None,
                            options: dict[str, Any]) -> None:
    if options["save_png"]:
        figure.savefig(stem.with_suffix(".png"), dpi=int(options["dpi"]), bbox_inches="tight")
    if pdf is not None:
        pdf.savefig(figure, dpi=int(options["dpi"]), bbox_inches="tight")
    plt.close(figure)


def allocation_log_lines(risk: PortfolioRisk, options: dict[str, Any]) -> list[str]:
    """What the run report says about one lookback."""
    lines = [
        f"{risk.timeframe_label}: {risk.returns.shape[1]} funds over {len(risk.returns)} months; "
        f"portfolio volatility {format_percent(risk.portfolio_vol)} a year; "
        f"{risk.effective_bets:.1f} effective bets from "
        f"{int((risk.funds['weight'] != 0).sum())} holdings."
    ]
    for row in risk.clusters.itertuples():
        multiplier = (f"{row.risk_multiplier:.2f}x capital"
                      if np.isfinite(row.risk_multiplier) else "not held")
        lines.append(
            f"  C{row.cluster} {row.label}: {format_percent(row.risk_share)} of risk on "
            f"{format_percent(row.weight_share)} of capital ({multiplier})."
        )
    flag = float(options["risk_multiplier_flag"])
    crowded = risk.clusters[risk.clusters["risk_multiplier"] >= flag]
    if not crowded.empty:
        lines.append(
            f"  Carrying more risk than capital by {flag:g}x or more: "
            + ", ".join(f"C{int(row.cluster)} {row.label}" for row in crowded.itertuples())
            + "."
        )
    lines.extend(f"  {note}" for note in risk.notes)
    return lines


def run_allocation_module(data: WorkbookData, options: dict[str, Any],
                          cluster_options: dict[str, Any], output_dir: Path,
                          findings: RunFindings | None = None) -> list[str]:
    """Allocate the portfolio's risk across the clusters module 2 found.

    The clustering settings are shared with module 2 on purpose, so the two
    modules can never disagree about which funds belong together.
    """
    output_dir.mkdir(parents=True, exist_ok=True)
    log: list[str] = []
    timeframes = normalise_timeframes(cluster_options["timeframes"])
    written = False

    for period in timeframes:
        label = timeframe_label(period)
        universe = build_cluster_universe(data, cluster_options, period)
        merged = universe.combined
        if merged.empty or merged.shape[1] < 2 or len(merged) < int(
                cluster_options["min_required_months"]):
            log.append(f"{label}: skipped - too little common history across the strategies.")
            continue
        try:
            analysis = build_cluster_analysis("All Strategies Combined", label, merged,
                                              cluster_options)
            risk = build_portfolio_risk(
                data, merged, analysis, {**options, **{
                    "linkage_method": cluster_options["linkage_method"]}}, label)
        except (ValueError, WorkbookFormatError) as exc:
            log.append(f"{label}: skipped - {exc}")
            continue

        charting = {**options, "linkage_method": cluster_options["linkage_method"]}
        stem = output_dir / f"Portfolio_Risk_{label}"
        pdf_path = output_dir / f"Portfolio_Risk_Report_{label}.pdf"
        with ExitStack() as stack:
            pdf = stack.enter_context(PdfPages(pdf_path)) if options["create_pdf"] else None
            for figure, suffix in (
                (plot_cluster_risk_budget(risk, charting), "cluster_risk_budget"),
                (plot_weighted_dendrogram(risk, charting), "dendrogram_with_weights"),
                (plot_add_trim_scatter(risk, charting), "add_or_trim"),
                (plot_concentration_summary(risk, charting), "concentration"),
            ):
                if findings is not None:
                    findings.capture(f"allocation::{label}::{suffix}", figure)
                _save_allocation_figure(figure, stem.with_name(f"{stem.name}_{suffix}"),
                                        pdf, options)
        if findings is not None:
            findings.risk.append(risk)

        if options["save_csv"]:
            columns = ["strategy", "fund", "label", "cluster", "weight", "weight_share",
                       "expected_return", "standalone_vol", "marginal_risk",
                       "risk_contribution", "risk_share", "risk_multiplier", "risk_share_gap",
                       "correlation_to_portfolio"]
            (risk.funds[columns].sort_values("risk_share", ascending=False)
             .to_csv(stem.with_name(f"{stem.name}_fund_contributions.csv"), index=False))
            risk.clusters.to_csv(stem.with_name(f"{stem.name}_cluster_budget.csv"), index=False)
            pd.DataFrame(concentration_lines(risk), columns=["measure", "value"]).to_csv(
                stem.with_name(f"{stem.name}_summary.csv"), index=False)
            written = True

        log.extend(allocation_log_lines(risk, options))

    if written:
        log.append("Wrote the fund contributions, cluster budget and summary as CSV.")
    if not log:
        log.append("No lookback had enough common history to allocate risk over.")
    return log


# =============================================================================
# The written commentary
#
# Every chart in this pack answers a question, and every chart assumes the
# reader already knows which question. This section writes that down: a short
# standing explanation of what each output is and how to read it, followed by
# what this particular workbook actually says, with the funds named.
#
# The explanations are fixed text, because what a cone chart means does not
# change between runs. The observations are not: they are read off the same
# objects the charts were drawn from, so the document and the pictures beside
# it can never disagree.
#
# The document is built as a list of blocks and rendered twice over - to Word
# through python-docx, and to Markdown when python-docx is not installed. A
# missing package downgrades the file format; it never costs the reader the
# commentary itself.
# =============================================================================

COMMENTARY_STEM = "commentary"


def _block(kind: str, **fields: Any) -> dict[str, Any]:
    """One element of the document, in a form either renderer can take."""
    return {"kind": kind, **fields}


def _headline_of(items: Sequence[Any], months: Callable[[Any], tuple[int, int]]) -> Any | None:
    """The lookback the commentary is written on: the longest, then the widest.

    The longest window is the one least likely to be describing a single
    market episode, so it carries the write-up and the shorter windows are
    used only to say where they disagree with it.
    """
    return max(items, key=months) if items else None


def headline_risk(findings: RunFindings) -> PortfolioRisk | None:
    return _headline_of(findings.risk,
                        lambda risk: (len(risk.returns), risk.returns.shape[1]))


def headline_clusters(findings: RunFindings) -> ClusterAnalysis | None:
    return _headline_of(list(findings.clusters.values()),
                        lambda result: (len(result.returns), result.returns.shape[1]))


def peer_sets(analysis: ClusterAnalysis) -> dict[str, frozenset[str]]:
    """Each fund's cluster-mates, which is comparable across lookbacks when a cluster number is not."""
    members = analysis.assignments.groupby("cluster")["fund"].apply(set).to_dict()
    return {str(row.fund): frozenset(members[row.cluster] - {row.fund})
            for row in analysis.assignments.itertuples()}


def cluster_movers(findings: RunFindings, headline: ClusterAnalysis) -> list[tuple[str, str]]:
    """Funds that sit with a different crowd over a different window.

    Cluster numbers are not comparable between two runs of the linkage, so the
    comparison is between the sets of funds a fund shares its cluster with. A
    fund that keeps less than half its neighbours has not been placed
    consistently, and any conclusion drawn about it is a conclusion about the
    window rather than about the fund.
    """
    reference = peer_sets(headline)
    moved: dict[str, list[str]] = {}
    for label, analysis in findings.clusters.items():
        if label == headline.timeframe_label:
            continue
        other = peer_sets(analysis)
        for fund, peers in reference.items():
            if fund not in other:
                continue
            union = peers | other[fund]
            overlap = len(peers & other[fund]) / len(union) if union else 1.0
            if overlap < 0.5:
                moved.setdefault(fund, []).append(label)
    return [(_split_label(fund)[1] or fund, _names(labels, 3))
            for fund, labels in sorted(moved.items())]


def tightest_pairs(analysis: ClusterAnalysis, count: int = 3) -> list[tuple[str, str, float]]:
    """The most closely correlated pairs, which are the portfolio's near-duplicates."""
    matrix = analysis.correlation
    pairs: list[tuple[str, str, float]] = []
    for first in range(matrix.shape[0]):
        for second in range(first + 1, matrix.shape[1]):
            value = float(matrix.iat[first, second])
            if np.isfinite(value):
                pairs.append((str(matrix.index[first]), str(matrix.columns[second]), value))
    pairs.sort(key=lambda item: -item[2])
    return pairs[:count]


# -----------------------------------------------------------------------------
# The standing explanations
#
# These do not change between runs, because what a cone chart means does not
# change between runs. They are kept here rather than in the chart footers so
# that the reader who was not in the room gets the same explanation as the
# reader who was.
# -----------------------------------------------------------------------------

EXPLAINERS: dict[str, list[str]] = {
    "cone": [
        "The cone charts test each fund against what it was underwritten to do. Row 3 of the "
        "workbook holds the annualised excess return the fund is expected to earn and row 4 the "
        "volatility it is expected to earn it with. Those two numbers imply a range of cumulative "
        "excess returns over time - narrow at the start, widening with the square root of elapsed "
        "months - and that range is the cone. The fund's own cumulative excess return over the "
        "risk-free rate is drawn through it.",
        "A line inside the cone means the fund is doing roughly what it said it would, at roughly "
        "the risk it said it would take. A line that leaves the bottom of the cone is a shortfall "
        "against the case for owning it, and one that leaves the top is not automatically good "
        "news: it usually means the fund is running more risk than row 4 assumed, which the "
        "rolling volatility panel beside it will show. The rolling total return, rolling Sharpe "
        "and drawdown charts on the same page put that path in context - whether a shortfall came "
        "from one bad quarter or from a slow bleed, and whether the fund's risk-adjusted return "
        "is trending up or down.",
        "The cone is only as good as the two assumptions behind it. A fund given a generous "
        "volatility assumption gets a wide cone and can sit comfortably inside it while "
        "delivering very little, so the width of the cone is worth reading before the position of "
        "the line within it.",
    ],
    "clustering": [
        "The clustering output answers a question a list of strategy labels cannot: which of these "
        "funds actually behave alike? The heatmap shows the correlation of every pair of funds "
        "over the window. The dendrogram beside it groups the funds by how similar their monthly "
        "return patterns are, joining the closest pair first and working up until every fund is on "
        "one tree.",
        "Read the dendrogram from the leaves upward. The height at which two funds join is how "
        "different they are: funds joining near the bottom are near-substitutes for one another, "
        "and a branch that only joins the rest of the tree near the top is genuinely doing "
        "something else. Cutting the tree at a chosen number of clusters turns that picture into "
        "the groups used everywhere else in this pack. A cluster is a statement about behaviour "
        "over this window, not about a fund's stated strategy - two funds sitting in different "
        "strategy buckets in the workbook can quite legitimately land in the same cluster, and "
        "that is usually the most useful thing on the chart.",
        "Correlation measured over one window is not a stable quantity, and a short window is "
        "noisier than a long one, which is why several lookbacks are produced rather than one. "
        "Funds without enough common history are left out of the clustering altogether rather "
        "than being compared on a handful of months; they are listed with their reason in the "
        "exclusion audit, and they are not represented anywhere on these charts.",
    ],
    "performance": [
        "The performance and drawdown output is the record of what actually happened, without any "
        "assumption attached. The NAV lines rebase every fund to 100 at the start of the window "
        "so that paths can be compared on one axis. The drawdown lines show how far each fund is "
        "below its own high-water mark at every point, which is the number an investor actually "
        "experiences.",
        "The heatmaps put funds on one axis and months on the other. Read them down a column "
        "rather than along a row: a column that turns dark across many funds at once is a shared "
        "shock, and it is where the correlations from the clustering section stop being an "
        "abstraction. Reading along a row shows whether a fund's difficulty was its own. The "
        "rolling performance and rolling risk-adjusted-return heatmaps do the same for returns "
        "and for return per unit of volatility over the trailing window.",
        "Two limits are worth holding in mind. A drawdown depends on when the window starts, so a "
        "fund whose worst period sits just before the window looks better here than it deserves. "
        "And monthly data cannot see inside a month, so every drawdown on these charts is the "
        "month-end path and understates what was live at the time.",
    ],
    "allocation": [
        "Weight is not risk. Three managed-futures funds at 5% each are a smaller block of risk "
        "than their 15% of the book suggests, because they diversify one another; three "
        "multi-strategy funds at 5% each are a larger one. This section splits the portfolio's "
        "volatility across the clusters from the previous section, so a cluster on these charts "
        "is the same cluster on that dendrogram.",
        "The decomposition is the standard one. Each fund's marginal contribution is how much the "
        "portfolio's volatility moves for the next pound put into it, and its component "
        "contribution is that marginal figure multiplied by what is already held. Those "
        "components add up to the portfolio's volatility exactly, so they can be grouped by "
        "cluster with nothing left over, and a cluster's multiplier is simply its share of the "
        "risk divided by its share of the capital. Above 1 it is carrying more risk than money.",
        "The add-or-trim chart plots each fund's marginal contribution to risk against the return "
        "it is expected to earn. The diagonal is the portfolio's own return per unit of risk: a "
        "fund above the line improves that ratio at the margin, a fund below it is being paid "
        "less than the risk it adds, and the portfolio itself sits on the line by construction. "
        "Funds drawn as open rings are not held, so a candidate can be judged on the same axes "
        "before it is bought. These are statements about the maths of the current book over one "
        "window, not recommendations - a fund below the line may be held for a reason this "
        "arithmetic does not see, such as capacity, liquidity or a view about the next regime.",
    ],
}

HOW_TO_READ = [
    "This document is written alongside the charts in the same output folder, and it is not a "
    "substitute for them. Each section explains what its output is and how to read it, and then "
    "says what this particular workbook shows, naming the funds involved. The explanations are "
    "the same in every run; the observations are read from the same data the charts were drawn "
    "from.",
    "Everything here is descriptive. The figures come from the monthly returns in the workbook "
    "over the lookbacks chosen for this run, using the estimators named at the end of the "
    "document, and they describe how the book has behaved rather than how it will. Where a "
    "sentence points at a decision - a cluster carrying more risk than its capital, a fund whose "
    "next pound improves the portfolio's return per unit of risk - it is describing what the "
    "arithmetic implies, on the assumption that the window is representative. It is not a "
    "recommendation, and it does not know about capacity, liquidity, lock-ups, fees or a manager "
    "meeting held last week.",
]


# -----------------------------------------------------------------------------
# What this workbook says
# -----------------------------------------------------------------------------

def risk_observations(risk: PortfolioRisk, options: dict[str, Any],
                      gap: float, limit: int) -> tuple[list[str], list[str], list[str]]:
    """Where the risk sits, what could be added or trimmed, and what else stands out."""
    funds = risk.funds
    held = funds[funds["weight"] != 0]
    clusters = risk.clusters

    where: list[str] = []
    top_cluster = clusters.iloc[0]
    members = funds[funds["cluster"] == top_cluster.cluster].sort_values(
        "risk_share", ascending=False)
    where.append(
        f"The largest single block of risk is cluster C{int(top_cluster.cluster)} "
        f"({top_cluster.label}), which carries {format_percent(top_cluster.risk_share)} of the "
        f"portfolio's volatility on {format_percent(top_cluster.weight_share)} of its capital"
        + (f" - {top_cluster.risk_multiplier:.2f} times its weight"
           if np.isfinite(top_cluster.risk_multiplier) else "")
        + ". Its holdings are "
        + _names([_fund_name(row) for row in members.itertuples() if row.weight], limit) + "."
    )
    ranked = held.nlargest(min(3, len(held)), "risk_share")
    where.append(
        "The three largest contributors at fund level are "
        + "; ".join(
            f"{_fund_name(row)} at {format_percent(row.risk_share)} of the risk on "
            f"{format_percent(row.weight_share)} of the capital"
            for row in ranked.itertuples())
        + f". Between them they account for {format_percent(ranked['risk_share'].sum())} of "
          "the portfolio's volatility."
    )
    crowded = held[held["risk_share_gap"] >= gap].sort_values("risk_share_gap", ascending=False)
    if not crowded.empty:
        where.append(
            "Positions carrying materially more risk than their weight implies: "
            + _trimmed(
                [f"{_fund_name(row)} (+{format_percent(row.risk_share_gap)})"
                 for row in crowded.itertuples()], limit)
            + ". These are the positions that are larger than they look on a weights sheet, "
              "either because they are volatile on their own or because they move with the rest "
              "of the book."
        )
    diversifying = held[held["risk_share_gap"] <= -gap].sort_values("risk_share_gap")
    if not diversifying.empty:
        where.append(
            "Positions carrying materially less risk than their weight implies: "
            + _trimmed(
                [f"{_fund_name(row)} ({format_percent(row.risk_share_gap)})"
                 for row in diversifying.itertuples()], limit)
            + ". These are doing the diversifying work in the portfolio."
        )
    flag = float(options["risk_multiplier_flag"])
    hot = clusters[clusters["risk_multiplier"] >= flag]
    if len(hot) > 1:
        where.append(
            f"More than one cluster is running above the {flag:g}x crowding flag: "
            + _names([f"C{int(row.cluster)} ({row.label}) at {row.risk_multiplier:.2f}x"
                      for row in hot.itertuples()], limit)
            + "."
        )

    # Add or trim. The diagonal on the chart is the portfolio's own return per
    # unit of risk, so a fund clears it when its expected return exceeds its
    # marginal risk multiplied by that ratio.
    moves: list[str] = []
    slope = risk.portfolio_return / risk.portfolio_vol if risk.portfolio_vol > 0 else np.nan
    if np.isfinite(slope):
        scored = funds.assign(
            edge=funds["expected_return"] - slope * funds["marginal_risk"])
        above_held = scored[(scored["weight"] != 0) & (scored["edge"] > 0)].sort_values(
            "edge", ascending=False)
        below_held = scored[(scored["weight"] != 0) & (scored["edge"] < 0)].sort_values("edge")
        candidates = scored[(scored["weight"] == 0) & (scored["edge"] > 0)].sort_values(
            "edge", ascending=False)
        moves.append(
            f"The portfolio earns {slope:.2f} of expected return for each unit of volatility it "
            f"runs ({format_percent(risk.portfolio_return)} on "
            f"{format_percent(risk.portfolio_vol)}). That ratio is the bar each fund is measured "
            "against below: a fund clears it when its expected return is worth more than the "
            "risk the next pound into it would add."
        )
        if not above_held.empty:
            moves.append(
                "Held funds where the next pound improves that ratio, best first: "
                + _trimmed(
                    [f"{_fund_name(row)} (currently {format_percent(row.weight)}, expected "
                     f"{format_percent(row.expected_return)} against a marginal risk of "
                     f"{format_percent(row.marginal_risk)})"
                     for row in above_held.itertuples()], limit)
                + "."
            )
        if not below_held.empty:
            moves.append(
                "Held funds being paid less than the risk they add, where trimming improves the "
                "ratio: "
                + _trimmed(
                    [f"{_fund_name(row)} (currently {format_percent(row.weight)}, expected "
                     f"{format_percent(row.expected_return)} against a marginal risk of "
                     f"{format_percent(row.marginal_risk)})"
                     for row in below_held.itertuples()], limit)
                + ". Trimming the largest of these is the cheapest way to reduce portfolio "
                  "volatility without giving up much expected return."
            )
        if not candidates.empty:
            moves.append(
                "Funds in the workbook that are not held and would clear the bar at the margin: "
                + _trimmed(
                    [f"{_fund_name(row)} (expected {format_percent(row.expected_return)}, "
                     f"marginal risk {format_percent(row.marginal_risk)}, correlation to the "
                     f"book {row.correlation_to_portfolio:.2f})"
                     for row in candidates.itertuples()], limit)
                + ". A low correlation to the book is what makes a candidate cheap to add, "
                  "rather than a high expected return on its own."
            )
        unheld_below = scored[(scored["weight"] == 0) & (scored["edge"] <= 0)]
        if not unheld_below.empty:
            moves.append(
                f"{_plural(len(unheld_below), 'unheld fund')} in the window would not clear the "
                "bar on these numbers: " + _names(
                    [_fund_name(row) for row in unheld_below.itertuples()], limit) + "."
            )

    notable: list[str] = [
        f"The book behaves like {risk.effective_bets:.1f} independent positions, from "
        f"{int((funds['weight'] != 0).sum())} holdings across {len(clusters)} clusters. Its "
        f"diversification ratio is {risk.diversification_ratio:.2f}, meaning the funds' own "
        f"volatilities, weighted and added up, come to {risk.diversification_ratio:.2f} times the "
        f"{format_percent(risk.portfolio_vol)} the portfolio actually runs."
    ]
    top_five = funds.nlargest(min(5, len(funds)), "risk_share")["risk_share"].sum()
    notable.append(
        f"The five largest contributors hold {format_percent(top_five)} of the portfolio's risk. "
        f"Concentration of risk measured on a Herfindahl basis is {risk.risk_concentration:.3f} "
        f"against {risk.weight_concentration:.3f} for capital"
        + (", so the risk is more concentrated than the weights sheet suggests."
           if risk.risk_concentration > risk.weight_concentration
           else ", so the weights are more concentrated than the risk they produce.")
    )
    quiet = clusters[clusters["risk_multiplier"] < 1.0]
    if not quiet.empty:
        notable.append(
            "Clusters carrying less risk than capital: "
            + _names([f"C{int(row.cluster)} ({row.label}) at {row.risk_multiplier:.2f}x"
                      for row in quiet.itertuples()], limit)
            + ". Room to add sits here before it sits anywhere else."
        )
    if risk.cash_weight > 0.0005:
        notable.append(
            f"{format_percent(risk.cash_weight)} of the book is in cash, which carries no risk "
            "and is excluded from every share above."
        )
    elif risk.cash_weight < -0.0005:
        notable.append(
            f"The weights total {format_percent(1.0 - risk.cash_weight)}, which is read as "
            f"{format_percent(-risk.cash_weight)} of gearing."
        )
    return where, moves, notable


def cone_observations(records: list[dict[str, Any]], gap: float,
                      limit: int) -> list[str]:
    """Which funds are running behind, ahead of, or hotter than what was underwritten."""
    rated = [row for row in records
             if row["expected_excess"] is not None and np.isfinite(row["realised_excess"])]
    lines: list[str] = []
    if not rated:
        lines.append(
            "No fund in the window has an expected excess return on row 3, so no fund can be "
            "measured against its own case. The rolling return, Sharpe and drawdown charts still "
            "apply; the cones do not."
        )
        return lines

    for row in rated:
        row["shortfall"] = row["realised_excess"] - float(row["expected_excess"])
    behind = sorted([row for row in rated if row["shortfall"] <= -gap],
                    key=lambda row: row["shortfall"])
    ahead = sorted([row for row in rated if row["shortfall"] >= gap],
                   key=lambda row: -row["shortfall"])

    if behind:
        lines.append(
            f"{_plural(len(behind), 'fund is', 'funds are')} running behind what was "
            f"underwritten by more than {format_percent(gap)} a year, worst first: "
            + _trimmed(
                [f"{row['fund']} (expected {format_percent(row['expected_excess'])}, delivered "
                 f"{format_percent(row['realised_excess'])} over "
                 f"{_plural(row['months'], 'month')})"
                 for row in behind], limit)
            + "."
        )
    else:
        lines.append(
            f"No fund is more than {format_percent(gap)} a year behind its expected excess "
            "return over its window."
        )
    if ahead:
        lines.append(
            "Ahead of expectations by the same margin: "
            + _trimmed(
                [f"{row['fund']} (expected {format_percent(row['expected_excess'])}, delivered "
                 f"{format_percent(row['realised_excess'])})" for row in ahead], limit)
            + ". Worth checking against realised volatility before it is read as skill."
        )

    hotter = [row for row in rated
              if row["expected_vol"] is not None and np.isfinite(row["realised_vol"])
              and row["realised_vol"] > float(row["expected_vol"]) * 1.25]
    if hotter:
        hotter.sort(key=lambda row: -(row["realised_vol"] / float(row["expected_vol"])))
        lines.append(
            "Running materially more volatility than row 4 assumed - at least a quarter more: "
            + _trimmed(
                [f"{row['fund']} ({format_percent(row['realised_vol'])} realised against "
                 f"{format_percent(float(row['expected_vol']))} expected)" for row in hotter],
                limit)
            + ". A fund inside its cone on this much extra risk is not inside it on merit."
        )
    calmer = [row for row in rated
              if row["expected_vol"] is not None and np.isfinite(row["realised_vol"])
              and row["realised_vol"] < float(row["expected_vol"]) * 0.75]
    if calmer:
        lines.append(
            "Running materially less volatility than assumed: "
            + _names([f"{row['fund']} ({format_percent(row['realised_vol'])} against "
                      f"{format_percent(float(row['expected_vol']))})" for row in calmer], limit)
            + "."
        )
    missing = [row["fund"] for row in records if row["expected_excess"] is None]
    if missing:
        lines.append(
            f"{_plural(len(missing), 'fund has', 'funds have')} no expected figures in the "
            "workbook and so get no cone: " + _names(missing, limit) + "."
        )
    return lines


def cluster_observations(headline: ClusterAnalysis, findings: RunFindings,
                         limit: int) -> tuple[list[str], list[str]]:
    """What the clustering shows on the headline window, and where the others disagree."""
    assignments = headline.assignments
    sizes = assignments.groupby("cluster")["fund"].apply(list)
    lines = [
        f"Over the {headline.timeframe_label} window, {headline.returns.shape[1]} funds with a "
        f"complete common history split into {len(sizes)} clusters "
        f"({headline.returns.index.min().strftime('%b %Y')} to "
        f"{headline.returns.index.max().strftime('%b %Y')})."
    ]
    for cluster, members in sizes.items():
        names = [_split_label(str(member))[1] for member in members]
        strategies = sorted({_split_label(str(member))[0] for member in members})
        mixed = (" - drawn from " + _names(strategies, 4) + ", so this grouping cuts across the "
                 "workbook's own strategy labels") if len(strategies) > 1 else ""
        lines.append(
            f"C{int(cluster)} ({_plural(len(names), 'fund')}{mixed}): {_names(names, limit)}.")

    pairs = tightest_pairs(headline)
    if pairs:
        lines.append(
            "The most closely correlated pairs over this window are "
            + "; ".join(f"{_split_label(first)[1]} and {_split_label(second)[1]} at "
                        f"{value:.2f}" for first, second, value in pairs)
            + ". A pair above about 0.8 is close to being one position held twice."
        )

    disagreement: list[str] = []
    others = [label for label in findings.clusters if label != headline.timeframe_label]
    if others:
        movers = cluster_movers(findings, headline)
        if movers:
            disagreement.append(
                "These funds sit with a materially different set of peers over the other "
                "lookbacks, so their placement is a statement about the window rather than about "
                "the fund: "
                + "; ".join(f"{fund} (over {labels})" for fund, labels in movers[:limit])
                + "."
            )
        else:
            disagreement.append(
                f"The grouping is stable: no fund changes the bulk of its cluster-mates between "
                f"the {headline.timeframe_label} window and {_names(others, 3)}."
            )
        for label in others:
            other = findings.clusters[label]
            extra = other.returns.shape[1] - headline.returns.shape[1]
            if extra > 0:
                disagreement.append(
                    f"The {label} window covers {_plural(extra, 'more fund')} than the "
                    f"{headline.timeframe_label} window, because a shorter window asks less "
                    "history of each fund, and nothing said above covers the difference."
                )
    return lines, disagreement


def performance_observations(records: list[dict[str, Any]], limit: int) -> list[str]:
    """The drawdown record: how deep, how long, and who is still in one."""
    if not records:
        return ["No fund had enough history to produce a drawdown record."]
    lines: list[str] = []
    deepest = sorted(records, key=lambda row: row["max_drawdown"])[:3]
    lines.append(
        "The deepest drawdowns over each fund's own history are "
        + "; ".join(
            f"{row['fund']} at {format_percent(row['max_drawdown'])} "
            f"(trough {row['trough'].strftime('%b %Y')})"
            for row in deepest)
        + "."
    )
    underwater = sorted([row for row in records if row["current_drawdown"] < -0.005],
                        key=lambda row: row["current_drawdown"])
    if underwater:
        lines.append(
            f"{len(underwater)} of {len(records)} funds are below their high-water mark at the "
            "end of the window: "
            + _trimmed(
                [f"{row['fund']} at {format_percent(row['current_drawdown'])} after "
                 f"{_plural(row['months_underwater'], 'month')}"
                 for row in underwater], limit)
            + "."
        )
        stuck = [row for row in underwater if row["months_underwater"] >= 24]
        if stuck:
            lines.append(
                "Underwater for two years or more: "
                + _names([f"{row['fund']} ({row['months_underwater']} months)"
                          for row in stuck], limit)
                + ". A shallow drawdown that has lasted this long is a different problem from a "
                  "deep one that is already recovering."
            )
        worst_now = [row for row in underwater if row["at_worst_now"]]
        if worst_now:
            lines.append(
                "At the worst point of their recorded history right now: "
                + _names([row["fund"] for row in worst_now], limit)
                + "."
            )
    else:
        lines.append("Every fund finished the window at or near its high-water mark.")
    return lines


# -----------------------------------------------------------------------------
# Assembling the document
# -----------------------------------------------------------------------------

def build_commentary(data: WorkbookData, settings: dict[str, Any],
                     findings: RunFindings) -> list[dict[str, Any]]:
    """The whole document as blocks, ready for either renderer."""
    options = settings["commentary"]
    limit = int(options["max_named_funds"])
    gap = float(options["material_gap_pct"]) / 100.0
    modules = selected_modules(settings)
    blocks: list[dict[str, Any]] = [
        _block("title", text="Portfolio analytics commentary"),
        _block("subtitle", text=(
            f"{Path(str(data.path)).name} - {data.sheet} - "
            f"{data.index[0]:%b %Y} to {data.index[-1]:%b %Y} - "
            f"{len(data.funds)} funds across {len(data.strategies)} strategies - "
            f"produced {_dt.datetime.now():%d %B %Y}")),
    ]

    risk = headline_risk(findings)
    clusters = headline_clusters(findings)

    where: list[str] = []
    moves: list[str] = []
    notable: list[str] = []
    if risk is not None:
        where, moves, notable = risk_observations(risk, settings["allocation"], gap, limit)

    blocks.append(_block("heading", level=1, text="How to read this document"))
    blocks.extend(_block("para", text=text) for text in HOW_TO_READ)
    if risk is not None:
        blocks.append(_block("para", text=(
            f"The risk figures are written on the {risk.timeframe_label} window, which is the "
            f"longest one this run could use ({len(risk.returns)} months to "
            f"{risk.returns.index.max().strftime('%b %Y')}, "
            f"{risk.returns.shape[1]} funds). Where a shorter window tells a different story it "
            "is called out rather than averaged in.")))
    elif "allocation" not in modules:
        blocks.append(_block("note", text=(
            "The portfolio risk module did not run, so this document describes how the funds have "
            "behaved but not how the book's risk is distributed between them. Add a weights row "
            "to the worksheet and tick module 4 to get the risk allocation, the add-or-trim "
            "reading and the concentration measures.")))

    # ---- Headline ---------------------------------------------------------
    headline = headline_lines(findings, risk, gap)
    if headline:
        blocks.append(_block("heading", level=1, text="Headline"))
        blocks.append(_block("bullets", items=headline))

    # ---- Module 4: where the risk is --------------------------------------
    if risk is not None:
        blocks.append(_block("heading", level=1, text="Where the risk is"))
        blocks.extend(_block("para", text=text) for text in EXPLAINERS["allocation"][:2])
        blocks.append(_block("heading", level=2, text="What this portfolio shows"))
        blocks.append(_block("table", headers=["Measure", "Value"],
                            rows=[list(pair) for pair in concentration_lines(risk)]))
        blocks.append(_block("bullets", items=where))
        blocks.append(_block("image", key=f"allocation::{risk.timeframe_label}::cluster_risk_budget",
                            caption=f"Cluster risk budget, {risk.timeframe_label}"))
        blocks.append(_block("table",
                            headers=["Cluster", "Funds", "Held", "Capital", "Risk", "Multiplier"],
                            rows=[[f"C{int(row.cluster)} {row.label}", str(int(row.members)),
                                   str(int(row.held)), format_percent(row.weight_share),
                                   format_percent(row.risk_share),
                                   f"{row.risk_multiplier:.2f}x"
                                   if np.isfinite(row.risk_multiplier) else "not held"]
                                  for row in risk.clusters.itertuples()]))

        blocks.append(_block("heading", level=1, text="Where funds could be added or trimmed"))
        blocks.append(_block("para", text=EXPLAINERS["allocation"][2]))
        blocks.append(_block("bullets", items=moves))
        blocks.append(_block("image", key=f"allocation::{risk.timeframe_label}::add_or_trim",
                            caption=f"Marginal risk against expected return, "
                                    f"{risk.timeframe_label}"))

        blocks.append(_block("heading", level=1, text="Anything else notable"))
        blocks.append(_block("bullets", items=notable))
        blocks.append(_block("image", key=f"allocation::{risk.timeframe_label}::concentration",
                            caption=f"Concentration, {risk.timeframe_label}"))

        others = [item for item in findings.risk if item is not risk]
        if others:
            differences = lookback_differences(risk, others, gap, limit)
            blocks.append(_block("heading", level=2, text="Where the other lookbacks disagree"))
            blocks.append(_block("bullets", items=differences))

    # ---- Module 1 ---------------------------------------------------------
    if findings.cone:
        blocks.append(_block("heading", level=1, text="Expected against actual"))
        blocks.extend(_block("para", text=text) for text in EXPLAINERS["cone"])
        blocks.append(_block("heading", level=2, text="What this workbook shows"))
        lines = cone_observations(findings.cone, gap, limit)
        blocks.append(_block("bullets", items=lines))
        worst = worst_cone_fund(findings.cone, gap)
        if worst is not None:
            blocks.append(_block("image", key=f"cone::{worst}",
                                caption=f"{worst} against its expected range"))

    # ---- Module 2 ---------------------------------------------------------
    if clusters is not None:
        blocks.append(_block("heading", level=1, text="What behaves like what"))
        blocks.extend(_block("para", text=text) for text in EXPLAINERS["clustering"])
        blocks.append(_block("heading", level=2, text="What this workbook shows"))
        lines, disagreement = cluster_observations(clusters, findings, limit)
        blocks.append(_block("bullets", items=lines))
        blocks.append(_block("image", key=f"dendrogram::{clusters.timeframe_label}",
                            caption=f"All strategies, {clusters.timeframe_label}"))
        if disagreement:
            blocks.append(_block("heading", level=2, text="Where the other lookbacks disagree"))
            blocks.append(_block("bullets", items=disagreement))

    # ---- Module 3 ---------------------------------------------------------
    if findings.performance:
        blocks.append(_block("heading", level=1, text="What actually happened"))
        blocks.extend(_block("para", text=text) for text in EXPLAINERS["performance"])
        blocks.append(_block("heading", level=2, text="What this workbook shows"))
        blocks.append(_block("bullets",
                            items=performance_observations(findings.performance, limit)))
        blocks.append(_block("image", key="performance_drawdown_heatmap",
                            caption="Actual drawdowns, all funds"))
        blocks.append(_block("image", key="performance_nav",
                            caption="NAV, all funds"))

    # ---- Coverage ---------------------------------------------------------
    blocks.append(_block("heading", level=1, text="Coverage and caveats"))
    blocks.append(_block("bullets", items=coverage_notes(data, settings, findings, risk, limit)))
    blocks.append(_block("note", text=(
        "Produced automatically from the workbook named at the top of this document. The "
        "observations are arithmetic on the monthly returns in it, over the windows and with the "
        "estimators listed above. They describe the record, not the future, and they are not "
        "investment advice.")))
    return blocks


def headline_lines(findings: RunFindings, risk: PortfolioRisk | None,
                   gap: float) -> list[str]:
    """The half-dozen sentences a reader who stops after the first page should have.

    Deliberately written rather than assembled from the sections below, so that
    the first page is a summary and not the same paragraphs twice.
    """
    lines: list[str] = []
    if risk is not None:
        held = risk.funds[risk.funds["weight"] != 0]
        lines.append(
            f"The book runs {format_percent(risk.portfolio_vol)} of volatility a year across "
            f"{_plural(len(held), 'holding')}, and behaves like {risk.effective_bets:.1f} "
            f"independent positions - so the {_plural(len(held), 'position')} on the weights "
            f"sheet are worth rather fewer than that in risk terms.")
        top = risk.clusters.iloc[0]
        lines.append(
            f"The largest block of risk is C{int(top.cluster)} ({top.label}), which holds "
            f"{format_percent(top.risk_share)} of it on {format_percent(top.weight_share)} of "
            "the capital"
            + (f", {top.risk_multiplier:.2f} times its weight." if np.isfinite(top.risk_multiplier)
               else "."))
        if not held.empty:
            largest = held.nlargest(1, "risk_share").iloc[0]
            lines.append(
                f"The single largest contributor is {largest['fund']}, at "
                f"{format_percent(largest['risk_share'])} of the portfolio's risk on "
                f"{format_percent(largest['weight_share'])} of its capital.")
        slope = risk.portfolio_return / risk.portfolio_vol if risk.portfolio_vol > 0 else np.nan
        if np.isfinite(slope):
            scored = risk.funds.assign(
                edge=risk.funds["expected_return"] - slope * risk.funds["marginal_risk"])
            best = scored[scored["edge"] > 0].nlargest(1, "edge")
            worst = scored[(scored["weight"] != 0) & (scored["edge"] < 0)].nsmallest(1, "edge")
            parts: list[str] = []
            if not best.empty:
                row = best.iloc[0]
                parts.append(
                    f"the next pound is best spent on {row['fund']}"
                    + ("" if row["weight"] else " (not currently held)"))
            if not worst.empty:
                parts.append(f"the position paying least for its risk is {worst.iloc[0]['fund']}")
            if parts:
                lines.append(
                    f"On the portfolio's own return per unit of risk ({slope:.2f}): "
                    + "; ".join(parts) + ". The add-or-trim section sets out the rest.")

    clusters = headline_clusters(findings)
    if risk is None and clusters is not None:
        # With module 4 off there is no risk budget to lead with, so the
        # grouping itself is the most useful thing on the first page.
        sizes = clusters.assignments.groupby("cluster")["fund"].apply(list)
        largest = max(sizes.items(), key=lambda item: len(item[1]))
        pairs = tightest_pairs(clusters, 1)
        opening = (f"Over the {clusters.timeframe_label} window, {clusters.returns.shape[1]} "
                   "funds with a complete common history fall into "
                   f"{_plural(len(sizes), 'behavioural group')}. ")
        lines.append(opening + (
            "No two of them group together at this cut, so on these returns each is doing "
            "something distinct enough to stand on its own."
            if len(largest[1]) < 2 else
            f"The largest is C{int(largest[0])}, holding {_plural(len(largest[1]), 'fund')}: "
            + _names([_split_label(str(member))[1] for member in largest[1]], 6) + "."))
        if pairs:
            first, second, value = pairs[0]
            lines.append(
                f"The closest pair over that window is {_split_label(first)[1]} and "
                f"{_split_label(second)[1]}, correlated at {value:.2f}.")

    if findings.cone:
        rated = [row for row in findings.cone
                 if row["expected_excess"] is not None and np.isfinite(row["realised_excess"])]
        behind = [row for row in rated
                  if row["realised_excess"] - float(row["expected_excess"]) <= -gap]
        if behind:
            worst_fund = min(behind,
                             key=lambda row: row["realised_excess"] - float(row["expected_excess"]))
            lines.append(
                f"{_plural(len(behind), 'of the rated funds is', 'of the rated funds are')} "
                f"running more than {format_percent(gap)} a year behind what they were "
                f"underwritten to deliver, of {len(rated)} with an expected figure. The widest "
                f"gap is {worst_fund['fund']}, at "
                f"{format_percent(worst_fund['realised_excess'])} against "
                f"{format_percent(float(worst_fund['expected_excess']))} expected.")

    if findings.performance:
        underwater = [row for row in findings.performance if row["current_drawdown"] < -0.005]
        deepest = min(findings.performance, key=lambda row: row["max_drawdown"])
        lines.append(
            f"{len(underwater)} of {len(findings.performance)} funds end the window below their "
            f"high-water mark. The deepest drawdown on record is {deepest['fund']} at "
            f"{format_percent(deepest['max_drawdown'])}.")
    return lines


def worst_cone_fund(records: list[dict[str, Any]], gap: float) -> str | None:
    """The fund whose picture is worth putting in front of the reader."""
    rated = [row for row in records
             if row["expected_excess"] is not None and np.isfinite(row["realised_excess"])]
    if not rated:
        return None
    worst = min(rated, key=lambda row: row["realised_excess"] - float(row["expected_excess"]))
    if worst["realised_excess"] - float(worst["expected_excess"]) > -gap:
        return None
    return str(worst["label"])


def lookback_differences(headline: PortfolioRisk, others: Sequence[PortfolioRisk],
                         gap: float, limit: int) -> list[str]:
    """Where a shorter window would change the reading above."""
    lines: list[str] = []
    for other in others:
        parts: list[str] = []
        vol_gap = other.portfolio_vol - headline.portfolio_vol
        if abs(vol_gap) >= 0.01:
            parts.append(
                f"portfolio volatility is {format_percent(other.portfolio_vol)} rather than "
                f"{format_percent(headline.portfolio_vol)}")
        if abs(other.effective_bets - headline.effective_bets) >= 0.5:
            parts.append(
                f"the book behaves like {other.effective_bets:.1f} independent positions rather "
                f"than {headline.effective_bets:.1f}")

        # Cluster numbers are not comparable between two runs of the linkage,
        # so the comparison that survives is at fund level.
        left = headline.funds.set_index("label")["risk_share"]
        right = other.funds.set_index("label")["risk_share"]
        shared = left.index.intersection(right.index)
        moved = (right[shared] - left[shared]).sort_values(key=abs, ascending=False)
        moved = moved[moved.abs() >= gap]
        if not moved.empty:
            parts.append(
                "the biggest changes in risk share are "
                + "; ".join(
                    f"{_split_label(str(name))[1]} "
                    f"({format_percent(left[name])} to {format_percent(right[name])})"
                    for name in moved.index[:limit]))
        missing = len(left.index.difference(right.index))
        extra = len(right.index.difference(left.index))
        if extra or missing:
            parts.append(
                f"it covers {other.returns.shape[1]} funds rather than "
                f"{headline.returns.shape[1]}")
        lines.append(
            f"Over {other.timeframe_label}: " + ("; ".join(parts) if parts else
                                                 "nothing material changes") + ".")
    return lines


def coverage_notes(data: WorkbookData, settings: dict[str, Any], findings: RunFindings,
                   risk: PortfolioRisk | None, limit: int) -> list[str]:
    """What the numbers above do not cover, and what shaped them."""
    lines: list[str] = []
    lines.append(
        f"The workbook holds {len(data.funds)} funds across {len(data.strategies)} strategies, "
        f"with monthly returns from {data.index[0]:%b %Y} to {data.index[-1]:%b %Y} "
        f"({len(data.index)} months).")

    histories = sorted(((len(fund.returns.dropna()), fund.fund) for fund in data.funds))
    short = [f"{name} ({months} months)" for months, name in histories[:limit] if months < 36]
    if short:
        lines.append(
            "Shortest track records, where every estimate is least reliable: "
            + _names(short, limit) + ".")

    if findings.cluster_exclusions:
        reasons: dict[str, set[str]] = {}
        for row in findings.cluster_exclusions:
            reasons.setdefault(str(row.get("reason", "unstated")), set()).add(str(row.get("fund")))
        for reason, names in sorted(reasons.items()):
            lines.append(
                f"Left out of the clustering ({reason}): " + _names(sorted(names), limit) + ".")

    if findings.performance_excluded:
        names = sorted({entry[0].split(" - ", 1)[-1] for entry in findings.performance_excluded})
        lines.append(
            "Left off the line charts for want of history: " + _names(names, limit) + ".")

    if risk is not None:
        lines.extend(risk.notes)
    if data.notes:
        lines.extend(f"Data note: {note}" for note in data.notes)

    clustering = settings["clustering"]
    lines.append(
        "Clustering settings: "
        f"{clustering['correlation_method']} correlation, {clustering['linkage_method']} linkage, "
        f"cut into at most {clustering['max_clusters']} clusters, over "
        f"{_names([timeframe_label(period) for period in normalise_timeframes(clustering['timeframes'])], 4)}.")
    if risk is not None:
        allocation = settings["allocation"]
        estimator = ("Ledoit-Wolf shrinkage" if allocation["covariance_method"] == "ledoit_wolf"
                     else "sample covariance")
        detail = (f" (intensity {risk.shrinkage:.2f})"
                  if allocation["covariance_method"] == "ledoit_wolf"
                  and np.isfinite(risk.shrinkage) else "")
        lines.append(
            f"Risk settings: {estimator}{detail}; expected returns taken as the "
            f"{risk.return_basis}; return smoothing "
            + ("reversed for funds showing positive autocorrelation, which raises volatility and "
               "correlation" if allocation["unsmooth_returns"] else "left as reported, so stale "
               "monthly marks understate volatility and correlation")
            + ".")
    return lines


# -----------------------------------------------------------------------------
# Rendering
#
# One document, two renderers. Word is what the reader asked for; Markdown is
# what they get when python-docx is not installed, which on a locked-down
# machine is a real possibility and not a reason to lose the write-up.
# -----------------------------------------------------------------------------

def render_markdown(blocks: Sequence[dict[str, Any]], path: Path,
                    findings: RunFindings) -> None:
    lines: list[str] = []
    for item in blocks:
        kind = item["kind"]
        if kind == "title":
            lines += [f"# {item['text']}", ""]
        elif kind == "subtitle":
            lines += [f"*{item['text']}*", ""]
        elif kind == "heading":
            lines += ["#" * (int(item["level"]) + 1) + f" {item['text']}", ""]
        elif kind == "para":
            lines += [item["text"], ""]
        elif kind == "note":
            lines += [f"> {item['text']}", ""]
        elif kind == "bullets":
            lines += [f"* {text}" for text in item["items"]] + [""]
        elif kind == "table":
            lines += ["| " + " | ".join(item["headers"]) + " |",
                      "| " + " | ".join("---" for _ in item["headers"]) + " |"]
            lines += ["| " + " | ".join(str(cell) for cell in row) + " |"
                      for row in item["rows"]] + [""]
        elif kind == "image":
            picture = findings.charts.get(item["key"])
            if picture is not None:
                relative = os.path.relpath(picture, path.parent).replace(os.sep, "/")
                lines += [f"![{item['caption']}]({relative})", "",
                          f"*{item['caption']}*", ""]
    path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def render_docx(blocks: Sequence[dict[str, Any]], path: Path, findings: RunFindings) -> None:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor

    document = Document()
    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(8)

    for item in blocks:
        kind = item["kind"]
        if kind == "title":
            heading = document.add_heading(item["text"], level=0)
            heading.paragraph_format.space_after = Pt(2)
        elif kind == "subtitle":
            paragraph = document.add_paragraph(item["text"])
            run = paragraph.runs[0]
            run.italic = True
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        elif kind == "heading":
            document.add_heading(item["text"], level=int(item["level"]))
        elif kind == "para":
            document.add_paragraph(item["text"])
        elif kind == "note":
            paragraph = document.add_paragraph(item["text"])
            run = paragraph.runs[0]
            run.italic = True
            run.font.size = Pt(9)
            run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        elif kind == "bullets":
            for text in item["items"]:
                document.add_paragraph(text, style="List Bullet")
        elif kind == "table":
            table = document.add_table(rows=1, cols=len(item["headers"]))
            table.style = "Light Grid Accent 1"
            for cell, text in zip(table.rows[0].cells, item["headers"], strict=True):
                cell.text = str(text)
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.bold = True
            for row in item["rows"]:
                cells = table.add_row().cells
                for cell, text in zip(cells, row, strict=True):
                    cell.text = str(text)
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.size = Pt(9)
            document.add_paragraph()
        elif kind == "image":
            picture = findings.charts.get(item["key"])
            if picture is None:
                continue
            try:
                document.add_picture(str(picture), width=Inches(6.4))
            except Exception:  # a picture Word will not take must not lose the text
                LOGGER.exception("Could not embed %s", picture)
                continue
            document.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
            caption = document.add_paragraph(item["caption"])
            caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = caption.runs[0]
            run.italic = True
            run.font.size = Pt(8.5)
            run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    document.save(str(path))


def write_commentary(data: WorkbookData, settings: dict[str, Any], findings: RunFindings,
                     output_root: Path) -> list[str]:
    """Write the commentary, in Word where that is possible and Markdown where it is not."""
    options = settings["commentary"]
    blocks = build_commentary(data, settings, findings)
    wanted = str(options["file_format"])
    have_docx = importlib.util.find_spec("docx") is not None
    log: list[str] = []

    if wanted in {"auto", "docx"} and not have_docx:
        if wanted == "docx":
            raise WorkbookFormatError(
                "The commentary was asked for as a Word document, but python-docx is not "
                "installed. Run 'pip install python-docx', or set the commentary format to "
                "Markdown."
            )
        log.append(
            "python-docx is not installed, so the commentary was written as Markdown instead of "
            "Word. Run 'pip install python-docx' for a .docx file."
        )

    use_docx = have_docx and wanted in {"auto", "docx"}
    path = output_root / f"{COMMENTARY_STEM}.{'docx' if use_docx else 'md'}"
    try:
        if use_docx:
            render_docx(blocks, path, findings)
        else:
            render_markdown(blocks, path, findings)
    except PermissionError:
        log.append(
            f"The commentary could not be written to {path}. Close the file if it is open in "
            "another application."
        )
        return log

    embedded = sum(1 for item in blocks
                   if item["kind"] == "image" and item["key"] in findings.charts)
    log.append(f"Commentary: {path}")
    log.append(
        f"{sum(1 for item in blocks if item['kind'] in {'para', 'bullets'})} passage(s) and "
        f"{embedded} embedded chart(s)."
    )
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
    problems = check_modules_against_layout(
        data.layout, selected_modules(settings), has_weights=data.has_weights)
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

    # The commentary is written after every module has finished, from what the
    # modules recorded on the way past. Collecting is cheap and the collector
    # is inert when the document is switched off, so the modules are handed one
    # either way rather than being run down two different paths.
    commentary = settings["commentary"]
    findings = RunFindings(
        output_root / f"{COMMENTARY_STEM}_charts",
        dpi=int(commentary["chart_dpi"]),
        embed_charts=bool(commentary["create_document"]) and bool(commentary["embed_charts"]),
    )

    # Module 4 is handed the clustering settings as well as its own, so its
    # clusters are always the ones module 2 drew.
    runners: dict[str, Callable[[WorkbookData, dict[str, Any], Path, RunFindings], list[str]]] = {
        "cone": lambda book, chosen, folder, found: run_cone_module(
            book, chosen["cone"], folder, found),
        "clustering": lambda book, chosen, folder, found: run_clustering_module(
            book, chosen["clustering"], folder, found),
        "performance": lambda book, chosen, folder, found: run_performance_module(
            book, chosen["performance"], folder, found),
        "allocation": lambda book, chosen, folder, found: run_allocation_module(
            book, chosen["allocation"], chosen["clustering"], folder, found),
    }
    for key in selected_modules(settings):
        report.append(f"--- {MODULE_TITLES[key]} ---")
        LOGGER.info("Running %s", MODULE_TITLES[key])
        try:
            report.extend(f"  {line}" for line in
                          runners[key](data, settings, output_root / MODULE_FOLDERS[key],
                                       findings))
        except Exception as exc:  # keep the other modules going, and say what failed
            LOGGER.exception("%s failed", MODULE_TITLES[key])
            report.append(f"  FAILED: {exc}")
        report.append("")

    if commentary["create_document"]:
        report.append("--- Written commentary ---")
        LOGGER.info("Writing the commentary")
        try:
            report.extend(f"  {line}" for line in
                          write_commentary(data, settings, findings, output_root))
        except Exception as exc:  # the charts are the deliverable; the write-up is not
            LOGGER.exception("The commentary failed")
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
    allocation_tab = ttk.Frame(notebook, padding=12)
    commentary_tab = ttk.Frame(notebook, padding=12)
    notebook.add(run_tab, text="Workbook & modules")
    notebook.add(cone_tab, text=MODULE_TITLES["cone"])
    notebook.add(cluster_tab, text=MODULE_TITLES["clustering"])
    notebook.add(performance_tab, text=MODULE_TITLES["performance"])
    notebook.add(allocation_tab, text=MODULE_TITLES["allocation"])
    notebook.add(commentary_tab, text="Written commentary")

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
    ttk.Checkbutton(modules_box, text=MODULE_TITLES["allocation"] +
                    '  (either layout, plus a "Weight" or "Current Position" row)',
                    variable=module_vars["allocation"]).grid(
        row=3, column=0, sticky="w", pady=2)

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
    ttk.Button(check_box, text="Check the workbook again",
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

    # ---- Module 4 tab ------------------------------------------------------
    allocation_tab.columnconfigure(1, weight=1)
    ttk.Label(allocation_tab,
              text=("Allocates the portfolio's volatility across the clusters module 2 finds, so "
                    "the clusters here are the ones on that dendrogram. It uses module 2's "
                    "lookbacks, clustering method and cluster count. Needs a weights row: put "
                    '"Weight" - or "Current Position", "Allocation", "Holding", "% of '
                    'Portfolio" - in column A directly above the first month, and each fund\'s '
                    "share of the portfolio in its own column. Leave a fund blank to measure it "
                    "as a candidate the portfolio does not yet hold."),
              foreground="#555555", wraplength=860).grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
    labelled_combo(allocation_tab, 1, "Covariance estimator",
                   string_var("allocation", "covariance_method"), ("ledoit_wolf", "sample"),
                   "shrinkage is steadier on short windows")
    labelled_combo(allocation_tab, 2, "Expected returns from",
                   string_var("allocation", "return_basis"), ("auto", "expected", "realised"),
                   "auto = row 3 where every fund has one")
    labelled_entry(allocation_tab, 3, "Flag a cluster at",
                   string_var("allocation", "risk_multiplier_flag"),
                   "x its share of capital, or more", width=10)
    labelled_entry(allocation_tab, 4, "Fund-label wrap width",
                   string_var("allocation", "label_wrap_width"), "characters", width=10)
    labelled_entry(allocation_tab, 5, "Image resolution", string_var("allocation", "dpi"),
                   "DPI", width=10)
    ttk.Checkbutton(allocation_tab,
                    text=("Reverse return smoothing before measuring risk (raises the volatility "
                          "of funds whose marks lag, such as credit)"),
                    variable=boolean_var("allocation", "unsmooth_returns")).grid(
        row=6, column=0, columnspan=3, sticky="w", pady=3)
    allocation_outputs = ttk.LabelFrame(allocation_tab, text="Files to create", padding=8)
    allocation_outputs.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(8, 0))
    ttk.Checkbutton(allocation_outputs, text="PDF report per lookback",
                    variable=boolean_var("allocation", "create_pdf")).grid(row=0, column=0, sticky="w")
    ttk.Checkbutton(allocation_outputs, text="PNG charts",
                    variable=boolean_var("allocation", "save_png")).grid(
        row=0, column=1, sticky="w", padx=(20, 0))
    ttk.Checkbutton(allocation_outputs, text="CSV contributions, budget and summary",
                    variable=boolean_var("allocation", "save_csv")).grid(
        row=0, column=2, sticky="w", padx=(20, 0))

    # ---- Written commentary tab -------------------------------------------
    commentary_tab.columnconfigure(1, weight=1)
    ttk.Label(commentary_tab,
              text=("Writes a short document beside the charts explaining what each output is, "
                    "how to read it, and what this workbook actually shows - where the risk sits, "
                    "where funds could be added or trimmed, and anything else notable, with the "
                    "funds named. It describes whichever modules ran; module 4 supplies the risk "
                    "and add-or-trim reading, so without a weights row the document covers "
                    "behaviour but not the shape of the book."),
              foreground="#555555", wraplength=860).grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))
    ttk.Checkbutton(commentary_tab, text="Write the commentary document",
                    variable=boolean_var("commentary", "create_document")).grid(
        row=1, column=0, columnspan=3, sticky="w", pady=3)
    labelled_combo(commentary_tab, 2, "File format",
                   string_var("commentary", "file_format"), ("auto", "docx", "markdown"),
                   "auto = Word when python-docx is installed")
    ttk.Checkbutton(commentary_tab,
                    text="Put the headline charts in the document",
                    variable=boolean_var("commentary", "embed_charts")).grid(
        row=3, column=0, columnspan=3, sticky="w", pady=3)
    labelled_entry(commentary_tab, 4, "Embedded chart resolution",
                   string_var("commentary", "chart_dpi"), "DPI", width=10)
    labelled_entry(commentary_tab, 5, "Name at most",
                   string_var("commentary", "max_named_funds"),
                   "funds in any one list, before it says \"and N others\"", width=10)
    labelled_entry(commentary_tab, 6, "Only mention gaps of",
                   string_var("commentary", "material_gap_pct"),
                   "percentage points or more, so that noise is not reported as a finding",
                   width=10)

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
        try:
            material_gap = float(variables["commentary.material_gap_pct"].get().strip())
        except ValueError as exc:
            raise ValueError(
                "The commentary's materiality threshold must be a number, such as 3.") from exc
        try:
            crowding_flag = float(variables["allocation.risk_multiplier_flag"].get().strip())
        except ValueError as exc:
            raise ValueError(
                "The portfolio-risk crowding flag must be a number, such as 1.25."
            ) from exc
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
            "commentary": {
                "create_document": bool(variables["commentary.create_document"].get()),
                "file_format": variables["commentary.file_format"].get().strip(),
                "embed_charts": bool(variables["commentary.embed_charts"].get()),
                "chart_dpi": whole_number("commentary", "chart_dpi",
                                          "The embedded chart resolution"),
                "max_named_funds": whole_number("commentary", "max_named_funds",
                                                "The number of funds to name"),
                "material_gap_pct": material_gap,
            },
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
            "allocation": {
                "covariance_method": variables["allocation.covariance_method"].get(),
                "return_basis": variables["allocation.return_basis"].get(),
                "unsmooth_returns": bool(variables["allocation.unsmooth_returns"].get()),
                "risk_multiplier_flag": crowding_flag,
                "label_wrap_width": whole_number("allocation", "label_wrap_width",
                                                 "The portfolio-risk label wrap width"),
                "create_pdf": bool(variables["allocation.create_pdf"].get()),
                "save_png": bool(variables["allocation.save_png"].get()),
                "save_csv": bool(variables["allocation.save_csv"].get()),
                "dpi": whole_number("allocation", "dpi", "The portfolio-risk image resolution"),
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

    # ---- Checking automatically -------------------------------------------
    # The check re-runs on its own whenever something that could change its
    # result is altered: the workbook, the tab, the modules or the two data
    # checks. It is debounced, so typing a path does not start a check per
    # keystroke, and it never raises a dialog - the panel is the only place an
    # automatic result appears. The buttons stay, for re-checking on demand
    # after the workbook has been edited in Excel.
    pending_check: list[str] = []
    check_running: list[bool] = []

    def schedule_auto_check(*_args: Any) -> None:
        for job in pending_check:
            try:
                root.after_cancel(job)
            except tk.TclError:
                pass
        pending_check.clear()
        pending_check.append(root.after(400, run_auto_check))

    def run_auto_check() -> None:
        pending_check.clear()
        if check_running:
            return
        path_text = excel_var.get().strip().strip('"')
        if not path_text or not Path(path_text).expanduser().is_file():
            write_check_output(
                "Choose a workbook, and its format check will appear here automatically."
            )
            return
        check_running.append(True)
        try:
            write_check_output("Checking the workbook...")
            check_output.update_idletasks()
            perform_check(quiet=True)
        finally:
            check_running.clear()

    excel_var.trace_add("write", schedule_auto_check)
    sheet_var.trace_add("write", schedule_auto_check)
    for auto_name in ("workbook.stop_on_missing_months", "workbook.stop_on_history_gaps"):
        variables[auto_name].trace_add("write", schedule_auto_check)
    for module_variable in module_vars.values():
        module_variable.trace_add("write", schedule_auto_check)

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
    schedule_auto_check()
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
    parser.add_argument("--commentary", action=argparse.BooleanOptionalAction, default=None,
                        help="Write the commentary document, or skip it")
    parser.add_argument("--commentary-format", choices=("auto", "docx", "markdown"),
                        help="Word where python-docx is installed (auto), or force one")
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
    if arguments.commentary is not None:
        settings["commentary"]["create_document"] = bool(arguments.commentary)
    if arguments.commentary_format is not None:
        settings["commentary"]["file_format"] = arguments.commentary_format
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
