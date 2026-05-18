# ============================================================
# real_sector.py
# Somalia Economic Outlook Report
# Real Sector Pipeline
# ============================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
from matplotlib.patches import Patch
from matplotlib.lines import Line2D


# ============================================================
# 0. PROJECT PATHS
# ============================================================

# ============================================================
# PATHS / DIRECTORIES
# ============================================================

from pathlib import Path

# Root project folder
PROJECT_DIR = Path(__file__).parent

# Main Excel data file
DATA_FILE = PROJECT_DIR / "Data template.xlsx"

# Sheet names
REAL_SHEET = "Real Sector Raw"
FISCAL_SHEET = "Fiscal Sector Raw"
MONETARY_SHEET = "Monetary Financial Raw"
EXTERNAL_SHEET = "External Sector Raw"

# Output folders
OUTPUT_DIR = PROJECT_DIR / "outputs"
CHART_DIR = OUTPUT_DIR / "charts"
TABLE_DIR = OUTPUT_DIR / "tables"

# Create folders automatically
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
CHART_DIR.mkdir(parents=True, exist_ok=True)
TABLE_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# 1. LOAD TEMPLATE
# ============================================================

def load_sector_sheet(file_path, sheet_name):
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    header_row = None
    for i in range(len(raw)):
        row_values = raw.iloc[i].astype(str).str.strip().tolist()
        if "Sector" in row_values and "Variable" in row_values:
            header_row = i
            break

    if header_row is None:
        raise ValueError(f"Could not find header row in sheet: {sheet_name}")

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    if "Source / Notes" not in df.columns:
        possible_source_cols = [c for c in df.columns if "Source" in c or "Notes" in c]
        if possible_source_cols:
            df = df.rename(columns={possible_source_cols[0]: "Source / Notes"})

    id_cols = ["Sector", "Variable", "Unit", "Type", "Source / Notes"]
    year_cols = [c for c in df.columns if str(c).strip().isdigit()]

    missing = [c for c in id_cols if c not in df.columns]
    if missing:
        raise KeyError(f"Missing columns in {sheet_name}: {missing}. Found: {list(df.columns)}")

    df = df[id_cols + year_cols].copy()
    df["Variable"] = df["Variable"].astype(str).str.strip()

    for c in year_cols:
        df[c] = (
            df[c]
            .astype(str)
            .str.replace(",", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.strip()
            .replace({"": np.nan, "nan": np.nan, "-": np.nan})
        )
        df[c] = pd.to_numeric(df[c], errors="coerce")

    years = [int(c) for c in year_cols]
    return df, years


# ============================================================
# 2. HELPERS
# ============================================================

def get_series(df, years, variable):
    row = df.loc[df["Variable"].eq(variable)]

    if row.empty:
        print(f"WARNING: Missing variable: {variable}")
        return pd.Series(index=years, dtype=float)

    s = row[[str(y) for y in years]].iloc[0]
    s.index = years

    return pd.to_numeric(s, errors="coerce")


def yoy(x):
    return x.pct_change() * 100


def pct_gdp(x, gdp):
    return x / gdp * 100


def has_data(x, min_obs=2):
    return x.notna().sum() >= min_obs



def select_years(years, start_year=None, end_year=None):
    years = sorted([int(y) for y in years])
    if start_year is not None:
        years = [y for y in years if y >= int(start_year)]
    if end_year is not None:
        years = [y for y in years if y <= int(end_year)]
    return years


def latest_two_years(years):
    years = sorted([int(y) for y in years])
    if len(years) < 2:
        raise ValueError("At least two years are needed for change tables.")
    return years[-2], years[-1]


def style_ax(ax, title):
    ax.set_title(title, loc="left", fontsize=10, fontstyle="italic", fontweight="bold")
    ax.axhline(0, linewidth=0.8, color="black")
    ax.grid(True, alpha=0.22)
    ax.tick_params(axis="x", labelsize=8)
    ax.tick_params(axis="y", labelsize=8)
    ax.spines["top"].set_visible(False)


def add_bar_labels(ax, bars, fmt="{:.1f}", fontsize=7, threshold=0.05):
    for bar in bars:
        h = bar.get_height()

        if np.isnan(h) or abs(h) < threshold:
            continue

        x = bar.get_x() + bar.get_width() / 2
        y = bar.get_y() + h + 0.15 if h >= 0 else bar.get_y() + h - 0.15
        va = "bottom" if h >= 0 else "top"

        ax.text(x, y, fmt.format(h), ha="center", va=va, fontsize=fontsize)


# ============================================================
# 3. COMPUTE INDICATORS
# ============================================================

def compute_real_sector_indicators(real_raw, years):
    nominal_gdp = get_series(real_raw, years, "Nominal GDP")
    real_gdp = get_series(real_raw, years, "Real GDP")

    hh_cons = get_series(real_raw, years, "Household final consumption")
    gov_cons = get_series(real_raw, years, "Government final consumption")
    investment = get_series(real_raw, years, "Gross fixed capital formation")
    exports_gs = get_series(real_raw, years, "Exports of goods and services")
    imports_gs = get_series(real_raw, years, "Imports of goods and services")

    cpi_headline = get_series(real_raw, years, "CPI headline index")
    food_cpi = get_series(real_raw, years, "Food CPI index/Non Alcoholic Beverages")
    housing_cpi = get_series(real_raw, years, "Housing/utilities/Fuel CPI index")
    transport_cpi = get_series(real_raw, years, "Transport CPI index")

    real_growth = yoy(real_gdp)

    valid_gdp = real_gdp.dropna()

    if len(valid_gdp) >= 3:
        t_valid = np.arange(len(valid_gdp))
        coef = np.polyfit(t_valid, np.log(valid_gdp.values), 1)

        potential_gdp = pd.Series(
            np.exp(coef[1] + coef[0] * np.arange(len(years))),
            index=years
        )

        output_gap = (real_gdp / potential_gdp - 1) * 100
    else:
        potential_gdp = pd.Series(index=years, dtype=float)
        output_gap = pd.Series(index=years, dtype=float)

    prev_gdp = real_gdp.shift(1)

    contrib_private_consumption = hh_cons.diff() / prev_gdp * 100
    contrib_government_consumption = gov_cons.diff() / prev_gdp * 100
    contrib_investment = investment.diff() / prev_gdp * 100
    contrib_exports = exports_gs.diff() / prev_gdp * 100
    contrib_imports = -imports_gs.diff() / prev_gdp * 100

    contrib_sum_before_residual = (
        contrib_private_consumption
        + contrib_government_consumption
        + contrib_investment
        + contrib_exports
        + contrib_imports
    )

    contrib_residual = real_growth - contrib_sum_before_residual

    decomposition_check = (
        contrib_private_consumption
        + contrib_government_consumption
        + contrib_investment
        + contrib_exports
        + contrib_imports
        + contrib_residual
    )

    decomposition_error = real_growth - decomposition_check

    investment_ratio = pct_gdp(investment, nominal_gdp)
    exports_ratio = pct_gdp(exports_gs, nominal_gdp)
    imports_ratio = pct_gdp(imports_gs, nominal_gdp)
    trade_balance_ratio = pct_gdp(exports_gs - imports_gs, nominal_gdp)

    return {
        "nominal_gdp": nominal_gdp,
        "real_gdp": real_gdp,
        "real_growth": real_growth,
        "potential_gdp": potential_gdp,
        "output_gap": output_gap,

        "hh_cons": hh_cons,
        "gov_cons": gov_cons,
        "investment": investment,
        "exports_gs": exports_gs,
        "imports_gs": imports_gs,

        "contrib_private_consumption": contrib_private_consumption,
        "contrib_government_consumption": contrib_government_consumption,
        "contrib_investment": contrib_investment,
        "contrib_exports": contrib_exports,
        "contrib_imports": contrib_imports,
        "contrib_residual": contrib_residual,
        "decomposition_check": decomposition_check,
        "decomposition_error": decomposition_error,

        "investment_ratio": investment_ratio,
        "exports_ratio": exports_ratio,
        "imports_ratio": imports_ratio,
        "trade_balance_ratio": trade_balance_ratio,

        "cpi_headline": cpi_headline,
        "food_cpi": food_cpi,
        "housing_cpi": housing_cpi,
        "transport_cpi": transport_cpi,
    }


# ============================================================
# 4. SUMMARY TABLE
# ============================================================

def create_summary_table(indicators, years):
    summary = pd.DataFrame(index=years)

    summary["Nominal GDP"] = indicators["nominal_gdp"]
    summary["Real GDP"] = indicators["real_gdp"]
    summary["Real GDP growth"] = indicators["real_growth"]
    summary["Output gap"] = indicators["output_gap"]

    summary["Investment (% GDP)"] = indicators["investment_ratio"]
    summary["Exports (% GDP)"] = indicators["exports_ratio"]
    summary["Imports (% GDP)"] = indicators["imports_ratio"]
    summary["Trade balance (% GDP)"] = indicators["trade_balance_ratio"]

    summary["Headline inflation"] = yoy(indicators["cpi_headline"])
    summary["Food inflation"] = yoy(indicators["food_cpi"])
    summary["Housing/fuel inflation"] = yoy(indicators["housing_cpi"])
    summary["Transport inflation"] = yoy(indicators["transport_cpi"])

    summary["Private consumption contribution"] = indicators["contrib_private_consumption"]
    summary["Government consumption contribution"] = indicators["contrib_government_consumption"]
    summary["Investment contribution"] = indicators["contrib_investment"]
    summary["Export contribution"] = indicators["contrib_exports"]
    summary["Import contribution"] = indicators["contrib_imports"]
    summary["Residual contribution"] = indicators["contrib_residual"]
    summary["Decomposition check"] = indicators["decomposition_check"]
    summary["Decomposition error"] = indicators["decomposition_error"]

    return summary.round(2)


# ============================================================
# 5. PLOT DASHBOARD
# ============================================================

def plot_real_sector_dashboard(indicators, years, save_path=None, start_year=None, end_year=None):
    years_main = select_years(years, start_year, end_year)
    years_cpi = years_main

    fig, axes = plt.subplots(2, 2, figsize=(15, 9))

    fig.suptitle(
        "Somalia: Real Sector Diagnostics",
        fontsize=18,
        fontweight="bold",
        color="#1f77b4"
    )

    # --------------------------------------------------------
    # 1. Growth and output gap
    # --------------------------------------------------------

    ax = axes[0, 0]

    bars = ax.bar(
        years_main,
        indicators["output_gap"].loc[years_main],
        label="Output gap",
        color="#1f77b4",
        alpha=0.85
    )

    ax.plot(
        years_main,
        indicators["real_growth"].loc[years_main],
        color="#ff7f0e",
        marker="o",
        linewidth=2.8,
        linestyle="--",
        label="Real GDP growth"
    )

    add_bar_labels(ax, bars, threshold=0.2)

    for x, y in zip(years_main, indicators["real_growth"].loc[years_main]):
        if pd.notna(y):
            ax.text(
                x, y + 0.25, f"{y:.1f}",
                ha="center",
                fontsize=8,
                color="#ff7f0e",
                fontweight="bold"
            )

    style_ax(ax, "Real activity recovered as the output gap narrowed after recent shocks.")
    ax.set_ylabel("Percent")
    ax.legend(fontsize=8, frameon=True)

    # --------------------------------------------------------
    # 2. Expenditure-side decomposition
    # --------------------------------------------------------

    ax = axes[0, 1]

    colors = {
        "Private consumption": "#17becf",
        "Government consumption": "#2ca02c",
        "Investment": "#9467bd",
        "Exports": "#e377c2",
        "Imports drag": "#ff7f0e",
        "Residual/statistical discrepancy": "#1f77b4",
    }

    contribution_items = [
        ("contrib_private_consumption", "Private consumption"),
        ("contrib_government_consumption", "Government consumption"),
        ("contrib_investment", "Investment"),
        ("contrib_exports", "Exports"),
        ("contrib_imports", "Imports drag"),
        ("contrib_residual", "Residual/statistical discrepancy"),
    ]

    positive_bottom = np.zeros(len(years_main))
    negative_bottom = np.zeros(len(years_main))
    all_bars = []

    for key, label in contribution_items:
        vals = indicators[key].loc[years_main].fillna(0).values

        pos_vals = np.where(vals > 0, vals, 0)
        neg_vals = np.where(vals < 0, vals, 0)

        b_pos = ax.bar(
            years_main,
            pos_vals,
            bottom=positive_bottom,
            color=colors[label]
        )

        b_neg = ax.bar(
            years_main,
            neg_vals,
            bottom=negative_bottom,
            color=colors[label]
        )

        all_bars.extend(b_pos)
        all_bars.extend(b_neg)

        positive_bottom += pos_vals
        negative_bottom += neg_vals

    ax.plot(
        years_main,
        indicators["real_growth"].loc[years_main].fillna(0),
        color="black",
        marker="o",
        linewidth=2.8,
        linestyle="--"
    )

    for x, y in zip(years_main, indicators["real_growth"].loc[years_main]):
        if pd.notna(y):
            ax.text(
                x, y + 0.25, f"{y:.1f}",
                ha="center",
                fontsize=8,
                color="black",
                fontweight="bold"
            )

    for bar in all_bars:
        h = bar.get_height()

        if abs(h) >= 2.5:
            x = bar.get_x() + bar.get_width() / 2
            y = bar.get_y() + h / 2

            ax.text(
                x,
                y,
                f"{h:.1f}",
                ha="center",
                va="center",
                fontsize=7,
                color="white",
                fontweight="bold"
            )

    legend_items = [
        Line2D([0], [0], color="black", linestyle="--", marker="o", label="Real GDP growth"),
        Patch(facecolor=colors["Private consumption"], label="Private consumption"),
        Patch(facecolor=colors["Government consumption"], label="Government consumption"),
        Patch(facecolor=colors["Investment"], label="Investment"),
        Patch(facecolor=colors["Exports"], label="Exports"),
        Patch(facecolor=colors["Imports drag"], label="Imports drag"),
        Patch(facecolor=colors["Residual/statistical discrepancy"], label="Residual/statistical discrepancy"),
    ]

    style_ax(ax, "Growth was increasingly driven by domestic demand amid rising import leakage.")
    ax.set_ylabel("Percentage points")
    ax.legend(handles=legend_items, fontsize=7, ncol=2, frameon=True)

    # --------------------------------------------------------
    # 3. Inflation components
    # --------------------------------------------------------

    ax = axes[1, 0]

    inflation_items = [
        (yoy(indicators["cpi_headline"]).loc[years_cpi], "Headline CPI"),
        (yoy(indicators["food_cpi"]).loc[years_cpi], "Food CPI"),
        (yoy(indicators["housing_cpi"]).loc[years_cpi], "Housing/fuel CPI"),
        (yoy(indicators["transport_cpi"]).loc[years_cpi], "Transport CPI"),
    ]

    for s, label in inflation_items:
        if has_data(s):
            ax.plot(years_cpi, s, marker="o", linewidth=2.2, label=label)

    style_ax(ax, "Inflation pressures reflected food, fuel, and transport price volatility.")
    ax.set_ylabel("Percent, year-on-year")
    ax.legend(fontsize=8, frameon=True)

    # --------------------------------------------------------
    # 4. Investment and trade dependence
    # --------------------------------------------------------

    ax = axes[1, 1]

    trade_balance = indicators["trade_balance_ratio"].loc[years_main]
    investment_ratio = indicators["investment_ratio"].loc[years_main]
    exports_ratio = indicators["exports_ratio"].loc[years_main]
    imports_ratio = indicators["imports_ratio"].loc[years_main]

    bars = ax.bar(
        years_main,
        trade_balance,
        color="#1f77b4",
        alpha=0.95,
        label="Trade balance"
    )

    ax2 = ax.twinx()

    ax2.plot(
        years_main,
        investment_ratio,
        color="#9467bd",
        linewidth=2.5,
        label="Investment"
    )

    ax2.plot(
        years_main,
        exports_ratio,
        color="#2ca02c",
        linewidth=2.5,
        linestyle="--",
        label="Exports"
    )

    ax2.plot(
        years_main,
        imports_ratio,
        color="#d62728",
        linewidth=2.5,
        linestyle="--",
        label="Imports"
    )

    ax2.set_ylim(0, 100)

    add_bar_labels(ax, bars, threshold=1.0)

    style_ax(ax, "High import dependence continued to widen Somalia’s external trade deficit.")
    ax.set_ylabel("Trade balance, percent of GDP")
    ax2.set_ylabel("Investment, exports and imports, percent of GDP")
    ax2.spines["top"].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()

    ax.legend(lines1 + lines2, labels1 + labels2, fontsize=8, frameon=True)

    fig.text(
        0.01,
        0.01,
        "Source: Somalia macro template; CBS/SNBS; staff calculations.",
        fontsize=8
    )

    plt.tight_layout(rect=[0, 0.04, 1, 0.94])

    if save_path is not None:
        fig.savefig(save_path, dpi=300, bbox_inches="tight")

    return fig


# ============================================================
# 6. MAIN
# ============================================================

def main():
    print("Loading real sector data...")
    real_raw, years = load_sector_sheet(DATA_FILE, REAL_SHEET)

    print("Computing real sector indicators...")
    indicators = compute_real_sector_indicators(real_raw, years)

    print("Creating summary table...")
    summary = create_summary_table(indicators, years)

    summary_output = TABLE_DIR / "real_sector_summary.xlsx"
    chart_output = CHART_DIR / "real_sector_dashboard.png"

    print("Exporting summary table...")
    summary.to_excel(summary_output)

    print("Generating dashboard...")
    plot_real_sector_dashboard(indicators, years, save_path=chart_output)

    print("Done.")
    print(f"Summary table saved to: {summary_output}")
    print(f"Dashboard saved to: {chart_output}")


if __name__ == "__main__":
    main()