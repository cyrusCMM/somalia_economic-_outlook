# ============================================================
# fiscal_sector.py
# Somalia Economic Outlook Report
# Fiscal Sector Pipeline
# 3x3 PPT-ready dashboard
# ============================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path


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
# LOAD TEMPLATE
# ============================================================

def load_sector_sheet(file_path, sheet_name):
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    header_row = None
    for i in range(len(raw)):
        values = raw.iloc[i].astype(str).str.strip().tolist()
        if "Sector" in values and "Variable" in values:
            header_row = i
            break

    if header_row is None:
        raise ValueError(f"Could not find header row in sheet: {sheet_name}")

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")

    if "Source / Notes" not in df.columns:
        possible = [c for c in df.columns if "Source" in c or "Notes" in c]
        if possible:
            df = df.rename(columns={possible[0]: "Source / Notes"})

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
# HELPERS
# ============================================================

def get_series(df, years, variable):
    row = df.loc[df["Variable"].eq(variable)]

    if row.empty:
        print(f"WARNING: Missing variable: {variable}")
        return pd.Series(index=years, dtype=float)

    s = row[[str(y) for y in years]].iloc[0]
    s.index = years
    return pd.to_numeric(s, errors="coerce")


def pct_gdp(x, gdp):
    return x / gdp.replace(0, np.nan) * 100


def share(x, total):
    return x / total.replace(0, np.nan) * 100



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
    ax.set_title(title, loc="left", fontsize=10.5, fontweight="bold")
    ax.axhline(0, linewidth=0.8, color="black")
    ax.grid(True, alpha=0.15)
    ax.tick_params(axis="x", labelsize=8.5)
    ax.tick_params(axis="y", labelsize=8.5)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def label_bars(ax, bars, threshold=0.2, fmt="{:.1f}", fontsize=7.5):
    for bar in bars:
        h = bar.get_height()
        if pd.isna(h) or abs(h) < threshold:
            continue

        x = bar.get_x() + bar.get_width() / 2
        y = bar.get_y() + h + 0.10 if h >= 0 else bar.get_y() + h - 0.10
        va = "bottom" if h >= 0 else "top"

        ax.text(x, y, fmt.format(h), ha="center", va=va, fontsize=fontsize)


def label_stacked(ax, bars, threshold=8, fontsize=7.5):
    for bar in bars:
        h = bar.get_height()
        if pd.isna(h) or abs(h) < threshold:
            continue

        x = bar.get_x() + bar.get_width() / 2
        y = bar.get_y() + h / 2

        ax.text(
            x,
            y,
            f"{h:.0f}",
            ha="center",
            va="center",
            fontsize=fontsize,
            color="white",
            fontweight="bold",
        )


def format_table(table_df):
    out = table_df.copy()
    out = out.replace([np.inf, -np.inf], np.nan)
    out = out.round(1)
    return out.astype(object).where(pd.notna(out), "—")


def add_table(ax, table_df, title):
    ax.axis("off")
    ax.set_title(title, loc="left", fontsize=10.5, fontweight="bold")

    display_df = format_table(table_df)

    tbl = ax.table(
        cellText=display_df.values,
        colLabels=display_df.columns,
        rowLabels=display_df.index,
        loc="center",
        cellLoc="center",
        colLoc="center",
    )

    tbl.auto_set_font_size(False)
    tbl.set_fontsize(8.2)
    tbl.scale(1.08, 1.55)

    for (r, c), cell in tbl.get_celld().items():
        cell.set_linewidth(0.35)
        if r == 0:
            cell.set_text_props(weight="bold")

    return tbl


# ============================================================
# OUTPUT GAP
# ============================================================

def compute_output_gap(real_raw, years):
    real_gdp = get_series(real_raw, years, "Real GDP")
    valid = real_gdp.dropna()

    if len(valid) < 3:
        return pd.Series(index=years, dtype=float)

    t = np.arange(len(valid))
    coef = np.polyfit(t, np.log(valid.values), 1)

    potential = pd.Series(
        np.exp(coef[1] + coef[0] * np.arange(len(years))),
        index=years,
    )

    return (real_gdp / potential - 1) * 100


# ============================================================
# COMPUTE FISCAL INDICATORS
# ============================================================

def compute_fiscal_indicators(fiscal_raw, real_raw, years):
    nominal_gdp = get_series(real_raw, years, "Nominal GDP")
    output_gap = compute_output_gap(real_raw, years)

    revenue_grants = get_series(fiscal_raw, years, "Total revenue and grants")
    domestic_revenue = get_series(fiscal_raw, years, "Domestic revenue")
    tax_revenue = get_series(fiscal_raw, years, "Tax revenue")
    non_tax_revenue = get_series(fiscal_raw, years, "Non-tax revenue")

    customs = get_series(fiscal_raw, years, "Customs duties / trade taxes")
    sales_tax = get_series(fiscal_raw, years, "Sales tax / VAT")
    income_tax = get_series(fiscal_raw, years, "Income tax")
    corporate_income_tax = get_series(fiscal_raw, years, "Corporate income tax")
    excise_taxes = get_series(fiscal_raw, years, "Excise taxes")

    external_grants = get_series(fiscal_raw, years, "External grants")
    bilateral_grants = get_series(fiscal_raw, years, "Bilateral grants")
    multilateral_grants = get_series(fiscal_raw, years, "Multilateral grants")

    total_expenditure = get_series(fiscal_raw, years, "Total expenditure")
    recurrent_expenditure = get_series(fiscal_raw, years, "Recurrent expenditure")
    compensation = get_series(fiscal_raw, years, "Compensation of employees")
    goods_services = get_series(fiscal_raw, years, "Use of goods and services")
    social_benefits = get_series(fiscal_raw, years, "Social benefits")
    interest = get_series(fiscal_raw, years, "Interest and other charges")
    subsidies = get_series(fiscal_raw, years, "Subsidies")
    transfers_fms = get_series(fiscal_raw, years, "Transfers/grants to FMS")
    capital_expenditure = get_series(fiscal_raw, years, "Capital expenditure")

    overall_balance = get_series(fiscal_raw, years, "Overall balance / financing gap")
    calculated_balance = revenue_grants - total_expenditure
    overall_balance = overall_balance.fillna(calculated_balance)

    external_debt = get_series(fiscal_raw, years, "External debt stock")
    debt_service_paid = get_series(fiscal_raw, years, "Debt service paid")
    debt_service_due = get_series(fiscal_raw, years, "Debt service due")

    primary_balance = overall_balance + interest
    overall_balance_gdp = pct_gdp(overall_balance, nominal_gdp)
    primary_balance_gdp = pct_gdp(primary_balance, nominal_gdp)
    fiscal_impulse = -primary_balance_gdp.diff()

    # 100% composition of total resources.
    resource_components = pd.DataFrame({
        "Tax revenue": tax_revenue,
        "Non-tax revenue": non_tax_revenue,
        "External grants": external_grants,
    })

    resource_base = resource_components.sum(axis=1).replace(0, np.nan)
    resource_composition_share = resource_components.div(resource_base, axis=0) * 100

    # 100% composition of plotted tax heads.
    tax_heads = pd.DataFrame({
        "Customs/trade taxes": customs,
        "Sales tax/VAT": sales_tax,
        "Income tax": income_tax,
        "Corporate income tax": corporate_income_tax,
        "Excise taxes": excise_taxes,
    })

    tax_heads_sum = tax_heads.sum(axis=1).replace(0, np.nan)
    tax_composition_share = tax_heads.div(tax_heads_sum, axis=0) * 100

    # 100% expenditure composition.
    expenditure_components = pd.DataFrame({
        "Compensation": compensation,
        "Goods/services": goods_services,
        "Social benefits": social_benefits,
        "Interest": interest,
        "Subsidies": subsidies,
        "Transfers to FMS": transfers_fms,
        "Capital expenditure": capital_expenditure,
    })

    expenditure_base = expenditure_components.sum(axis=1).replace(0, np.nan)
    expenditure_composition_share = expenditure_components.div(expenditure_base, axis=0) * 100

    # 100% composition of plotted grant components.
    grants_components = pd.DataFrame({
        "Bilateral grants": bilateral_grants,
        "Multilateral grants": multilateral_grants,
    })

    grants_base = grants_components.sum(axis=1).replace(0, np.nan)
    grants_composition_share = grants_components.div(grants_base, axis=0) * 100

    # Consistency checks.
    revenue_check = revenue_grants - (domestic_revenue + external_grants)
    domestic_revenue_check = domestic_revenue - (tax_revenue + non_tax_revenue)
    balance_check = overall_balance - calculated_balance
    grants_check = external_grants - grants_components.sum(axis=1)

    return {
        "nominal_gdp": nominal_gdp,
        "output_gap": output_gap,

        "revenue_grants": revenue_grants,
        "domestic_revenue": domestic_revenue,
        "tax_revenue": tax_revenue,
        "non_tax_revenue": non_tax_revenue,
        "external_grants": external_grants,

        "customs": customs,
        "sales_tax": sales_tax,
        "income_tax": income_tax,
        "corporate_income_tax": corporate_income_tax,
        "excise_taxes": excise_taxes,

        "bilateral_grants": bilateral_grants,
        "multilateral_grants": multilateral_grants,

        "total_expenditure": total_expenditure,
        "recurrent_expenditure": recurrent_expenditure,
        "compensation": compensation,
        "goods_services": goods_services,
        "social_benefits": social_benefits,
        "interest": interest,
        "subsidies": subsidies,
        "transfers_fms": transfers_fms,
        "capital_expenditure": capital_expenditure,

        "overall_balance": overall_balance,
        "primary_balance": primary_balance,
        "overall_balance_gdp": overall_balance_gdp,
        "primary_balance_gdp": primary_balance_gdp,
        "fiscal_impulse": fiscal_impulse,

        "external_debt": external_debt,
        "external_debt_gdp": pct_gdp(external_debt, nominal_gdp),
        "debt_service_paid": debt_service_paid,
        "debt_service_due": debt_service_due,
        "debt_service_paid_revenue": share(debt_service_paid, domestic_revenue),

        "revenue_grants_gdp": pct_gdp(revenue_grants, nominal_gdp),
        "domestic_revenue_gdp": pct_gdp(domestic_revenue, nominal_gdp),
        "tax_revenue_gdp": pct_gdp(tax_revenue, nominal_gdp),
        "customs_gdp": pct_gdp(customs, nominal_gdp),
        "grants_gdp": pct_gdp(external_grants, nominal_gdp),
        "expenditure_gdp": pct_gdp(total_expenditure, nominal_gdp),
        "recurrent_gdp": pct_gdp(recurrent_expenditure, nominal_gdp),
        "capital_gdp": pct_gdp(capital_expenditure, nominal_gdp),

        "resource_composition_share": resource_composition_share,
        "tax_composition_share": tax_composition_share,
        "expenditure_composition_share": expenditure_composition_share,
        "grants_composition_share": grants_composition_share,

        "revenue_check": revenue_check,
        "domestic_revenue_check": domestic_revenue_check,
        "balance_check": balance_check,
        "grants_check": grants_check,
    }


# ============================================================
# TABLES
# ============================================================

def create_summary_table(ind, years):
    summary = pd.DataFrame(index=years)

    summary["Revenue and grants"] = ind["revenue_grants"]
    summary["Domestic revenue"] = ind["domestic_revenue"]
    summary["Tax revenue"] = ind["tax_revenue"]
    summary["External grants"] = ind["external_grants"]
    summary["Total expenditure"] = ind["total_expenditure"]
    summary["Overall balance"] = ind["overall_balance"]
    summary["Primary balance"] = ind["primary_balance"]
    summary["External debt"] = ind["external_debt"]

    summary["Overall balance (% GDP)"] = ind["overall_balance_gdp"]
    summary["Primary balance (% GDP)"] = ind["primary_balance_gdp"]
    summary["Fiscal impulse"] = ind["fiscal_impulse"]
    summary["External debt (% GDP)"] = ind["external_debt_gdp"]

    summary["Revenue check"] = ind["revenue_check"]
    summary["Domestic revenue check"] = ind["domestic_revenue_check"]
    summary["Balance check"] = ind["balance_check"]
    summary["Grants check"] = ind["grants_check"]

    return summary.round(2)


def make_change_table(rows, years=None):
    if years is None:
        sample = next(iter(rows.values()))
        years = sample.dropna().index.tolist()

    y0, y1 = latest_two_years(years)

    table = pd.DataFrame({
        str(y0): {
            k: v.loc[y0] if y0 in v.index else np.nan
            for k, v in rows.items()
        },
        str(y1): {
            k: v.loc[y1] if y1 in v.index else np.nan
            for k, v in rows.items()
        },
    })

    table["Change"] = table[str(y1)] - table[str(y0)]

    table["% change"] = np.where(
        table[str(y0)].abs() > 0,
        table["Change"] / table[str(y0)].abs() * 100,
        np.nan,
    )

    return table.round(1)

def create_revenue_table(ind, years=None):
    rows = {
        "Revenue and grants": ind["revenue_grants"],
        "Domestic revenue": ind["domestic_revenue"],
        "Tax revenue": ind["tax_revenue"],
        "Non-tax revenue": ind["non_tax_revenue"],
        "External grants": ind["external_grants"],
    }
    return make_change_table(rows, years)


def create_expenditure_table(ind, years=None):
    rows = {
        "Total expenditure": ind["total_expenditure"],
        "Recurrent expenditure": ind["recurrent_expenditure"],
        "Compensation": ind["compensation"],
        "Goods/services": ind["goods_services"],
        "Social benefits": ind["social_benefits"],
        "Transfers to FMS": ind["transfers_fms"],
        "Capital expenditure": ind["capital_expenditure"],
    }
    return make_change_table(rows, years)


def create_tax_heads_table(ind, years=None):
    rows = {
        "Tax revenue": ind["tax_revenue"],
        "Customs/trade taxes": ind["customs"],
        "Sales tax/VAT": ind["sales_tax"],
        "Income tax": ind["income_tax"],
        "Corporate income tax": ind["corporate_income_tax"],
        "Excise taxes": ind["excise_taxes"],
    }
    return make_change_table(rows, years)


# ============================================================
# PLOT DASHBOARD
# ============================================================

def plot_fiscal_dashboard(ind, years, save_path=None, start_year=None, end_year=None):
    years_main = select_years(years, start_year, end_year)

    fig, axes = plt.subplots(3, 3, figsize=(20, 14))

    fig.suptitle(
        "Somalia: Fiscal Sector Diagnostics",
        fontsize=22,
        fontweight="bold",
        color="#1f77b4",
    )

    legend_kw = dict(fontsize=8.2, frameon=True, framealpha=0.88)

    # 1. Fiscal stance and macro cycle
    ax = axes[0, 0]

    impulse_vals = ind["fiscal_impulse"].loc[years_main]
    impulse_colors = ["#1f77b4" if x > 0 else "#8B0000" for x in impulse_vals.fillna(0)]

    bars = ax.bar(
        years_main,
        impulse_vals,
        color=impulse_colors,
        alpha=0.90,
        label="Fiscal impulse",
    )

    ax.plot(
        years_main,
        ind["primary_balance_gdp"].loc[years_main],
        color="#ff7f0e",
        marker="o",
        linewidth=2.2,
        linestyle="--",
        label="Primary balance",
    )

    ax2 = ax.twinx()
    ax2.plot(
        years_main,
        ind["output_gap"].loc[years_main],
        color="black",
        marker="s",
        linewidth=2.1,
        linestyle=":",
        label="Output gap",
    )

    label_bars(ax, bars, threshold=0.2)

    style_ax(ax, "Fiscal stance and macro cycle")
    ax.set_ylabel("Percent of GDP", fontsize=9.5)
    ax2.set_ylabel("Output gap, percent", fontsize=9.5)
    ax2.spines["top"].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, **legend_kw)

    # 2. Revenue and grants composition
    ax = axes[0, 1]

    resource_share = ind["resource_composition_share"].loc[years_main].dropna(axis=1, how="all")
    resource_colors = {
        "Tax revenue": "#1f77b4",
        "Non-tax revenue": "#9467bd",
        "External grants": "#ff7f0e",
    }

    bottom = np.zeros(len(years_main))
    stacked_bars = []

    for col in resource_share.columns:
        vals = resource_share[col].fillna(0).values
        bars = ax.bar(years_main, vals, bottom=bottom, color=resource_colors.get(col), label=col)
        stacked_bars.extend(bars)
        bottom += vals

    label_stacked(ax, stacked_bars, threshold=8)

    style_ax(ax, "Revenue and grants composition")
    ax.set_ylabel("Share of total resources, percent", fontsize=9.5)
    ax.set_ylim(0, 105)
    ax.legend(**legend_kw)

    # 3. Fiscal balances and debt stock
    ax = axes[0, 2]

    bars = ax.bar(
        years_main,
        ind["overall_balance_gdp"].loc[years_main],
        color="#1f77b4",
        alpha=0.90,
        label="Overall balance",
    )

    ax.plot(
        years_main,
        ind["primary_balance_gdp"].loc[years_main],
        color="#ff7f0e",
        marker="o",
        linewidth=2.1,
        linestyle="--",
        label="Primary balance",
    )

    ax2 = ax.twinx()
    ax2.plot(
        years_main,
        ind["external_debt"].loc[years_main],
        color="#d62728",
        marker="s",
        linewidth=2.3,
        label="External debt stock",
    )

    label_bars(ax, bars, threshold=0.2)

    style_ax(ax, "Fiscal balances and debt stock")
    ax.set_ylabel("Balance, percent of GDP", fontsize=9.5)
    ax2.set_ylabel("External debt stock, US$ millions", fontsize=9.5)
    ax2.spines["top"].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, **legend_kw)

    # 4. Tax revenue composition
    ax = axes[1, 0]

    tax_share = ind["tax_composition_share"].loc[years_main].dropna(axis=1, how="all")
    tax_colors = {
        "Customs/trade taxes": "#17becf",
        "Sales tax/VAT": "#9467bd",
        "Income tax": "#2ca02c",
        "Corporate income tax": "#8c564b",
        "Excise taxes": "#e377c2",
    }

    bottom = np.zeros(len(years_main))
    stacked_bars = []

    for col in tax_share.columns:
        vals = tax_share[col].fillna(0).values
        bars = ax.bar(years_main, vals, bottom=bottom, color=tax_colors.get(col), label=col)
        stacked_bars.extend(bars)
        bottom += vals

    label_stacked(ax, stacked_bars, threshold=8)

    style_ax(ax, "Tax revenue composition")
    ax.set_ylabel("Share of plotted tax heads, percent", fontsize=9.5)
    ax.set_ylim(0, 105)
    ax.legend(fontsize=8.0, ncol=2, frameon=True, framealpha=0.88)

    # 5. Expenditure composition
    ax = axes[1, 1]

    exp_share = ind["expenditure_composition_share"].loc[years_main].dropna(axis=1, how="all")
    exp_colors = {
        "Compensation": "#1f77b4",
        "Goods/services": "#ff7f0e",
        "Social benefits": "#2ca02c",
        "Interest": "#d62728",
        "Subsidies": "#8c564b",
        "Transfers to FMS": "#9467bd",
        "Capital expenditure": "#e377c2",
    }

    bottom = np.zeros(len(years_main))
    stacked_bars = []

    for col in exp_share.columns:
        vals = exp_share[col].fillna(0).values
        bars = ax.bar(years_main, vals, bottom=bottom, color=exp_colors.get(col), label=col)
        stacked_bars.extend(bars)
        bottom += vals

    label_stacked(ax, stacked_bars, threshold=8)

    style_ax(ax, "Expenditure composition")
    ax.set_ylabel("Share of total expenditure, percent", fontsize=9.5)
    ax.set_ylim(0, 105)
    ax.legend(fontsize=8.0, ncol=2, frameon=True, framealpha=0.88)

    # 6. Grants composition
    ax = axes[1, 2]

    grants_share = ind["grants_composition_share"].loc[years_main].dropna(axis=1, how="all")
    grants_colors = {
        "Bilateral grants": "#1f77b4",
        "Multilateral grants": "#2ca02c",
    }

    bottom = np.zeros(len(years_main))
    stacked_bars = []

    for col in grants_share.columns:
        vals = grants_share[col].fillna(0).values
        bars = ax.bar(years_main, vals, bottom=bottom, color=grants_colors.get(col), label=col)
        stacked_bars.extend(bars)
        bottom += vals

    label_stacked(ax, stacked_bars, threshold=8)

    style_ax(ax, "External grants composition")
    ax.set_ylabel("Share of plotted grants, percent", fontsize=9.5)
    ax.set_ylim(0, 105)
    ax.legend(**legend_kw)

    # 7. Revenue table
    add_table(axes[2, 0], create_revenue_table(ind, years_main), "Revenue, US$ millions")

    # 8. Expenditure table
    add_table(axes[2, 1], create_expenditure_table(ind, years_main), "Expenditure, US$ millions")

    # 9. Tax heads table
    add_table(axes[2, 2], create_tax_heads_table(ind, years_main), "Tax revenue heads, US$ millions")

    fig.text(
        0.01,
        0.01,
        "Source: Somalia macro template; MoF/CBS; staff calculations.",
        fontsize=9,
    )

    plt.tight_layout(rect=[0, 0.035, 1, 0.95])

    if save_path is not None:
        fig.savefig(save_path, dpi=300, bbox_inches="tight")

    return fig


# ============================================================
# MAIN
# ============================================================

def main():
    print("Loading fiscal and real sector data...")
    fiscal_raw, fiscal_years = load_sector_sheet(DATA_FILE, FISCAL_SHEET)
    real_raw, real_years = load_sector_sheet(DATA_FILE, REAL_SHEET)

    years = sorted(list(set(fiscal_years).intersection(set(real_years))))

    print("Computing fiscal indicators...")
    indicators = compute_fiscal_indicators(fiscal_raw, real_raw, years)

    print("Creating fiscal summary table...")
    summary = create_summary_table(indicators, years)

    summary_output = TABLE_DIR / "fiscal_sector_summary.xlsx"
    chart_output = CHART_DIR / "fiscal_sector_dashboard_3x3.png"

    print("Exporting fiscal summary table...")
    summary.to_excel(summary_output)

    print("Generating fiscal dashboard...")
    plot_fiscal_dashboard(indicators, years, save_path=chart_output)

    print("Done.")
    print(f"Summary table saved to: {summary_output}")
    print(f"Dashboard saved to: {chart_output}")


if __name__ == "__main__":
    main()