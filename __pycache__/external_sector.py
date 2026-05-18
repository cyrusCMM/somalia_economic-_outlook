# ============================================================
# external_sector.py
# Somalia Economic Outlook Report
# External Sector Pipeline
# IMF-style 3x3 dashboard
# ============================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path



# ============================================================
# PATHS / DIRECTORIES
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
# LOAD DATA
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

    df = df[[c for c in id_cols if c in df.columns] + year_cols].copy()
    df["Variable"] = df["Variable"].astype(str).str.strip()

    for c in year_cols:
        df[c] = (
            df[c].astype(str)
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


def yoy(x):
    return x.pct_change() * 100


def pct_gdp(x, gdp):
    return x / gdp.replace(0, np.nan) * 100


def share(x, total):
    return x / total.replace(0, np.nan) * 100


def index_100(x, base_year=2018):
    if base_year in x.index and pd.notna(x.loc[base_year]) and x.loc[base_year] != 0:
        return x / x.loc[base_year] * 100

    valid = x.dropna()
    if valid.empty:
        return pd.Series(index=x.index, dtype=float)

    return x / valid.iloc[0] * 100



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
    ax.set_title(title, loc="left", fontsize=11, fontweight="bold")
    ax.axhline(0, linewidth=0.8, color="black")
    ax.grid(True, alpha=0.15)
    ax.tick_params(axis="x", labelsize=9)
    ax.tick_params(axis="y", labelsize=9)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)


def label_bars(ax, bars, threshold=0.2, fmt="{:.1f}", fontsize=8):
    for bar in bars:
        h = bar.get_height()
        if pd.isna(h) or abs(h) < threshold:
            continue

        x = bar.get_x() + bar.get_width() / 2
        y = bar.get_y() + h + 0.10 if h >= 0 else bar.get_y() + h - 0.10
        va = "bottom" if h >= 0 else "top"

        ax.text(x, y, fmt.format(h), ha="center", va=va, fontsize=fontsize)


def label_stacked(ax, bars, threshold=8, fontsize=8):
    for bar in bars:
        h = bar.get_height()
        if pd.isna(h) or abs(h) < threshold:
            continue

        x = bar.get_x() + bar.get_width() / 2
        y = bar.get_y() + h / 2

        ax.text(
            x, y, f"{h:.0f}",
            ha="center",
            va="center",
            fontsize=fontsize,
            color="white",
            fontweight="bold",
        )


def format_table(table_df):
    out = table_df.copy().replace([np.inf, -np.inf], np.nan)
    out = out.round(1)
    return out.astype(object).where(pd.notna(out), "—")


def add_table(ax, table_df, title):
    ax.axis("off")
    ax.set_title(title, loc="left", fontsize=11, fontweight="bold")

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
    tbl.scale(1.05, 1.45)

    for (r, c), cell in tbl.get_celld().items():
        cell.set_linewidth(0.35)
        if r == 0:
            cell.set_text_props(weight="bold")

    return tbl


# ============================================================
# COMPUTE EXTERNAL INDICATORS
# ============================================================

def compute_external_indicators(external_raw, real_raw, years):
    nominal_gdp = get_series(real_raw, years, "Nominal GDP")
    cpi = get_series(real_raw, years, "CPI headline index")
    inflation = yoy(cpi)

    current_account = get_series(external_raw, years, "Current account balance")

    goods_exports = get_series(external_raw, years, "Goods exports FOB")
    goods_imports = get_series(external_raw, years, "Goods imports FOB")
    goods_balance = goods_exports + goods_imports

    services_credit = get_series(external_raw, years, "Services credit")
    services_debit = get_series(external_raw, years, "Services debit")
    services_balance = services_credit + services_debit

    primary_income_credit = get_series(external_raw, years, "Primary income credit")
    primary_income_debit = get_series(external_raw, years, "Primary income debit")
    primary_income_balance = primary_income_credit + primary_income_debit

    secondary_income_credit = get_series(external_raw, years, "Secondary income credit")
    secondary_income_debit = get_series(external_raw, years, "Secondary income debit")
    secondary_income_balance = secondary_income_credit + secondary_income_debit

    individual_remittances = get_series(external_raw, years, "Individual remittances")
    business_transfers = get_series(external_raw, years, "Business transfers")
    ngo_transfers = get_series(external_raw, years, "NGO transfers")
    remittances_total = individual_remittances + business_transfers + ngo_transfers

    fdi = get_series(external_raw, years, "Foreign direct investment")
    other_investment = get_series(external_raw, years, "Other investment")
    portfolio_investment = get_series(external_raw, years, "Portfolio investment")

    reserve_assets_flow = get_series(external_raw, years, "Reserve assets flow")
    gross_reserves = get_series(external_raw, years, "Gross reserve assets stock")

    external_debt = get_series(external_raw, years, "External public debt stock")

    exchange_rate_avg = get_series(external_raw, years, "SOS/USD average")
    exchange_rate_eop = get_series(external_raw, years, "SOS/USD end-period")
    exchange_rate_depreciation = yoy(exchange_rate_avg)

    food_imports = get_series(external_raw, years, "Food imports")
    fuel_imports = get_series(external_raw, years, "Fuel imports")
    capital_goods_imports = get_series(external_raw, years, "Capital goods imports")
    livestock_exports = get_series(external_raw, years, "Livestock exports")

    exports_gs = goods_exports + services_credit
    imports_gs = goods_imports + services_debit
    imports_gs_abs = imports_gs.abs()

    reserves_months = gross_reserves / (imports_gs_abs / 12)

    current_account_gdp = pct_gdp(current_account, nominal_gdp)
    goods_balance_gdp = pct_gdp(goods_balance, nominal_gdp)
    services_balance_gdp = pct_gdp(services_balance, nominal_gdp)
    primary_income_gdp = pct_gdp(primary_income_balance, nominal_gdp)
    secondary_income_gdp = pct_gdp(secondary_income_balance, nominal_gdp)

    goods_exports_gdp = pct_gdp(goods_exports, nominal_gdp)
    goods_imports_gdp = pct_gdp(goods_imports.abs(), nominal_gdp)
    services_credit_gdp = pct_gdp(services_credit, nominal_gdp)
    services_debit_gdp = pct_gdp(services_debit.abs(), nominal_gdp)
    exports_gs_gdp = pct_gdp(exports_gs, nominal_gdp)
    imports_gs_gdp = pct_gdp(imports_gs_abs, nominal_gdp)

    remittances_gdp = pct_gdp(remittances_total, nominal_gdp)
    secondary_income_credit_gdp = pct_gdp(secondary_income_credit, nominal_gdp)
    fdi_gdp = pct_gdp(fdi, nominal_gdp)
    other_investment_gdp = pct_gdp(other_investment, nominal_gdp)
    reserves_gdp = pct_gdp(gross_reserves, nominal_gdp)
    external_debt_gdp = pct_gdp(external_debt, nominal_gdp)

    transfer_components = pd.DataFrame({
        "Individual remittances": individual_remittances,
        "Business transfers": business_transfers,
        "NGO transfers": ngo_transfers,
    })
    transfer_base = transfer_components.sum(axis=1).replace(0, np.nan)
    transfer_composition_share = transfer_components.div(transfer_base, axis=0) * 100

    financing_components = pd.DataFrame({
        "FDI": fdi.abs(),
        "Other investment": other_investment.abs(),
        "Reserve assets flow": reserve_assets_flow.abs(),
    })
    financing_base = financing_components.sum(axis=1).replace(0, np.nan)
    financing_composition_share = financing_components.div(financing_base, axis=0) * 100

    export_index = pd.DataFrame({
        "Goods exports": index_100(goods_exports),
        "Goods imports": index_100(goods_imports.abs()),
        "Services credit": index_100(services_credit),
        "Services debit": index_100(services_debit.abs()),
    })

    export_concentration = share(livestock_exports, goods_exports)

    import_components = pd.DataFrame({
        "Food imports": food_imports.abs(),
        "Fuel imports": fuel_imports.abs(),
        "Capital goods imports": capital_goods_imports.abs(),
    })
    import_components["Other imports"] = goods_imports.abs() - import_components.sum(axis=1)
    import_components["Other imports"] = import_components["Other imports"].where(
        import_components["Other imports"] >= 0, np.nan
    )
    import_base = import_components.sum(axis=1).replace(0, np.nan)
    import_composition_share = import_components.div(import_base, axis=0) * 100

    return {
        "nominal_gdp": nominal_gdp,
        "inflation": inflation,

        "current_account": current_account,
        "goods_exports": goods_exports,
        "goods_imports": goods_imports,
        "goods_balance": goods_balance,
        "services_credit": services_credit,
        "services_debit": services_debit,
        "services_balance": services_balance,
        "primary_income_balance": primary_income_balance,
        "secondary_income_balance": secondary_income_balance,

        "individual_remittances": individual_remittances,
        "business_transfers": business_transfers,
        "ngo_transfers": ngo_transfers,
        "remittances_total": remittances_total,

        "fdi": fdi,
        "other_investment": other_investment,
        "portfolio_investment": portfolio_investment,

        "reserve_assets_flow": reserve_assets_flow,
        "gross_reserves": gross_reserves,
        "reserves_months": reserves_months,

        "external_debt": external_debt,

        "exchange_rate_avg": exchange_rate_avg,
        "exchange_rate_eop": exchange_rate_eop,
        "exchange_rate_depreciation": exchange_rate_depreciation,

        "food_imports": food_imports,
        "fuel_imports": fuel_imports,
        "capital_goods_imports": capital_goods_imports,
        "livestock_exports": livestock_exports,

        "exports_gs": exports_gs,
        "imports_gs": imports_gs,
        "imports_gs_abs": imports_gs_abs,

        "current_account_gdp": current_account_gdp,
        "goods_balance_gdp": goods_balance_gdp,
        "services_balance_gdp": services_balance_gdp,
        "primary_income_gdp": primary_income_gdp,
        "secondary_income_gdp": secondary_income_gdp,

        "goods_exports_gdp": goods_exports_gdp,
        "goods_imports_gdp": goods_imports_gdp,
        "services_credit_gdp": services_credit_gdp,
        "services_debit_gdp": services_debit_gdp,
        "exports_gs_gdp": exports_gs_gdp,
        "imports_gs_gdp": imports_gs_gdp,

        "remittances_gdp": remittances_gdp,
        "secondary_income_credit_gdp": secondary_income_credit_gdp,
        "fdi_gdp": fdi_gdp,
        "other_investment_gdp": other_investment_gdp,
        "reserves_gdp": reserves_gdp,
        "external_debt_gdp": external_debt_gdp,

        "transfer_composition_share": transfer_composition_share,
        "financing_composition_share": financing_composition_share,
        "import_composition_share": import_composition_share,
        "export_index": export_index,
        "export_concentration": export_concentration,
    }


# ============================================================
# TABLES
# ============================================================

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

def create_external_table(ind, years=None):
    rows = {
        "Current account": ind["current_account"],
        "Goods exports": ind["goods_exports"],
        "Goods imports": ind["goods_imports"],
        "Services credit": ind["services_credit"],
        "Services debit": ind["services_debit"],
        "Remittances": ind["remittances_total"],
        "FDI": ind["fdi"],
        "Other investment": ind["other_investment"],
        "Gross reserves": ind["gross_reserves"],
        "External debt": ind["external_debt"],
        "SOS/USD average": ind["exchange_rate_avg"],
        "CA balance (% GDP)": ind["current_account_gdp"],
        "Reserve cover, months": ind["reserves_months"],
    }

    return make_change_table(rows, years)


def create_summary_table(ind, years):
    summary = pd.DataFrame(index=years)

    summary["Current account"] = ind["current_account"]
    summary["Goods balance"] = ind["goods_balance"]
    summary["Services balance"] = ind["services_balance"]
    summary["Primary income balance"] = ind["primary_income_balance"]
    summary["Secondary income balance"] = ind["secondary_income_balance"]
    summary["Goods exports"] = ind["goods_exports"]
    summary["Goods imports"] = ind["goods_imports"]
    summary["Services credit"] = ind["services_credit"]
    summary["Services debit"] = ind["services_debit"]
    summary["Remittances"] = ind["remittances_total"]
    summary["FDI"] = ind["fdi"]
    summary["Other investment"] = ind["other_investment"]
    summary["Gross reserves"] = ind["gross_reserves"]
    summary["Reserves months imports"] = ind["reserves_months"]
    summary["External debt"] = ind["external_debt"]
    summary["SOS/USD average"] = ind["exchange_rate_avg"]

    summary["Current account (% GDP)"] = ind["current_account_gdp"]
    summary["Goods balance (% GDP)"] = ind["goods_balance_gdp"]
    summary["Services balance (% GDP)"] = ind["services_balance_gdp"]
    summary["Secondary income (% GDP)"] = ind["secondary_income_gdp"]
    summary["Imports G&S (% GDP)"] = ind["imports_gs_gdp"]
    summary["Remittances (% GDP)"] = ind["remittances_gdp"]
    summary["External debt (% GDP)"] = ind["external_debt_gdp"]
    summary["Export concentration"] = ind["export_concentration"]

    return summary.round(2)


# ============================================================
# PLOT DASHBOARD
# ============================================================

def plot_external_dashboard(ind, years, save_path=None, start_year=None, end_year=None):
    years_main = select_years(years, start_year, end_year)

    fig, axes = plt.subplots(3, 3, figsize=(20, 14))

    fig.suptitle(
        "Somalia: External Sector Diagnostics",
        fontsize=22,
        fontweight="bold",
        color="#1f77b4",
    )

    legend_kw = dict(fontsize=8.5, frameon=True, framealpha=0.88)

    # 1. Current account decomposition
    ax = axes[0, 0]

    components = {
        "Goods balance": ind["goods_balance_gdp"],
        "Services balance": ind["services_balance_gdp"],
        "Primary income": ind["primary_income_gdp"],
        "Secondary income": ind["secondary_income_gdp"],
    }

    bottom_pos = np.zeros(len(years_main))
    bottom_neg = np.zeros(len(years_main))
    colors = {
        "Goods balance": "#1f77b4",
        "Services balance": "#ff7f0e",
        "Primary income": "#9467bd",
        "Secondary income": "#2ca02c",
    }

    for label, series in components.items():
        vals = series.loc[years_main].fillna(0).values
        pos = np.where(vals > 0, vals, 0)
        neg = np.where(vals < 0, vals, 0)

        ax.bar(years_main, pos, bottom=bottom_pos, color=colors[label], label=label)
        ax.bar(years_main, neg, bottom=bottom_neg, color=colors[label])

        bottom_pos += pos
        bottom_neg += neg

    ca_vals = ind["current_account_gdp"].loc[years_main]
    ax.plot(
        years_main,
        ca_vals,
        color="black",
        marker="o",
        linewidth=2.3,
        linestyle="--",
        label="Current account",
    )

    for x, y in zip(years_main, ca_vals):
        if pd.notna(y):
            ax.text(x, y - 1.2, f"{y:.1f}", ha="center", va="top", fontsize=8)

    style_ax(ax, "Current account decomposition")
    ax.set_ylabel("Percent of GDP", fontsize=9.5)
    ax.legend(fontsize=8, frameon=True, framealpha=0.88, ncol=2)

    # 2. Exchange rate and inflation
    ax = axes[0, 1]

    ax.plot(
        years_main,
        ind["exchange_rate_depreciation"].loc[years_main],
        color="#1f77b4",
        marker="o",
        linewidth=2.3,
        label="SOS/USD depreciation",
    )

    ax.plot(
        years_main,
        ind["inflation"].loc[years_main],
        color="#d62728",
        marker="s",
        linewidth=2.3,
        linestyle="--",
        label="Headline inflation",
    )

    style_ax(ax, "Exchange rate and inflation")
    ax.set_ylabel("Percent", fontsize=9.5)
    ax.legend(**legend_kw)

    # 3. Financial account composition
    ax = axes[0, 2]

    financing = ind["financing_composition_share"].loc[years_main].dropna(axis=1, how="all")
    fin_colors = {
        "FDI": "#d62728",
        "Other investment": "#1f77b4",
        "Reserve assets flow": "#9467bd",
    }

    bottom = np.zeros(len(years_main))
    stacked_bars = []

    for col in financing.columns:
        vals = financing[col].fillna(0).values
        bars = ax.bar(years_main, vals, bottom=bottom, color=fin_colors.get(col), label=col)
        stacked_bars.extend(bars)
        bottom += vals

    label_stacked(ax, stacked_bars, threshold=8)
    style_ax(ax, "Financial account composition")
    ax.set_ylabel("Share of identified flows, percent", fontsize=9.5)
    ax.set_ylim(0, 105)
    ax.legend(**legend_kw)

    # 4. Exports and imports index
    ax = axes[1, 0]

    export_index = ind["export_index"].loc[years_main]

    ax.plot(years_main, export_index["Goods exports"], marker="o", linewidth=2.3, color="#2ca02c", label="Goods exports")
    ax.plot(years_main, export_index["Goods imports"], marker="s", linewidth=2.3, linestyle="--", color="#d62728", label="Goods imports")
    ax.plot(years_main, export_index["Services credit"], marker="^", linewidth=2.1, linestyle=":", color="#1f77b4", label="Services credit")
    ax.plot(years_main, export_index["Services debit"], marker="D", linewidth=2.1, linestyle="-.", color="#ff7f0e", label="Services debit")

    style_ax(ax, "Exports and imports")
    ax.set_ylabel("Index, 2018 = 100", fontsize=9.5)
    ax.legend(fontsize=8, frameon=True, framealpha=0.88, ncol=2)

    # 5. Remittances and transfers composition
    ax = axes[1, 1]

    transfer_share = ind["transfer_composition_share"].loc[years_main].dropna(axis=1, how="all")
    trans_colors = {
        "Individual remittances": "#1f77b4",
        "Business transfers": "#ff7f0e",
        "NGO transfers": "#2ca02c",
    }

    bottom = np.zeros(len(years_main))
    stacked_bars = []

    for col in transfer_share.columns:
        vals = transfer_share[col].fillna(0).values
        bars = ax.bar(years_main, vals, bottom=bottom, color=trans_colors.get(col), label=col)
        stacked_bars.extend(bars)
        bottom += vals

    label_stacked(ax, stacked_bars, threshold=8)
    style_ax(ax, "Remittances and transfer composition")
    ax.set_ylabel("Share of inward transfers, percent", fontsize=9.5)
    ax.set_ylim(0, 105)
    ax.legend(**legend_kw)

    # 6. External debt and current account
    ax = axes[1, 2]

    bars = ax.bar(
        years_main,
        ind["external_debt_gdp"].loc[years_main],
        color="#d62728",
        alpha=0.85,
        label="External debt",
    )

    ax2 = ax.twinx()
    ax2.plot(
        years_main,
        ind["current_account_gdp"].loc[years_main],
        color="black",
        marker="o",
        linewidth=2.2,
        linestyle="--",
        label="Current account",
    )

    label_bars(ax, bars, threshold=1)
    style_ax(ax, "External debt and current account")
    ax.set_ylabel("External debt, percent of GDP", fontsize=9.5)
    ax2.set_ylabel("Current account, percent of GDP", fontsize=9.5)
    ax2.spines["top"].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, **legend_kw)

    # 7. Reserve adequacy
    ax = axes[2, 0]

    bars = ax.bar(
        years_main,
        ind["reserves_months"].loc[years_main],
        color="#9467bd",
        alpha=0.90,
        label="Months of imports",
    )

    ax.axhline(3, linestyle=":", linewidth=1.3, color="gray", label="3-month benchmark")

    ax2 = ax.twinx()
    ax2.plot(
        years_main,
        ind["gross_reserves"].loc[years_main],
        color="#1f77b4",
        marker="o",
        linewidth=2.3,
        label="Gross reserves",
    )

    label_bars(ax, bars, threshold=0.05)
    style_ax(ax, "Reserve adequacy")
    ax.set_ylabel("Months of imports", fontsize=9.5)
    ax2.set_ylabel("Gross reserves, US$ millions", fontsize=9.5)
    ax2.spines["top"].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, **legend_kw)

    # 8. External vulnerability indicators
    ax = axes[2, 1]

    ax.plot(
        years_main,
        ind["imports_gs_gdp"].loc[years_main],
        color="#d62728",
        marker="o",
        linewidth=2.3,
        label="Imports G&S",
    )

    ax.plot(
        years_main,
        ind["remittances_gdp"].loc[years_main],
        color="#1f77b4",
        marker="s",
        linewidth=2.3,
        linestyle="--",
        label="Remittances",
    )

    ax.plot(
        years_main,
        ind["other_investment_gdp"].loc[years_main],
        color="#2ca02c",
        marker="^",
        linewidth=2.1,
        linestyle=":",
        label="Other investment",
    )

    style_ax(ax, "Imports, remittances, and financing")
    ax.set_ylabel("Percent of GDP", fontsize=9.5)
    ax.legend(**legend_kw)

    # 9. External sector table
    add_table(
        axes[2, 2],
        create_external_table(ind, years_main),
        "Selected external indicators",
    )

    fig.text(
        0.01,
        0.01,
        "Source: Somalia macro template; CBS/SNBS/MoF; staff calculations.",
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
    print("Loading external and real sector data...")
    external_raw, external_years = load_sector_sheet(DATA_FILE, EXTERNAL_SHEET)
    real_raw, real_years = load_sector_sheet(DATA_FILE, REAL_SHEET)

    years = sorted(list(set(external_years).intersection(set(real_years))))

    print("Computing external sector indicators...")
    indicators = compute_external_indicators(external_raw, real_raw, years)

    print("Creating external summary table...")
    summary = create_summary_table(indicators, years)

    summary_output = TABLE_DIR / "external_sector_summary.xlsx"
    chart_output = CHART_DIR / "external_sector_dashboard_3x3.png"

    print("Exporting external summary table...")
    summary.to_excel(summary_output)

    print("Generating external dashboard...")
    plot_external_dashboard(indicators, years, save_path=chart_output)

    print("Done.")
    print(f"Summary table saved to: {summary_output}")
    print(f"Dashboard saved to: {chart_output}")


if __name__ == "__main__":
    main()