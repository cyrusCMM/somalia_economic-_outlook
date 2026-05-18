# ============================================================
# monetary_sector.py
# Somalia Economic Outlook Report
# Monetary and Financial Sector Pipeline
# 3x3 PPT-ready dashboard
# ============================================================

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path


PROJECT_DIR = Path(r"C:\Users\hp\Documents\somali economic outlook report")
DATA_FILE = PROJECT_DIR / "Data template.xlsx"

MONETARY_SHEET = "Monetary Financial Raw"
REAL_SHEET = "Real Sector Raw"

OUTPUT_DIR = PROJECT_DIR / "outputs"
CHART_DIR = OUTPUT_DIR / "charts"
TABLE_DIR = OUTPUT_DIR / "tables"

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


def yoy(x):
    return x.pct_change() * 100


def pct_gdp(x, gdp):
    return x / gdp.replace(0, np.nan) * 100


def share(x, total):
    return x / total.replace(0, np.nan) * 100


def clean_zero_as_missing(x):
    y = x.copy()
    y = y.replace(0, np.nan)
    return y


def index_100(x, base_year=2018):
    if base_year in x.index and pd.notna(x.loc[base_year]) and x.loc[base_year] != 0:
        return x / x.loc[base_year] * 100
    first = x.dropna().iloc[0] if not x.dropna().empty else np.nan
    return x / first * 100


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
            ha="center", va="center",
            fontsize=fontsize,
            color="white", fontweight="bold"
        )


def format_table(table_df):
    out = table_df.copy()
    out = out.replace([np.inf, -np.inf], np.nan)
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
    tbl.set_fontsize(8.5)
    tbl.scale(1.10, 1.55)

    for (r, c), cell in tbl.get_celld().items():
        cell.set_linewidth(0.35)
        if r == 0:
            cell.set_text_props(weight="bold")

    return tbl


# ============================================================
# COMPUTE INDICATORS
# ============================================================

def compute_monetary_indicators(monetary_raw, real_raw, years):
    nominal_gdp = get_series(real_raw, years, "Nominal GDP")
    cpi = get_series(real_raw, years, "CPI headline index")
    inflation = yoy(cpi)

    nfa_cbs = get_series(monetary_raw, years, "Net foreign assets of CBS")
    nda_cbs = get_series(monetary_raw, years, "Net domestic assets of CBS")

    m2 = clean_zero_as_missing(get_series(monetary_raw, years, "Broad money (M2)"))
    deposits = get_series(monetary_raw, years, "Customer deposits")
    private_credit = get_series(monetary_raw, years, "Credit to private sector")
    claims_government = get_series(monetary_raw, years, "Claims on government")
    nfa_banks = get_series(monetary_raw, years, "Net foreign assets of banks")

    total_assets = get_series(monetary_raw, years, "Total assets")
    financing_assets = get_series(monetary_raw, years, "Financing assets")
    investment_assets = get_series(monetary_raw, years, "Investment assets")
    total_liabilities = get_series(monetary_raw, years, "Total liabilities")
    equity = get_series(monetary_raw, years, "Shareholders' equity")

    car = get_series(monetary_raw, years, "Capital adequacy ratio")

    rtgs = get_series(monetary_raw, years, "RTGS value")
    ach = get_series(monetary_raw, years, "ACH value")
    payment_transactions = get_series(monetary_raw, years, "Total payment transactions")
    mobile_money_value = get_series(monetary_raw, years, "Mobile money transaction value")

    m2_growth = yoy(m2)
    deposit_growth = yoy(deposits)
    credit_growth = yoy(private_credit)
    asset_growth = yoy(total_assets)

    m2_gdp = pct_gdp(m2, nominal_gdp)
    deposits_gdp = pct_gdp(deposits, nominal_gdp)
    credit_gdp = pct_gdp(private_credit, nominal_gdp)
    assets_gdp = pct_gdp(total_assets, nominal_gdp)

    credit_deposit_ratio = share(private_credit, deposits)
    equity_assets_ratio = share(equity, total_assets)

    bank_asset_composition = pd.DataFrame({
        "Financing assets": financing_assets,
        "Investment assets": investment_assets,
        "Liquid/other assets": total_assets - financing_assets - investment_assets,
    })

    bank_asset_base = bank_asset_composition.sum(axis=1).replace(0, np.nan)
    bank_asset_composition_share = bank_asset_composition.div(bank_asset_base, axis=0) * 100

    return {
        "nominal_gdp": nominal_gdp,
        "inflation": inflation,

        "nfa_cbs": nfa_cbs,
        "nda_cbs": nda_cbs,
        "m2": m2,
        "deposits": deposits,
        "private_credit": private_credit,
        "claims_government": claims_government,
        "nfa_banks": nfa_banks,

        "total_assets": total_assets,
        "financing_assets": financing_assets,
        "investment_assets": investment_assets,
        "total_liabilities": total_liabilities,
        "equity": equity,
        "car": car,

        "rtgs": rtgs,
        "ach": ach,
        "payment_transactions": payment_transactions,
        "mobile_money_value": mobile_money_value,

        "m2_growth": m2_growth,
        "deposit_growth": deposit_growth,
        "credit_growth": credit_growth,
        "asset_growth": asset_growth,

        "m2_gdp": m2_gdp,
        "deposits_gdp": deposits_gdp,
        "credit_gdp": credit_gdp,
        "assets_gdp": assets_gdp,

        "credit_deposit_ratio": credit_deposit_ratio,
        "equity_assets_ratio": equity_assets_ratio,
        "bank_asset_composition_share": bank_asset_composition_share,

        "total_assets_index": index_100(total_assets),
        "financing_assets_index": index_100(financing_assets),
        "investment_assets_index": index_100(investment_assets),
    }


# ============================================================
# TABLES
# ============================================================

def make_change_table(rows, years=(2023, 2024)):
    y0, y1 = years

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


def create_key_indicators_table(ind):
    rows = {
        "M2 growth (%)": ind["m2_growth"],
        "Credit growth (%)": ind["credit_growth"],
        "Deposits/GDP (%)": ind["deposits_gdp"],
        "Credit/GDP (%)": ind["credit_gdp"],
        "Credit/deposit ratio (%)": ind["credit_deposit_ratio"],
        "Bank assets/GDP (%)": ind["assets_gdp"],
        "CAR (%)": ind["car"],
        "Mobile money value": ind["mobile_money_value"],
    }
    return make_change_table(rows)


def create_summary_table(ind, years):
    summary = pd.DataFrame(index=years)

    summary["M2"] = ind["m2"]
    summary["M2 growth"] = ind["m2_growth"]
    summary["Customer deposits"] = ind["deposits"]
    summary["Deposit growth"] = ind["deposit_growth"]
    summary["Private credit"] = ind["private_credit"]
    summary["Private credit growth"] = ind["credit_growth"]
    summary["Credit-to-GDP"] = ind["credit_gdp"]
    summary["Deposits-to-GDP"] = ind["deposits_gdp"]
    summary["Credit-deposit ratio"] = ind["credit_deposit_ratio"]
    summary["Bank assets"] = ind["total_assets"]
    summary["Bank assets/GDP"] = ind["assets_gdp"]
    summary["CAR"] = ind["car"]
    summary["RTGS value"] = ind["rtgs"]
    summary["ACH value"] = ind["ach"]
    summary["Mobile money value"] = ind["mobile_money_value"]

    return summary.round(2)


# ============================================================
# PLOT DASHBOARD
# ============================================================

def plot_monetary_dashboard(ind, years, save_path=None):
    years_main = [y for y in years if y <= 2024]

    fig, axes = plt.subplots(3, 3, figsize=(20, 14))

    fig.suptitle(
        "Somalia: Monetary and Financial Sector Diagnostics",
        fontsize=22,
        fontweight="bold",
        color="#1f77b4",
    )

    legend_kw = dict(fontsize=8.5, frameon=True, framealpha=0.88)

    # --------------------------------------------------------
    # 1. Monetary expansion and inflation
    # --------------------------------------------------------
    ax = axes[0, 0]

    ax.plot(
        years_main,
        ind["m2_growth"].loc[years_main],
        marker="o",
        linewidth=2.3,
        label="M2 growth",
    )

    ax.plot(
        years_main,
        ind["credit_growth"].loc[years_main],
        marker="s",
        linewidth=2.3,
        linestyle="--",
        label="Private credit growth",
    )

    ax.plot(
        years_main,
        ind["inflation"].loc[years_main],
        marker="^",
        linewidth=2.1,
        linestyle=":",
        color="black",
        label="Headline inflation",
    )

    style_ax(ax, "Monetary expansion and inflation")
    ax.set_ylabel("Percent", fontsize=9.5)
    ax.legend(**legend_kw)

    # --------------------------------------------------------
    # 2. Financial deepening
    # --------------------------------------------------------
    ax = axes[0, 1]

    ax.plot(
        years_main,
        ind["deposits_gdp"].loc[years_main],
        marker="o",
        linewidth=2.4,
        label="Customer deposits",
    )

    ax.plot(
        years_main,
        ind["credit_gdp"].loc[years_main],
        marker="s",
        linewidth=2.4,
        linestyle="--",
        label="Private credit",
    )

    ax.plot(
        years_main,
        ind["assets_gdp"].loc[years_main],
        marker="^",
        linewidth=2.2,
        linestyle=":",
        label="Bank assets",
    )

    style_ax(ax, "Financial deepening")
    ax.set_ylabel("Percent of GDP", fontsize=9.5)
    ax.legend(**legend_kw)

    # --------------------------------------------------------
    # 3. Credit intermediation
    # --------------------------------------------------------
    ax = axes[0, 2]

    bars = ax.bar(
        years_main,
        ind["credit_deposit_ratio"].loc[years_main],
        color="#1f77b4",
        alpha=0.90,
        label="Credit/deposit ratio",
    )

    ax.axhline(50, linestyle=":", linewidth=1.2, color="gray")

    ax2 = ax.twinx()
    ax2.plot(
        years_main,
        ind["credit_growth"].loc[years_main],
        color="#ff7f0e",
        marker="o",
        linewidth=2.3,
        linestyle="--",
        label="Private credit growth",
    )

    label_bars(ax, bars, threshold=5, fmt="{:.0f}")

    style_ax(ax, "Credit intermediation")
    ax.set_ylabel("Credit/deposit ratio, percent", fontsize=9.5)
    ax2.set_ylabel("Credit growth, percent", fontsize=9.5)
    ax2.spines["top"].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, **legend_kw)

    # --------------------------------------------------------
    # 4. Banking balance sheet growth index
    # --------------------------------------------------------
    ax = axes[1, 0]

    ax.plot(
        years_main,
        ind["total_assets_index"].loc[years_main],
        marker="o",
        linewidth=2.4,
        label="Total assets",
    )

    ax.plot(
        years_main,
        ind["financing_assets_index"].loc[years_main],
        marker="s",
        linewidth=2.2,
        linestyle="--",
        label="Financing assets",
    )

    ax.plot(
        years_main,
        ind["investment_assets_index"].loc[years_main],
        marker="^",
        linewidth=2.2,
        linestyle=":",
        label="Investment assets",
    )

    style_ax(ax, "Banking balance sheet growth")
    ax.set_ylabel("Index, 2018 = 100", fontsize=9.5)
    ax.legend(**legend_kw)

    # --------------------------------------------------------
    # 5. Credit and deposits
    # --------------------------------------------------------
    ax = axes[1, 1]

    ax.plot(
        years_main,
        ind["deposits"].loc[years_main],
        marker="o",
        linewidth=2.4,
        label="Customer deposits",
    )

    ax.plot(
        years_main,
        ind["private_credit"].loc[years_main],
        marker="s",
        linewidth=2.4,
        linestyle="--",
        label="Private credit",
    )

    ax2 = ax.twinx()

    ax2.plot(
        years_main,
        ind["credit_deposit_ratio"].loc[years_main],
        color="#d62728",
        marker="^",
        linewidth=2.1,
        linestyle=":",
        label="Credit/deposit ratio",
    )

    style_ax(ax, "Credit and deposits")
    ax.set_ylabel("US$ millions", fontsize=9.5)
    ax2.set_ylabel("Credit/deposit ratio, percent", fontsize=9.5)
    ax2.spines["top"].set_visible(False)

    lines1, labels1 = ax.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax.legend(lines1 + lines2, labels1 + labels2, **legend_kw)

    # --------------------------------------------------------
    # 6. Bank asset composition
    # --------------------------------------------------------
    ax = axes[1, 2]

    comp = ind["bank_asset_composition_share"].loc[years_main].dropna(axis=1, how="all")

    colors = {
        "Financing assets": "#1f77b4",
        "Investment assets": "#ff7f0e",
        "Liquid/other assets": "#9467bd",
    }

    bottom = np.zeros(len(years_main))
    stacked_bars = []

    for col in comp.columns:
        vals = comp[col].fillna(0).values
        bars = ax.bar(
            years_main,
            vals,
            bottom=bottom,
            color=colors.get(col),
            label=col,
        )
        stacked_bars.extend(bars)
        bottom += vals

    label_stacked(ax, stacked_bars, threshold=8)

    style_ax(ax, "Bank asset composition")
    ax.set_ylabel("Share of total assets, percent", fontsize=9.5)
    ax.set_ylim(0, 105)
    ax.legend(**legend_kw)

    # --------------------------------------------------------
    # 7. Payments system values
    # --------------------------------------------------------
    ax = axes[2, 0]

    ax.plot(
        years_main,
        ind["rtgs"].loc[years_main],
        marker="o",
        linewidth=2.4,
        label="RTGS value",
    )

    ax.plot(
        years_main,
        ind["ach"].loc[years_main],
        marker="s",
        linewidth=2.4,
        linestyle="--",
        label="ACH value",
    )

    style_ax(ax, "Payments system values")
    ax.set_ylabel("US$ millions", fontsize=9.5)
    ax.legend(**legend_kw)

    # --------------------------------------------------------
    # 8. Mobile money ecosystem
    # --------------------------------------------------------
    ax = axes[2, 1]

    ax.plot(
        years_main,
        ind["mobile_money_value"].loc[years_main],
        marker="o",
        linewidth=2.4,
        label="Mobile money value",
    )

    ax.plot(
        years_main,
        ind["rtgs"].loc[years_main],
        marker="s",
        linewidth=2.2,
        linestyle="--",
        label="RTGS value",
    )

    ax.plot(
        years_main,
        ind["ach"].loc[years_main],
        marker="^",
        linewidth=2.2,
        linestyle=":",
        label="ACH value",
    )

    style_ax(ax, "Mobile money and formal payments")
    ax.set_ylabel("US$ millions", fontsize=9.5)
    ax.legend(**legend_kw)

    # --------------------------------------------------------
    # 9. Key monetary indicators table
    # --------------------------------------------------------
    add_table(
        axes[2, 2],
        create_key_indicators_table(ind),
        "Key monetary indicators",
    )

    fig.text(
        0.01,
        0.01,
        "Source: Somalia macro template; CBS; staff calculations.",
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
    print("Loading monetary and real sector data...")
    monetary_raw, monetary_years = load_sector_sheet(DATA_FILE, MONETARY_SHEET)
    real_raw, real_years = load_sector_sheet(DATA_FILE, REAL_SHEET)

    years = sorted(list(set(monetary_years).intersection(set(real_years))))

    print("Computing monetary and financial indicators...")
    indicators = compute_monetary_indicators(monetary_raw, real_raw, years)

    print("Creating monetary summary table...")
    summary = create_summary_table(indicators, years)

    summary_output = TABLE_DIR / "monetary_sector_summary.xlsx"
    chart_output = CHART_DIR / "monetary_sector_dashboard_3x3.png"

    print("Exporting monetary summary table...")
    summary.to_excel(summary_output)

    print("Generating monetary dashboard...")
    plot_monetary_dashboard(indicators, years, save_path=chart_output)

    print("Done.")
    print(f"Summary table saved to: {summary_output}")
    print(f"Dashboard saved to: {chart_output}")


if __name__ == "__main__":
    main()