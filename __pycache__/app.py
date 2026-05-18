# ============================================================
# app.py
# Somalia Economic Outlook Dashboard
# Streamlit deployment app
# ============================================================

import streamlit as st
from pathlib import Path
from io import BytesIO
import pandas as pd
import matplotlib.pyplot as plt

from real_sector import (
    load_sector_sheet as load_real_sheet,
    compute_real_sector_indicators,
    create_summary_table as create_real_summary,
    plot_real_sector_dashboard,
)

from fiscal_sector import (
    load_sector_sheet as load_fiscal_sheet,
    compute_fiscal_indicators,
    create_summary_table as create_fiscal_summary,
    plot_fiscal_dashboard,
)

from monetary_sector import (
    load_sector_sheet as load_monetary_sheet,
    compute_monetary_indicators,
    create_summary_table as create_monetary_summary,
    plot_monetary_dashboard,
)

from external_sector import (
    load_sector_sheet as load_external_sheet,
    compute_external_indicators,
    create_summary_table as create_external_summary,
    plot_external_dashboard,
)


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Somalia Economic Outlook",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("Somalia Economic Outlook")
st.caption("Automated macroeconomic diagnostics, year selection, and downloadable sector dashboards.")


# ============================================================
# PATHS / DIRECTORIES
# ============================================================

PROJECT_DIR = Path(__file__).parent
DATA_FILE = PROJECT_DIR / "Data template.xlsx"

REAL_SHEET = "Real Sector Raw"
FISCAL_SHEET = "Fiscal Sector Raw"
MONETARY_SHEET = "Monetary Financial Raw"
EXTERNAL_SHEET = "External Sector Raw"


# ============================================================
# SECTOR CONFIG
# ============================================================

SECTOR_CONFIG = {
    "Real Sector": {
        "title": "Real Sector Diagnostics",
        "summary_name": "real_sector_summary.xlsx",
        "figure_name": "real_sector_dashboard.png",
    },
    "Fiscal Sector": {
        "title": "Fiscal Sector Diagnostics",
        "summary_name": "fiscal_sector_summary.xlsx",
        "figure_name": "fiscal_sector_dashboard.png",
    },
    "Monetary and Financial Sector": {
        "title": "Monetary and Financial Sector Diagnostics",
        "summary_name": "monetary_sector_summary.xlsx",
        "figure_name": "monetary_sector_dashboard.png",
    },
    "External Sector": {
        "title": "External Sector Diagnostics",
        "summary_name": "external_sector_summary.xlsx",
        "figure_name": "external_sector_dashboard.png",
    },
}


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("Data and dashboard controls")

uploaded_file = st.sidebar.file_uploader(
    "Upload updated data template",
    type=["xlsx"],
    help="Optional. If no file is uploaded, the app uses Data template.xlsx from the GitHub repo.",
)

if uploaded_file is not None:
    DATA_SOURCE = uploaded_file
    st.sidebar.success("Using uploaded Excel file")
else:
    DATA_SOURCE = DATA_FILE
    st.sidebar.info("Using Data template.xlsx from the app repository")

sector = st.sidebar.selectbox("Select sector", list(SECTOR_CONFIG.keys()))


# ============================================================
# DOWNLOAD HELPERS
# ============================================================

def dataframe_to_excel_bytes(df, sheet_name="Summary"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name)
    output.seek(0)
    return output


def figure_to_png_bytes(fig):
    output = BytesIO()
    fig.savefig(output, format="png", dpi=300, bbox_inches="tight")
    output.seek(0)
    return output


def show_figure(fig):
    if fig is None:
        st.error(
            "The plot function did not return a figure. "
            "In the sector PY file, replace plt.show() with return fig."
        )
        st.stop()
    st.pyplot(fig, use_container_width=True)


def select_year_range(years, sector_name):
    years = sorted([int(y) for y in years])
    if not years:
        st.error("No year columns were found in the selected data sheet.")
        st.stop()

    min_year = min(years)
    max_year = max(years)

    st.sidebar.markdown("---")
    st.sidebar.subheader("Year range")

    if min_year == max_year:
        st.sidebar.info(f"Only {min_year} is available.")
        return min_year, max_year, years

    start_year, end_year = st.sidebar.slider(
        "Select years shown in charts and tables",
        min_value=min_year,
        max_value=max_year,
        value=(min_year, max_year),
        step=1,
        key=f"year_slider_{sector_name}",
    )

    selected_years = [y for y in years if start_year <= y <= end_year]

    if len(selected_years) < 2:
        st.sidebar.warning("Select at least two years so change tables can be calculated.")

    return start_year, end_year, selected_years


def filter_summary(summary, selected_years):
    selected_years = [y for y in selected_years if y in summary.index]
    if not selected_years:
        return summary
    return summary.loc[selected_years]


# ============================================================
# MAIN APP
# ============================================================

try:
    config = SECTOR_CONFIG[sector]
    st.subheader(config["title"])

    if sector == "Real Sector":
        real_raw, years = load_real_sheet(DATA_SOURCE, REAL_SHEET)
        indicators = compute_real_sector_indicators(real_raw, years)
        summary = create_real_summary(indicators, years)

        # Real-sector charts default to years with core real GDP data.
        # If users add a new year of Real GDP, it appears automatically.
        available_years = indicators["real_gdp"].dropna().index.tolist()
        start_year, end_year, selected_years = select_year_range(available_years, sector)

        summary = filter_summary(summary, selected_years)
        fig = plot_real_sector_dashboard(
            indicators,
            years,
            save_path=None,
            start_year=start_year,
            end_year=end_year,
        )

    elif sector == "Fiscal Sector":
        fiscal_raw, fiscal_years = load_fiscal_sheet(DATA_SOURCE, FISCAL_SHEET)
        real_raw, real_years = load_fiscal_sheet(DATA_SOURCE, REAL_SHEET)
        years = sorted(list(set(fiscal_years).intersection(set(real_years))))
        start_year, end_year, selected_years = select_year_range(years, sector)

        indicators = compute_fiscal_indicators(fiscal_raw, real_raw, years)
        summary = create_fiscal_summary(indicators, years)
        summary = filter_summary(summary, selected_years)
        fig = plot_fiscal_dashboard(
            indicators,
            years,
            save_path=None,
            start_year=start_year,
            end_year=end_year,
        )

    elif sector == "Monetary and Financial Sector":
        monetary_raw, monetary_years = load_monetary_sheet(DATA_SOURCE, MONETARY_SHEET)
        real_raw, real_years = load_monetary_sheet(DATA_SOURCE, REAL_SHEET)
        years = sorted(list(set(monetary_years).intersection(set(real_years))))
        start_year, end_year, selected_years = select_year_range(years, sector)

        indicators = compute_monetary_indicators(monetary_raw, real_raw, years)
        summary = create_monetary_summary(indicators, years)
        summary = filter_summary(summary, selected_years)
        fig = plot_monetary_dashboard(
            indicators,
            years,
            save_path=None,
            start_year=start_year,
            end_year=end_year,
        )

    elif sector == "External Sector":
        external_raw, external_years = load_external_sheet(DATA_SOURCE, EXTERNAL_SHEET)
        real_raw, real_years = load_external_sheet(DATA_SOURCE, REAL_SHEET)
        years = sorted(list(set(external_years).intersection(set(real_years))))
        start_year, end_year, selected_years = select_year_range(years, sector)

        indicators = compute_external_indicators(external_raw, real_raw, years)
        summary = create_external_summary(indicators, years)
        summary = filter_summary(summary, selected_years)
        fig = plot_external_dashboard(
            indicators,
            years,
            save_path=None,
            start_year=start_year,
            end_year=end_year,
        )

    show_figure(fig)

    st.markdown("---")
    st.caption(
        f"Selected period: {selected_years[0]}-{selected_years[-1]}. "
        "Change tables inside dashboards use the latest two years in the selected period."
    )

    col1, col2 = st.columns(2)

    with col1:
        excel_bytes = dataframe_to_excel_bytes(summary, sheet_name="Summary")
        st.download_button(
            label="Download selected-period summary Excel",
            data=excel_bytes,
            file_name=config["summary_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col2:
        png_bytes = figure_to_png_bytes(fig)
        st.download_button(
            label="Download dashboard PNG",
            data=png_bytes,
            file_name=config["figure_name"],
            mime="image/png",
        )

    with st.expander("View summary table used for dashboard calculations"):
        st.dataframe(summary, use_container_width=True)

    plt.close(fig)

except FileNotFoundError:
    st.error(
        "Data file not found. Upload the Excel file in the sidebar, "
        "or make sure Data template.xlsx is included in the GitHub repository."
    )

except Exception as e:
    st.error("Dashboard failed to load.")
    st.exception(e)
