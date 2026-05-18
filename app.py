# ============================================================
# app.py
# Somalia Economic Outlook Dashboard
# Streamlit deployment app
# ============================================================

import streamlit as st
from pathlib import Path
from io import BytesIO
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
st.caption("Automated macroeconomic diagnostics and downloadable sector dashboards.")


# ============================================================
# PATHS
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
# SECTOR CONFIG
# ============================================================

SECTOR_CONFIG = {
    "Real Sector": {
        "title": "Real Sector Diagnostics",
        "sheet": "Real Sector Raw",
        "summary_name": "real_sector_summary.xlsx",
        "figure_name": "real_sector_dashboard.png",
    },
    "Fiscal Sector": {
        "title": "Fiscal Sector Diagnostics",
        "sheet": "Fiscal Sector Raw",
        "summary_name": "fiscal_sector_summary.xlsx",
        "figure_name": "fiscal_sector_dashboard.png",
    },
    "Monetary and Financial Sector": {
        "title": "Monetary and Financial Sector Diagnostics",
        "sheet": "Monetary Financial Raw",
        "summary_name": "monetary_sector_summary.xlsx",
        "figure_name": "monetary_sector_dashboard.png",
    },
    "External Sector": {
        "title": "External Sector Diagnostics",
        "sheet": "External Sector Raw",
        "summary_name": "external_sector_summary.xlsx",
        "figure_name": "external_sector_dashboard.png",
    },
}

sector = st.sidebar.selectbox(
    "Select sector",
    list(SECTOR_CONFIG.keys()),
)

st.sidebar.markdown("---")
st.sidebar.caption("Dashboards use the sector modules and Excel template calculations.")


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
            "In the sector PY file, replace `plt.show()` with `return fig`."
        )
        st.stop()

    st.pyplot(fig, use_container_width=True)


# Need pandas here for Excel writer
import pandas as pd


# ============================================================
# MAIN APP
# ============================================================

try:
    config = SECTOR_CONFIG[sector]
    st.subheader(config["title"])

    if sector == "Real Sector":
        real_raw, years = load_real_sheet(DATA_FILE, "Real Sector Raw")
        indicators = compute_real_sector_indicators(real_raw, years)
        summary = create_real_summary(indicators, years)
        fig = plot_real_sector_dashboard(indicators, years, save_path=None)

    elif sector == "Fiscal Sector":
        fiscal_raw, fiscal_years = load_fiscal_sheet(DATA_FILE, "Fiscal Sector Raw")
        real_raw, real_years = load_fiscal_sheet(DATA_FILE, "Real Sector Raw")

        years = sorted(list(set(fiscal_years).intersection(set(real_years))))

        indicators = compute_fiscal_indicators(fiscal_raw, real_raw, years)
        summary = create_fiscal_summary(indicators, years)
        fig = plot_fiscal_dashboard(indicators, years, save_path=None)

    elif sector == "Monetary and Financial Sector":
        monetary_raw, monetary_years = load_monetary_sheet(DATA_FILE, "Monetary Financial Raw")
        real_raw, real_years = load_monetary_sheet(DATA_FILE, "Real Sector Raw")

        years = sorted(list(set(monetary_years).intersection(set(real_years))))

        indicators = compute_monetary_indicators(monetary_raw, real_raw, years)
        summary = create_monetary_summary(indicators, years)
        fig = plot_monetary_dashboard(indicators, years, save_path=None)

    elif sector == "External Sector":
        external_raw, external_years = load_external_sheet(DATA_FILE, "External Sector Raw")
        real_raw, real_years = load_external_sheet(DATA_FILE, "Real Sector Raw")

        years = sorted(list(set(external_years).intersection(set(real_years))))

        indicators = compute_external_indicators(external_raw, real_raw, years)
        summary = create_external_summary(indicators, years)
        fig = plot_external_dashboard(indicators, years, save_path=None)

    # -----------------------------
    # DISPLAY DASHBOARD
    # -----------------------------

    show_figure(fig)

    st.markdown("---")

    # -----------------------------
    # DOWNLOAD BUTTONS
    # -----------------------------

    col1, col2 = st.columns(2)

    with col1:
        excel_bytes = dataframe_to_excel_bytes(summary, sheet_name="Summary")

        st.download_button(
            label="Download summary table Excel",
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

    # -----------------------------
    # SHOW SUMMARY TABLE
    # -----------------------------

    with st.expander("View summary table used for dashboard calculations"):
        st.dataframe(summary, use_container_width=True)

    plt.close(fig)

except FileNotFoundError:
    st.error(
        "Data file not found. Upload the Excel file in the sidebar, "
        "or check the local path."
    )

except Exception as e:
    st.error("Dashboard failed to load.")
    st.exception(e)