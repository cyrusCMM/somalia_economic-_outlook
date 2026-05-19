# ============================================================
# app.py
# CBS Somalia Economic Outlook System
# Streamlit deployment app with role-based access
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
    page_title="CBS Somalia Economic Outlook System",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# LOGIN AND ROLE ACCESS
# ============================================================

USERS = {
    "mutukucmm@gmail.com": {
        "name": "Admin",
        "role": "admin",
    },
    "anisa.osman@centralbank.gov.so": {
        "name": "Anisa Osman",
        "role": "viewer",
    },
}

if not st.user.is_logged_in:
    st.title("CBS Somalia Economic Outlook System")
    st.info("Please log in with your authorized email to access the system.")
    st.button("Log in", on_click=st.login)
    st.stop()

user_email = st.user.get("email", "").lower()

if user_email not in USERS:
    st.error("You are logged in, but you are not authorized to access this system.")
    st.write(f"Logged in as: `{user_email}`")
    st.button("Log out", on_click=st.logout)
    st.stop()

USER_NAME = USERS[user_email]["name"]
USER_ROLE = USERS[user_email]["role"]


# ============================================================
# PATHS
# ============================================================

PROJECT_DIR = Path(__file__).parent
DEFAULT_DATA_FILE = PROJECT_DIR / "Data template.xlsx"

OUTPUT_DIR = PROJECT_DIR / "outputs"
CHART_DIR = OUTPUT_DIR / "charts"
TABLE_DIR = OUTPUT_DIR / "tables"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
CHART_DIR.mkdir(parents=True, exist_ok=True)
TABLE_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# APP HEADER
# ============================================================

st.title("CBS Somalia Economic Outlook System")
st.caption("Macro diagnostics, sector dashboards, and downloadable analytical outputs.")


# ============================================================
# SIDEBAR: USER AND DATA CONFIGURATION
# ============================================================

with st.sidebar:
    st.success(f"Logged in: {USER_NAME}")
    st.caption(f"Role: {USER_ROLE}")
    st.button("Log out", on_click=st.logout)

    st.markdown("---")
    st.header("Data configuration")

if USER_ROLE == "admin":
    uploaded_file = st.sidebar.file_uploader(
        "Upload updated Somalia data template",
        type=["xlsx"],
    )

    if uploaded_file is not None:
        DATA_FILE = uploaded_file
        st.sidebar.success("Using uploaded Excel file for this session.")
    else:
        DATA_FILE = DEFAULT_DATA_FILE
        st.sidebar.info("Using default Excel template from repository.")
else:
    DATA_FILE = DEFAULT_DATA_FILE
    st.sidebar.info("Viewer access: downloads enabled, data upload disabled.")


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

sector = st.sidebar.selectbox(
    "Select sector",
    list(SECTOR_CONFIG.keys()),
)


# ============================================================
# HELPERS
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


def select_year_range(all_years):
    all_years = sorted([int(y) for y in all_years])
    min_year = min(all_years)
    max_year = max(all_years)

    start_year, end_year = st.sidebar.select_slider(
        "Select year range",
        options=all_years,
        value=(min_year, max_year),
    )

    selected_years = [y for y in all_years if start_year <= y <= end_year]

    if len(selected_years) < 2:
        st.warning("Please select at least two years.")
        st.stop()

    return selected_years


def load_all_years(data_file):
    real_raw, real_years = load_real_sheet(data_file, "Real Sector Raw")
    fiscal_raw, fiscal_years = load_fiscal_sheet(data_file, "Fiscal Sector Raw")
    monetary_raw, monetary_years = load_monetary_sheet(data_file, "Monetary Financial Raw")
    external_raw, external_years = load_external_sheet(data_file, "External Sector Raw")

    return {
        "real_raw": real_raw,
        "real_years": real_years,
        "fiscal_raw": fiscal_raw,
        "fiscal_years": fiscal_years,
        "monetary_raw": monetary_raw,
        "monetary_years": monetary_years,
        "external_raw": external_raw,
        "external_years": external_years,
    }


# ============================================================
# MAIN APP
# ============================================================

try:
    loaded = load_all_years(DATA_FILE)

    all_years = sorted(
        set(loaded["real_years"])
        | set(loaded["fiscal_years"])
        | set(loaded["monetary_years"])
        | set(loaded["external_years"])
    )

    st.sidebar.markdown("---")
    st.sidebar.header("Dashboard settings")
    selected_years = select_year_range(all_years)

    config = SECTOR_CONFIG[sector]
    st.subheader(config["title"])
    st.caption(
        f"Showing {min(selected_years)}–{max(selected_years)}. "
        "Change tables use the latest two years in the selected range."
    )

    if sector == "Real Sector":
        years = sorted(set(loaded["real_years"]).intersection(selected_years))
        indicators = compute_real_sector_indicators(loaded["real_raw"], years)
        summary = create_real_summary(indicators, years)
        fig = plot_real_sector_dashboard(indicators, years, save_path=None)

    elif sector == "Fiscal Sector":
        years = sorted(
            set(loaded["fiscal_years"])
            .intersection(set(loaded["real_years"]))
            .intersection(selected_years)
        )
        indicators = compute_fiscal_indicators(
            loaded["fiscal_raw"],
            loaded["real_raw"],
            years,
        )
        summary = create_fiscal_summary(indicators, years)
        fig = plot_fiscal_dashboard(indicators, years, save_path=None)

    elif sector == "Monetary and Financial Sector":
        years = sorted(
            set(loaded["monetary_years"])
            .intersection(set(loaded["real_years"]))
            .intersection(selected_years)
        )
        indicators = compute_monetary_indicators(
            loaded["monetary_raw"],
            loaded["real_raw"],
            years,
        )
        summary = create_monetary_summary(indicators, years)
        fig = plot_monetary_dashboard(indicators, years, save_path=None)

    elif sector == "External Sector":
        years = sorted(
            set(loaded["external_years"])
            .intersection(set(loaded["real_years"]))
            .intersection(selected_years)
        )
        indicators = compute_external_indicators(
            loaded["external_raw"],
            loaded["real_raw"],
            years,
        )
        summary = create_external_summary(indicators, years)
        fig = plot_external_dashboard(indicators, years, save_path=None)

    if len(years) < 2:
        st.warning("The selected sector has fewer than two valid years in this range.")
        st.stop()

    show_figure(fig)

    st.markdown("---")

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

    with st.expander("View summary table used for dashboard calculations"):
        st.dataframe(summary, use_container_width=True)

    if USER_ROLE == "admin":
        with st.expander("Admin information"):
            st.write("Current data source:", DATA_FILE)
            st.write("Available years:", all_years)
            st.write("Selected years:", years)

    plt.close(fig)

except FileNotFoundError:
    st.error(
        "Data file not found. Upload the Excel file in the sidebar, "
        "or check that `Data template.xlsx` exists in the repository."
    )

except Exception as e:
    st.error("Dashboard failed to load.")
    st.exception(e)