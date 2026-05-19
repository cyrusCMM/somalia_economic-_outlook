# ============================================================
# app.py
# CBS Somalia Economic Outlook System
# Streamlit app with simple password login + role-based access
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
# SIMPLE PASSWORD LOGIN
# ============================================================

# For immediate deployment stability, this uses simple password login.
# Later, these passwords can be moved fully into Streamlit Secrets.
USERS = {
    "mutukucmm@gmail.com": {
        "name": "Admin",
        "role": "admin",
        "password": st.secrets.get("ADMIN_PASSWORD", "admin123"),
    },
    "anisa.osman@centralbank.gov.so": {
        "name": "Anisa Osman",
        "role": "viewer",
        "password": st.secrets.get("VIEWER_PASSWORD", "viewer123"),
    },
}


def login_gate():
    """Stop the app until a valid user logs in."""

    if "logged_in" not in st.session_state:
        st.session_state["logged_in"] = False

    if "current_user" not in st.session_state:
        st.session_state["current_user"] = ""

    # If a stale/blank user exists, reset safely.
    current_user = st.session_state.get("current_user", "")
    if st.session_state.get("logged_in") and current_user not in USERS:
        st.session_state["logged_in"] = False
        st.session_state["current_user"] = ""

    if not st.session_state["logged_in"]:
        st.title("CBS Somalia Economic Outlook System")
        st.info("Please log in with your authorized email and password.")

        with st.form("login_form", clear_on_submit=False):
            email = st.text_input("Email").strip().lower()
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log in")

        if submitted:
            if email in USERS and password == USERS[email]["password"]:
                st.session_state["logged_in"] = True
                st.session_state["current_user"] = email
                st.rerun()
            else:
                st.error("Invalid email or password.")

        st.stop()

    user_email = st.session_state.get("current_user", "")

    if user_email not in USERS:
        st.session_state["logged_in"] = False
        st.session_state["current_user"] = ""
        st.rerun()

    return user_email, USERS[user_email]["name"], USERS[user_email]["role"]


USER_EMAIL, USER_NAME, USER_ROLE = login_gate()
IS_ADMIN = USER_ROLE == "admin"


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
# SIDEBAR USER + DATA CONFIGURATION
# ============================================================

with st.sidebar:
    st.success(f"Logged in: {USER_NAME}")
    st.caption(f"Role: {USER_ROLE}")

    if st.button("Log out"):
        st.session_state["logged_in"] = False
        st.session_state["current_user"] = ""
        st.rerun()

    st.markdown("---")
    st.header("Data configuration")

if IS_ADMIN:
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

    if len(all_years) < 2:
        st.error("At least two years of data are required.")
        st.stop()

    start_year, end_year = st.sidebar.select_slider(
        "Select year range",
        options=all_years,
        value=(min(all_years), max(all_years)),
    )

    selected_years = [y for y in all_years if start_year <= y <= end_year]

    if len(selected_years) < 2:
        st.warning("Please select at least two years.")
        st.stop()

    return selected_years


def load_all_data(data_file):
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


def ensure_enough_years(years, sector_name):
    if len(years) < 2:
        st.warning(f"{sector_name} has fewer than two valid years in the selected range.")
        st.stop()


# ============================================================
# MAIN APP
# ============================================================

try:
    loaded = load_all_data(DATA_FILE)

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
        ensure_enough_years(years, sector)

        indicators = compute_real_sector_indicators(loaded["real_raw"], years)
        summary = create_real_summary(indicators, years)
        fig = plot_real_sector_dashboard(indicators, years, save_path=None)

    elif sector == "Fiscal Sector":
        years = sorted(
            set(loaded["fiscal_years"])
            .intersection(set(loaded["real_years"]))
            .intersection(selected_years)
        )
        ensure_enough_years(years, sector)

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
        ensure_enough_years(years, sector)

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
        ensure_enough_years(years, sector)

        indicators = compute_external_indicators(
            loaded["external_raw"],
            loaded["real_raw"],
            years,
        )
        summary = create_external_summary(indicators, years)
        fig = plot_external_dashboard(indicators, years, save_path=None)

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

    if IS_ADMIN:
        with st.expander("Admin information"):
            st.write("Current data source:", DATA_FILE)
            st.write("Available years:", all_years)
            st.write("Selected sector years:", years)
            st.write("Current user:", USER_EMAIL)

    plt.close(fig)

except FileNotFoundError:
    st.error(
        "Data file not found. Upload the Excel file in the sidebar, "
        "or check that `Data template.xlsx` exists in the repository."
    )

except Exception as e:
    st.error("Dashboard failed to load.")
    st.exception(e)
