import streamlit as st
import pandas as pd
from html import escape

from itam.audit import calculate_asset_age
from itam.audit import get_warranty_status
from itam.audit import run_itam_audit
from itam.data import (
    build_canonical_dataframe,
    detect_asset_type_from_data,
    detect_header_row,
    validate_source_columns,
)
from itam.ui import render_navigation

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================
st.set_page_config(
    page_title="AssetLens",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================================
# CONSTANTS
# ============================================================================
# ============================================================================
# AIR SELANGOR THEME CSS
# ============================================================================
def inject_professional_css():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');

        :root {
            --al-navy: #0F2744;
            --al-blue: #176B87;
            --al-teal: #1F8A8A;
            --al-bg: #F4F7FA;
            --al-surface: #FFFFFF;
            --al-border: #D9E2EC;
            --al-text: #172B4D;
            --al-muted: #6B7C93;
            --al-success: #2E7D5B;
            --al-warning: #B7791F;
            --al-danger: #B84242;
            --al-info: #176B87;
            --primary-blue: var(--al-navy);
            --secondary-blue: var(--al-blue);
            --light-blue: #E7F0F5;
            --text-primary: var(--al-text);
            --text-secondary: var(--al-muted);
            --background: var(--al-bg);
            --border: var(--al-border);
        }

        [data-testid="stAppViewContainer"] {
            background: var(--al-bg);
        }

        [data-testid="stHeader"] {
            background: transparent;
        }

        [data-testid="stMainBlockContainer"] {
            max-width: 1440px;
            padding-top: 2rem;
            padding-bottom: 3rem;
            font-family: 'Poppins', sans-serif;
            color: var(--al-text);
        }

        section[data-testid="stMain"] p,
        section[data-testid="stMain"] label,
        section[data-testid="stMain"] h1,
        section[data-testid="stMain"] h2,
        section[data-testid="stMain"] h3,
        section[data-testid="stMain"] [data-testid="stCaptionContainer"] {
            color: var(--al-text) !important;
        }

        section[data-testid="stMain"] [data-testid="stCaptionContainer"] {
            color: var(--al-muted) !important;
        }

        h1 {
            font-weight: 600;
            color: var(--primary-blue);
            font-size: 2.2rem !important;
            margin-bottom: 0.5rem !important;
            letter-spacing: -0.5px;
        }

        h2, h3 {
            color: var(--text-primary);
            font-weight: 600;
        }

        [data-testid="stSidebar"] {
            background: var(--al-navy);
            border-right: 1px solid #0A1D33;
        }

        /* Keep dark-sidebar text and light interactive controls separate. */
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
            color: #F4F7FA !important;
        }

        [data-testid="stSidebar"] details {
            background: transparent !important;
            border: 1px solid rgba(255, 255, 255, 0.2) !important;
            border-radius: 6px !important;
        }

        [data-testid="stSidebar"] details > summary,
        [data-testid="stSidebar"] details[open] > summary,
        [data-testid="stSidebar"] details > summary:hover,
        [data-testid="stSidebar"] details > summary:focus-visible {
            background: #18395C !important;
            color: #F4F7FA !important;
            border-radius: 5px !important;
        }

        [data-testid="stSidebar"] details > summary *,
        [data-testid="stSidebar"] details[open] > summary * {
            color: #F4F7FA !important;
            fill: #F4F7FA !important;
        }

        [data-testid="stSidebar"] details > [data-testid="stExpanderDetails"] {
            background: var(--al-navy) !important;
            color: #F4F7FA !important;
        }

        [data-testid="stSidebar"] details > [data-testid="stExpanderDetails"] * {
            color: #F4F7FA !important;
        }

        [data-testid="stSidebar"] [data-baseweb="select"] > div {
            background: var(--al-surface) !important;
            border: 1px solid var(--al-border) !important;
            border-radius: 6px !important;
        }

        [data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"],
        [data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"] span,
        [data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"] input {
            color: var(--al-text) !important;
        }

        [data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"] svg,
        [data-testid="stSidebar"] [data-baseweb="select"] svg {
            fill: var(--al-text) !important;
            color: var(--al-text) !important;
        }

        [data-testid="stSidebar"] [data-baseweb="select"] [aria-disabled="true"],
        [data-testid="stSidebar"] [data-baseweb="select"] [aria-disabled="true"] span {
            background: #EEF2F5 !important;
            color: #5B6B7C !important;
            opacity: 1 !important;
        }

        [data-testid="stSidebar"] [role="combobox"] {
            background: var(--al-surface) !important;
            border: 1px solid var(--al-border) !important;
            border-radius: 6px !important;
            color: var(--al-text) !important;
            -webkit-text-fill-color: var(--al-text) !important;
        }

        [data-testid="stSidebar"] [role="combobox"]::placeholder {
            color: var(--al-muted) !important;
            -webkit-text-fill-color: var(--al-muted) !important;
        }

        [data-testid="stSidebar"] [role="combobox"]:disabled {
            background: #EEF2F5 !important;
            color: #5B6B7C !important;
            -webkit-text-fill-color: #5B6B7C !important;
            opacity: 1 !important;
        }

        [data-testid="stSidebar"] [role="combobox"] + button,
        [data-testid="stSidebar"] [role="combobox"] ~ button {
            background: var(--al-surface) !important;
            color: var(--al-text) !important;
            border: 1px solid var(--al-border) !important;
            border-left: 0 !important;
        }

        [data-testid="stSidebar"] [role="combobox"] + button svg,
        [data-testid="stSidebar"] [role="combobox"] ~ button svg {
            color: var(--al-text) !important;
            fill: var(--al-text) !important;
        }

        [role="listbox"][aria-label="Navigation Open"],
        [role="listbox"][aria-label="Select Sheet Open"] {
            background: var(--al-surface) !important;
            border: 1px solid var(--al-border) !important;
        }

        [role="listbox"][aria-label="Navigation Open"] [role="option"],
        [role="listbox"][aria-label="Select Sheet Open"] [role="option"] {
            background: var(--al-surface) !important;
            color: var(--al-text) !important;
        }

        [role="listbox"][aria-label="Navigation Open"] [role="option"]:hover,
        [role="listbox"][aria-label="Select Sheet Open"] [role="option"]:hover,
        [role="listbox"][aria-label="Navigation Open"] [aria-selected="true"],
        [role="listbox"][aria-label="Select Sheet Open"] [aria-selected="true"] {
            background: var(--light-blue) !important;
            color: var(--al-text) !important;
        }

        .al-brand {
            border-bottom: 1px solid rgba(255, 255, 255, 0.2);
            padding: 0.25rem 0 1.25rem;
            margin-bottom: 1rem;
        }

        .al-brand-name { font-size: 1.45rem; font-weight: 700; letter-spacing: 0.02em; }
        .al-brand-subtitle { color: #AFC4D6; font-size: 0.78rem; margin-top: 0.2rem; }

        .al-kpi {
            background: var(--al-surface);
            border: 1px solid var(--al-border);
            border-top: 4px solid var(--al-info);
            border-radius: 8px;
            box-shadow: 0 3px 12px rgba(15, 39, 68, 0.08);
            min-height: 96px;
            padding: 1rem 1.1rem;
            margin-bottom: 1.25rem;
        }

        .al-kpi-success { border-top-color: var(--al-success); }
        .al-kpi-warning { border-top-color: var(--al-warning); }
        .al-kpi-danger { border-top-color: var(--al-danger); }
        .al-kpi-muted { border-top-color: var(--al-muted); }
        .al-kpi-navy { border-top-color: var(--al-navy); }
        .al-kpi-teal { border-top-color: var(--al-teal); }
        .al-kpi-label { color: var(--al-muted); font-size: 0.74rem; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; }
        .al-kpi-value { color: var(--al-text); font-size: 1.8rem; font-weight: 700; line-height: 1.25; margin-top: 0.5rem; }

        [data-testid="stMarkdownContainer"] h2,
        [data-testid="stMarkdownContainer"] h3,
        .stSubheader {
            border-bottom: 1px solid var(--al-border);
            padding-bottom: 0.45rem;
            margin-top: 1.8rem;
        }

        .metric-card {
            background: #FFFFFF;
            border-radius: 8px;
            padding: 24px;
            color: #1A4D7A;
            text-align: center;
            font-weight: 500;
            margin-bottom: 16px;
            box-shadow: 0 2px 8px rgba(23, 50, 77, 0.08);
            transition: box-shadow 0.2s ease;
            border: 1px solid var(--border);
            border-top: 3px solid var(--secondary-blue);
        }

        .metric-card:hover {
            box-shadow: 0 4px 14px rgba(23, 50, 77, 0.14);
        }

        .metric-card h2 {
            font-size: 2.4rem;
            margin: 12px 0;
            color: #0066B3 !important;
            font-weight: 700;
        }

        .metric-card .metric-label {
            font-size: 0.85rem;
            color: #1A4D7A;
            text-transform: uppercase;
            letter-spacing: 0.8px;
            font-weight: 600;
        }

        .card-warning { background: linear-gradient(135deg, #FFF8E6 0%, #FFF0CC 100%); }
        .card-warning h2 { color: #B8860B !important; }
        .card-info { background: linear-gradient(135deg, #E6F7FC 0%, #CCF0FA 100%); }
        .card-info h2 { color: #007BA7 !important; }
        .card-danger { background: linear-gradient(135deg, #FFE8EB 0%, #FFD6DC 100%); }
        .card-danger h2 { color: #B91C2E !important; }

        .section-header {
            background: linear-gradient(135deg, var(--primary-blue) 0%, var(--secondary-blue) 100%);
            padding: 14px 20px;
            border-radius: 8px;
            color: white;
            font-weight: 600;
            font-size: 1.1rem;
            margin: 24px 0 16px 0;
            box-shadow: 0 2px 8px rgba(0, 102, 179, 0.2);
        }

        .stDataFrame {
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(15, 39, 68, 0.06);
            border: 1px solid var(--border);
        }

        .stButton>button, [data-testid="stDownloadButton"] button {
            background: var(--al-blue);
            color: white;
            border: none;
            border-radius: 6px;
            padding: 10px 20px;
            font-weight: 500;
            transition: all 0.3s ease;
            box-shadow: 0 2px 6px rgba(15, 39, 68, 0.2);
        }

        .stButton>button:hover, [data-testid="stDownloadButton"] button:hover {
            background: var(--al-teal);
            box-shadow: 0 4px 12px rgba(15, 39, 68, 0.3);
        }

        @media (max-width: 768px) {
            h1 { font-size: 1.6rem !important; }
            .metric-card h2 { font-size: 1.8rem; }
        }
        </style>
    """, unsafe_allow_html=True)

# ============================================================================
# DATA PROCESSING FUNCTIONS
# ============================================================================

# ============================================================================
# FILE OPERATIONS
# ============================================================================

# ============================================================================
if __name__ == '__main__':
    inject_professional_css()
    st.title("AssetLens")
    st.caption("IT Asset Audit & Lifecycle Intelligence")
    st.caption("Audit, analyze and plan from official IT asset inventory exports.")

    with st.sidebar:
        st.markdown(
            '<div class="al-brand"><div class="al-brand-name">AssetLens</div>'
            '<div class="al-brand-subtitle">Asset Intelligence</div></div>',
            unsafe_allow_html=True,
        )
        st.markdown("**Dataset**")

    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
    if uploaded_file is None:
        st.info("Upload an Excel export to get started.")
        st.markdown("The dashboard keeps source values intact while adding lifecycle, warranty and audit analysis in memory.")
        with st.expander("Help & Support", expanded=False):
            st.markdown("**Troubleshooting**\n\nUse an unrestricted `.xlsx` export with a recognizable asset header row. The required identity columns vary by asset type.")
            st.markdown("**Contact Support**\n\nEmail: khalis.abdrahim@gmail.com")
        st.sidebar.caption("Version 2.5.0")
        st.stop()

    try:
        uploaded_file.seek(0)
        file_bytes = uploaded_file.read()
        if not file_bytes.startswith(b'PK'):
            st.error("The uploaded file is not a valid Excel (.xlsx) file.")
            st.stop()

        uploaded_file.seek(0)
        workbook = pd.ExcelFile(uploaded_file, engine='openpyxl')
        selected_sheet = st.sidebar.selectbox("Select Sheet", workbook.sheet_names)
        uploaded_file.seek(0)
        detected_header = detect_header_row(uploaded_file, selected_sheet)
        use_manual_header = st.sidebar.checkbox("Manual Header Row Selection", value=False)
        if use_manual_header:
            header_row = st.sidebar.number_input(
                "Header Row (0-based)", min_value=0, max_value=20,
                value=detected_header if detected_header is not None else 0,
            )
        elif detected_header is None:
            st.warning("Automatic header detection could not find a confident export header. Enable manual header selection.")
            st.stop()
        else:
            header_row = detected_header

        uploaded_file.seek(0)
        source_df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row, engine='openpyxl')
        source_df.columns = [str(column).strip() for column in source_df.columns]
        source_df = source_df.loc[:, ~source_df.columns.duplicated(keep='first')]
        asset_type = detect_asset_type_from_data(source_df)
        if asset_type == "Unknown":
            st.error("Could not confidently detect this export. Expected Workstation Type or IT Smartphones / IT Tablets.")
            st.stop()

        missing_required = validate_source_columns(source_df, asset_type)
        if missing_required:
            st.error(f"This {asset_type.lower()} export is missing required columns: {', '.join(missing_required)}")
            st.stop()

        # Process once; all five pages consume this same audited DataFrame.
        processed_df = build_canonical_dataframe(source_df, asset_type)
        processed_df = calculate_asset_age(processed_df)
        processed_df, _ = get_warranty_status(processed_df)
        processed_df = run_itam_audit(processed_df)

        st.sidebar.success(f"Detected: {asset_type} assets")
        with st.sidebar.expander("Dataset details", expanded=False):
            st.write(f"**Rows:** {len(processed_df):,}")
            st.write(f"**Columns found:** {len(source_df.columns):,}")
            st.write(f"**Header row:** {header_row}")
            st.write("**Analysis fields:** Asset Type, Model / Product, Source State, Lifecycle, Warranty Status")
        with st.sidebar.expander("Help & Support", expanded=False):
            st.markdown("**Troubleshooting**\n\nCheck the selected sheet and header row if the export is not detected.")
            st.markdown("**Contact Support**\n\nEmail: khalis.abdrahim@gmail.com")
        st.sidebar.caption("Version 2.5.0")

    except Exception as error:
        st.error(f"Error reading Excel file: {error}")
        st.info("Check that the file is an unprotected .xlsx export with the correct header row.")
    else:
        render_navigation(processed_df, asset_type)
