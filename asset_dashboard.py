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
from itam.ui import render_navigation, render_selected_page, render_sidebar_dataset_info

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
# THEME DETECTION
# ============================================================================
def _active_theme_type():
    """Return 'dark' or 'light' for the user's active Streamlit theme (falls back to 'light')."""
    try:
        return st.context.theme.type or "light"
    except Exception:
        return "light"


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

            /* Fixed-light controls used inside the always-dark navy sidebar; kept theme-independent. */
            --al-control-surface: #FFFFFF;
            --al-control-border: #D9E2EC;
            --al-control-text: #172B4D;
            --al-control-muted: #6B7C93;
            --al-control-highlight: #E7F0F5;
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
            background: var(--al-surface);
            border-right: 1px solid var(--al-border);
        }

        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
            color: var(--al-text) !important;
        }

        [data-testid="stSidebar"] [data-testid="stButton"] button { background: transparent; border: 1px solid transparent; box-shadow: none; color: var(--al-text); justify-content: flex-start; margin: 0.08rem 0; padding: 0.35rem 0.55rem; }
        /* The sidebar contains only the five navigation buttons in this shell. */
        [data-testid="stSidebar"] [data-testid="stButton"] button,
        [data-testid="stSidebar"] [data-testid="stButton"] button > div { justify-content: flex-start !important; text-align: left !important; }
        [data-testid="stSidebar"] [data-testid="stButton"] button:hover { background: var(--al-light-blue); border-color: var(--al-border); color: var(--al-text); }
        [data-testid="stSidebar"] [data-testid="stButton"] button:focus-visible { border-color: var(--al-info); box-shadow: 0 0 0 2px color-mix(in srgb, var(--al-info) 28%, transparent); }
        [data-testid="stSidebar"] [data-testid="stButton"] button[kind="primary"] { background: var(--al-light-blue); border-color: var(--al-info); color: var(--al-info); font-weight: 700; }
        .al-nav-section-label { color: var(--al-muted); font-size: 0.65rem; font-weight: 700; letter-spacing: 0.1em; margin: 0.65rem 0 0.15rem; text-transform: uppercase; }

        [data-testid="stSidebar"] details {
            background: transparent !important;
            border: 1px solid rgba(255, 255, 255, 0.2) !important;
            border-radius: 6px !important;
        }

        [data-testid="stSidebar"] details > summary,
        [data-testid="stSidebar"] details[open] > summary,
        [data-testid="stSidebar"] details > summary:hover,
        [data-testid="stSidebar"] details > summary:focus-visible {
            background: var(--al-light-blue) !important;
            color: var(--al-text) !important;
            border-radius: 5px !important;
        }

        [data-testid="stSidebar"] details > summary *,
        [data-testid="stSidebar"] details[open] > summary * {
            color: var(--al-text) !important;
            fill: var(--al-text) !important;
        }

        [data-testid="stSidebar"] details > [data-testid="stExpanderDetails"] {
            background: var(--al-surface) !important;
            color: var(--al-text) !important;
        }

        [data-testid="stSidebar"] details > [data-testid="stExpanderDetails"] * {
            color: var(--al-text) !important;
        }

        [data-testid="stSidebar"] [data-baseweb="select"] > div {
            background: var(--al-control-surface) !important;
            border: 1px solid var(--al-control-border) !important;
            border-radius: 6px !important;
        }

        [data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"],
        [data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"] span,
        [data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"] input {
            color: var(--al-control-text) !important;
        }

        [data-testid="stSidebar"] [data-baseweb="select"] [role="combobox"] svg,
        [data-testid="stSidebar"] [data-baseweb="select"] svg {
            fill: var(--al-control-text) !important;
            color: var(--al-control-text) !important;
        }

        [data-testid="stSidebar"] [data-baseweb="select"] [aria-disabled="true"],
        [data-testid="stSidebar"] [data-baseweb="select"] [aria-disabled="true"] span {
            background: #EEF2F5 !important;
            color: #5B6B7C !important;
            opacity: 1 !important;
        }

        [data-testid="stSidebar"] [role="combobox"] {
            background: var(--al-control-surface) !important;
            border: 1px solid var(--al-control-border) !important;
            border-radius: 6px !important;
            color: var(--al-control-text) !important;
            -webkit-text-fill-color: var(--al-control-text) !important;
        }

        [data-testid="stSidebar"] [role="combobox"]::placeholder {
            color: var(--al-control-muted) !important;
            -webkit-text-fill-color: var(--al-control-muted) !important;
        }

        [data-testid="stSidebar"] [role="combobox"]:disabled {
            background: #EEF2F5 !important;
            color: #5B6B7C !important;
            -webkit-text-fill-color: #5B6B7C !important;
            opacity: 1 !important;
        }

        [data-testid="stSidebar"] [role="combobox"] + button,
        [data-testid="stSidebar"] [role="combobox"] ~ button {
            background: var(--al-control-surface) !important;
            color: var(--al-control-text) !important;
            border: 1px solid var(--al-control-border) !important;
            border-left: 0 !important;
        }

        [data-testid="stSidebar"] [role="combobox"] + button svg,
        [data-testid="stSidebar"] [role="combobox"] ~ button svg {
            color: var(--al-control-text) !important;
            fill: var(--al-control-text) !important;
        }

        [role="listbox"][aria-label="Navigation Open"],
        [role="listbox"][aria-label="Select Sheet Open"] {
            background: var(--al-control-surface) !important;
            border: 1px solid var(--al-control-border) !important;
        }

        [role="listbox"][aria-label="Navigation Open"] [role="option"],
        [role="listbox"][aria-label="Select Sheet Open"] [role="option"] {
            background: var(--al-control-surface) !important;
            color: var(--al-control-text) !important;
        }

        [role="listbox"][aria-label="Navigation Open"] [role="option"]:hover,
        [role="listbox"][aria-label="Select Sheet Open"] [role="option"]:hover,
        [role="listbox"][aria-label="Navigation Open"] [aria-selected="true"],
        [role="listbox"][aria-label="Select Sheet Open"] [aria-selected="true"] {
            background: var(--al-control-highlight) !important;
            color: var(--al-control-text) !important;
        }

        .al-brand {
            border-bottom: 1px solid var(--al-border);
            padding: 0.1rem 0 0.85rem;
            margin-bottom: 0.85rem;
        }

        .al-brand-name { color: var(--al-text); font-size: 1.35rem; font-weight: 700; letter-spacing: 0.01em; }
        .al-brand-subtitle { color: var(--al-muted); font-size: 0.68rem; line-height: 1.35; margin-top: 0.18rem; }
        .al-sidebar-section-label { color: var(--al-muted); font-size: 0.68rem; font-weight: 700; letter-spacing: 0.08em; margin: 1rem 0 0.35rem; text-transform: uppercase; }
        .al-dataset-panel, .al-dataset-empty { background: var(--al-light-blue); border: 1px solid var(--al-border); border-radius: 7px; padding: 0.65rem 0.7rem; }
        .al-dataset-empty { color: var(--al-muted); font-size: 0.78rem; }
        .al-dataset-row { margin: 0 0 0.48rem; overflow: hidden; }
        .al-dataset-row:last-child { margin-bottom: 0; }
        .al-dataset-row span { color: var(--al-muted); display: block; font-size: 0.65rem; }
        .al-dataset-row strong { color: var(--al-text); display: block; font-size: 0.77rem; overflow-wrap: anywhere; }
        .al-source-card { background: var(--al-surface); border: 1px solid var(--al-border); border-radius: 8px; box-shadow: 0 3px 12px rgba(15, 39, 68, 0.06); margin: 0.5rem 0 1.5rem; padding: 1rem 1.1rem 0.8rem; }
        .al-source-card h3 { border: 0; color: var(--al-text); margin: 0 0 0.15rem; padding: 0; }
        .al-source-card p { color: var(--al-muted) !important; font-size: 0.82rem; margin: 0 0 0.6rem; }

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
            background: var(--al-surface);
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

    if _active_theme_type() == "dark":
        st.markdown("""
            <style>
            :root {
                --al-bg: #0B1622;
                --al-surface: #16212E;
                --al-border: #2B3A4A;
                --al-text: #E7ECF2;
                --al-muted: #93A4B5;
                --primary-blue: #6FB7DE;
                --secondary-blue: #4FB4D8;
                --light-blue: #1E3A4D;
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
    st.sidebar.markdown(
        '<div class="al-brand"><div class="al-brand-name">AssetLens</div>'
        '<div class="al-brand-subtitle">IT Asset Audit &amp; Lifecycle Intelligence</div></div>',
        unsafe_allow_html=True,
    )
    navigation_slot = st.sidebar.empty()
    dataset_slot = st.sidebar.empty()
    import_slot = st.sidebar.empty()
    st.title("AssetLens")
    st.caption("IT Asset Audit & Lifecycle Intelligence")
    with st.container(border=True):
        st.subheader("Dataset Source")
        st.caption("Upload an official asset master export to begin analysis.")
        uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])
    if uploaded_file is None:
        st.caption("Source values remain unchanged; analysis is performed in memory.")
        render_sidebar_dataset_info(container=dataset_slot)
        import_slot.empty()
        with st.sidebar.expander("Help & Support", expanded=False):
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
        selected_sheet = workbook.sheet_names[0]
        uploaded_file.seek(0)
        detected_header = detect_header_row(uploaded_file, selected_sheet)
        with import_slot.container():
            st.markdown('<div class="al-sidebar-section-label">Import</div>', unsafe_allow_html=True)
            with st.expander("Import Options", expanded=detected_header is None):
                if len(workbook.sheet_names) > 1:
                    selected_sheet = st.selectbox("Select Sheet", workbook.sheet_names)
                    uploaded_file.seek(0)
                    detected_header = detect_header_row(uploaded_file, selected_sheet)
                use_manual_header = st.checkbox("Manual Header Row Selection", value=False)
                if use_manual_header:
                    header_row = st.number_input(
                        "Header Row (0-based)", min_value=0, max_value=20,
                        value=detected_header if detected_header is not None else 0,
                    )
        if not use_manual_header and detected_header is None:
            st.warning("Automatic header detection could not find a confident export header. Use Import Options to select the header row.")
            st.stop()
        elif not use_manual_header:
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

        navigation_slot.empty()
        render_navigation(processed_df, asset_type, navigation_slot)
        render_sidebar_dataset_info(asset_type, uploaded_file.name, len(processed_df), dataset_slot)
        with st.sidebar.expander("Help & Support", expanded=False):
            st.markdown("**Troubleshooting**\n\nCheck the selected sheet and header row if the export is not detected.")
            st.markdown("**Contact Support**\n\nEmail: khalis.abdrahim@gmail.com")
        st.sidebar.caption("Version 2.5.0")

    except Exception as error:
        st.error(f"Error reading Excel file: {error}")
        st.info("Check that the file is an unprotected .xlsx export with the correct header row.")
    else:
        render_selected_page(processed_df, asset_type)
