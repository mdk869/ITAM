import streamlit as st
import pandas as pd
import re
from html import escape
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go

from itam.audit import calculate_asset_age as modular_calculate_asset_age
from itam.audit import get_warranty_status as modular_get_warranty_status
from itam.audit import run_itam_audit as modular_run_itam_audit
from itam.data import (
    CANONICAL_COLUMNS as MODULAR_CANONICAL_COLUMNS,
    CANONICAL_REQUIRED_COLUMNS as MODULAR_CANONICAL_REQUIRED_COLUMNS,
    CANONICAL_SOURCE_MAP as MODULAR_CANONICAL_SOURCE_MAP,
    apply_literal_search as modular_apply_literal_search,
    build_canonical_dataframe as modular_build_canonical_dataframe,
    detect_asset_type as modular_detect_asset_type,
    detect_asset_type_from_data as modular_detect_asset_type_from_data,
    detect_header_row as modular_detect_header_row,
    find_column as modular_find_column,
    get_model_column as modular_get_model_column,
    get_type_column as modular_get_type_column,
    normalize_text as modular_normalize_text,
    validate_source_columns as modular_validate_source_columns,
)
from itam.export import export_to_excel as modular_export_to_excel
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

        .card-primary { background: linear-gradient(135deg, #E8F4FC 0%, #D6EDFA 100%); }
        .card-primary h2 { color: #0066B3 !important; }
        .card-success { background: linear-gradient(135deg, #E8F8F0 0%, #D1F2E0 100%); }
        .card-success h2 { color: #1B7A4C !important; }
        .card-warning { background: linear-gradient(135deg, #FFF8E6 0%, #FFF0CC 100%); }
        .card-warning h2 { color: #B8860B !important; }
        .card-info { background: linear-gradient(135deg, #E6F7FC 0%, #CCF0FA 100%); }
        .card-info h2 { color: #007BA7 !important; }
        .card-danger { background: linear-gradient(135deg, #FFE8EB 0%, #FFD6DC 100%); }
        .card-danger h2 { color: #B91C2E !important; }

        .type-card {
            background: linear-gradient(135deg, #E8F4FC 0%, #D6EDFA 100%);
            border-radius: 12px;
            padding: 24px;
            color: #1A4D7A;
            text-align: center;
            font-weight: 500;
            box-shadow: 0 2px 12px rgba(0, 102, 179, 0.1);
            transition: all 0.3s ease;
            border: none;
            min-height: 120px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        }

        .type-card:hover {
            box-shadow: 0 4px 20px rgba(0, 102, 179, 0.18);
            transform: translateY(-3px);
        }

        .type-card .type-label {
            font-size: 0.9rem;
            margin-bottom: 8px;
            font-weight: 600;
            color: #1A4D7A;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }

        .type-card .type-count {
            font-size: 2rem;
            font-weight: 700;
            color: #0066B3;
        }

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

        .sidebar-section {
            background: var(--light-blue);
            padding: 10px 16px;
            border-radius: 6px;
            color: var(--primary-blue);
            font-weight: 600;
            font-size: 0.95rem;
            margin: 16px 0 12px 0;
            border-left: 3px solid var(--primary-blue);
        }

        .severity-badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 4px;
            font-weight: 600;
            font-size: 0.75rem;
            letter-spacing: 0.5px;
            margin-right: 8px;
        }

        .severity-high { background: #DC3545; color: white; }
        .severity-medium { background: #FFC107; color: #2C3E50; }
        .severity-low { background: #28A745; color: white; }

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
# UTILITY FUNCTIONS
# ============================================================================

@st.cache_data
def normalize_text(text):
    """Normalize text for column matching"""
    return re.sub(r'[^a-z0-9]', '', str(text).lower())

def find_column(df, search_terms):
    """Find column by multiple search terms"""
    if isinstance(search_terms, str):
        search_terms = [search_terms]

    normalized_cols = {normalize_text(col): col for col in df.columns}

    for term in search_terms:
        normalized_term = normalize_text(term)
        for norm_col, orig_col in normalized_cols.items():
            if normalized_term in norm_col:
                return orig_col
    return None

# ============================================================================
# ASSET TYPE DETECTION
# ============================================================================

def detect_asset_type(df_columns):
    """Detect an asset type from the company's export-specific columns."""
    normalized_columns = {normalize_text(column) for column in df_columns}
    if "workstationtype" in normalized_columns:
        return "Workstation"
    return "Unknown"


def detect_asset_type_from_data(df):
    """Detect smartphone or tablet exports from exact Product Type values."""
    normalized_columns = {
        normalize_text(column): column for column in df.columns
    }
    if "workstationtype" in normalized_columns:
        return "Workstation"

    product_type_col = normalized_columns.get("producttype")
    if not product_type_col:
        return "Unknown"

    values = {
        str(value).strip().casefold()
        for value in df[product_type_col].dropna()
        if str(value).strip()
    }
    if values and values.issubset({"it smartphones"}):
        return "Smartphone"
    if values and values.issubset({"it tablets"}):
        return "Tablet"
    return "Unknown"

CANONICAL_COLUMNS = [
    "asset_type", "serial_number", "model", "asset_tag", "state", "user",
    "employee_id", "email", "job_title", "department", "location", "site",
    "purchase_year", "warranty_expiry", "programme", "place",
    "workstation_status", "imei", "sim_number",
]

CANONICAL_SOURCE_MAP = {
    "Workstation": {
        "asset_type": "Workstation Type", "serial_number": "Serial Number",
        "model": "Model", "asset_tag": "Asset Tag", "state": "State",
        "user": "User", "employee_id": "User Employee ID", "email": "User Email",
        "job_title": "User Jobtitle", "department": "Department", "location": "Location",
        "site": "Site", "purchase_year": "Year Of Purchase",
        "warranty_expiry": "Warranty Expiry", "programme": "Programme",
        "place": "Place", "workstation_status": "Workstation Status",
    },
    "Smartphone": {
        "asset_type": "Product Type", "serial_number": "Serial Number", "model": "Product",
        "asset_tag": "AssetTag", "state": "State", "user": "User",
        "employee_id": "User -> Employee ID", "email": "User -> Email",
        "job_title": "User -> Job Title", "department": "User -> Department",
        "location": "Location", "site": "User -> Site", "purchase_year": "Year Of Purchase",
        "warranty_expiry": "Warranty Expiry Date", "programme": "Programme",
        "place": "Place", "imei": "No. IMEI", "sim_number": "No. Sim",
    },
    "Tablet": {
        "asset_type": "Product Type", "serial_number": "Serial Number", "model": "Product",
        "asset_tag": "AssetTag", "state": "State", "user": "User",
        "employee_id": "User -> Employee ID", "email": "User -> Email",
        "job_title": "User -> Job Title", "department": "User -> Department",
        "location": "Location", "site": "User -> Site", "purchase_year": "Year Of Purchase",
        "warranty_expiry": "Warranty Expiry Date", "programme": "Programme",
        "place": "Place", "imei": "No. IMEI", "sim_number": "No. Sim",
    },
}

CANONICAL_REQUIRED_COLUMNS = {
    "Workstation": ["Workstation Type", "Serial Number", "Model", "Asset Tag", "State", "Year Of Purchase"],
    "Smartphone": ["Product Type", "Serial Number", "Product", "AssetTag", "State", "Year Of Purchase"],
    "Tablet": ["Product Type", "Serial Number", "Product", "AssetTag", "State", "Year Of Purchase"],
}


def validate_source_columns(df, asset_type):
    """Return required source headers that are absent from an export."""
    normalized_columns = {normalize_text(column) for column in df.columns}
    return [
        column for column in CANONICAL_REQUIRED_COLUMNS.get(asset_type, [])
        if normalize_text(column) not in normalized_columns
    ]


def build_canonical_dataframe(df, asset_type):
    """Copy source values into the canonical ITAM schema without dropping raw columns."""
    mapping = CANONICAL_SOURCE_MAP[asset_type]
    source_columns = {normalize_text(column): column for column in df.columns}
    canonical_df = df.copy()

    for canonical_column in CANONICAL_COLUMNS:
        source_column = mapping.get(canonical_column)
        source_name = source_columns.get(normalize_text(source_column)) if source_column else None
        canonical_df[canonical_column] = df[source_name] if source_name else pd.NA

    canonical_df["asset_type"] = asset_type
    source_type_name = source_columns.get(normalize_text(mapping["asset_type"]))
    if source_type_name:
        canonical_df["source_asset_subtype"] = df[source_type_name]
    return canonical_df


def get_model_column(df, asset_type):
    """Get model column based on asset type"""
    return "model" if "model" in df.columns else None

def get_type_column(df, asset_type):
    """Get type column based on asset type"""
    return "source_asset_subtype" if "source_asset_subtype" in df.columns else None

def apply_literal_search(df, query, columns=None):
    """Filter rows using a case-insensitive literal search."""
    if not query:
        return df.copy()
    search_df = df if columns is None else df.loc[:, columns]
    matches = search_df.apply(
        lambda row: row.astype(str).str.contains(
            query, case=False, regex=False, na=False
        ).any(),
        axis=1,
    )
    return df.loc[matches].copy()

def detect_header_row(excel_file, sheet_name):
    """Find a confident company-export header row in the first 20 rows."""
    try:
        preview = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            header=None,
            nrows=20,
            engine="openpyxl",
        )

        identity_columns = {"serialnumber", "assettag", "model", "product", "yearofpurchase"}
        type_columns = {"workstationtype", "producttype"}
        candidates = []

        for i, row in preview.iterrows():
            normalized_values = {
                normalize_text(value)
                for value in row.dropna()
                if normalize_text(value)
            }
            matched_identity = normalized_values & identity_columns
            matched_type = normalized_values & type_columns

            # Require a type column plus most core identity columns. This
            # avoids treating report metadata rows as the table header.
            if matched_type and len(matched_identity) >= 3:
                score = len(matched_identity) * 2 + len(matched_type)
                candidates.append((score, i))

        if candidates:
            return max(candidates, key=lambda candidate: (candidate[0], -candidate[1]))[1]
        return None
    except Exception:
        return None

# ============================================================================
# DATA PROCESSING FUNCTIONS
# ============================================================================

@st.cache_data
def calculate_asset_age(df):
    """Calculate nullable asset age and the authoritative ITAM lifecycle."""
    df = df.copy()
    year_col = "purchase_year" if "purchase_year" in df.columns else find_column(df, ["year of purchase", "yearofpurchase"])
    if year_col:
        current_year = pd.Timestamp.now().year
        purchase_year = pd.to_numeric(df[year_col], errors="coerce")
        purchase_year = purchase_year.where(purchase_year.mod(1).eq(0))
        asset_age = current_year - purchase_year
    else:
        asset_age = pd.Series(pd.NA, index=df.index, dtype="Float64")

    df["Asset Age"] = asset_age.astype("Int64")
    df["ITAM Lifecycle Status"] = "Unknown"
    df.loc[df["Asset Age"].between(0, 1), "ITAM Lifecycle Status"] = "New"
    df.loc[df["Asset Age"].between(2, 3), "ITAM Lifecycle Status"] = "Active"
    df.loc[df["Asset Age"].between(4, 5), "ITAM Lifecycle Status"] = "Aging"
    df.loc[df["Asset Age"] > 5, "ITAM Lifecycle Status"] = "Expired"
    return df

@st.cache_data
def get_warranty_status(df):
    """Calculate warranty status"""
    warranty_col = "warranty_expiry" if "warranty_expiry" in df.columns else find_column(df, ["warranty expiry", "warrantyexpiry"])
    if not warranty_col:
        return df, None

    df_temp = df.copy()
    df_temp["Warranty Expiry Date"] = pd.to_datetime(df_temp[warranty_col], errors='coerce')
    today = pd.Timestamp.now().normalize()
    df_temp["Days to Expiry"] = (df_temp["Warranty Expiry Date"] - today).dt.days

    df_temp["Warranty Status"] = "Unknown"
    df_temp.loc[df_temp["Days to Expiry"] < 0, "Warranty Status"] = "Expired"
    df_temp.loc[df_temp["Days to Expiry"].between(0, 90), "Warranty Status"] = "Expiring Soon"
    df_temp.loc[df_temp["Days to Expiry"] > 90, "Warranty Status"] = "Active"

    expired_warranty_df = df_temp[df_temp["Warranty Status"] == "Expired"].copy()
    return df_temp, expired_warranty_df


def run_itam_audit(df):
    """Add row-level ITAM audit findings without changing source values."""
    audited_df = df.copy()
    row_findings = [[] for _ in audited_df.index]
    row_positions = pd.Series(range(len(audited_df)), index=audited_df.index)

    severity_priority = {
        "Critical": 5, "High": 4, "Medium": 3,
        "Low": 2, "Info": 1, "None": 0,
    }

    def is_missing(value):
        if pd.isna(value):
            return True
        return str(value).strip() == "" or str(value).strip().casefold() == "na"

    def normalized_identifier(column):
        if column not in audited_df.columns:
            return pd.Series(pd.NA, index=audited_df.index, dtype="string")
        normalized = audited_df[column].astype("string").str.strip().str.casefold()
        return normalized.where(normalized.notna() & normalized.ne("") & normalized.ne("na"))

    def add_finding(position, rule, severity, message):
        row_findings[position].append({
            "rule": rule,
            "severity": severity,
            "message": message,
        })

    def add_duplicate_flags(column, flag_column, rule, label):
        normalized = normalized_identifier(column)
        duplicate_mask = normalized.notna() & normalized.duplicated(keep=False)
        audited_df[flag_column] = duplicate_mask.to_numpy()
        for index in audited_df.index[duplicate_mask]:
            position = row_positions.loc[index]
            add_finding(position, rule, "High", f"Duplicate {label}.")

    add_duplicate_flags(
        "asset_tag", "ITAM Duplicate Asset Tag",
        "DUPLICATE_ASSET_TAG", "Asset Tag"
    )
    add_duplicate_flags(
        "serial_number", "ITAM Duplicate Serial",
        "DUPLICATE_SERIAL", "Serial Number"
    )

    asset_types = audited_df.get("asset_type", pd.Series(index=audited_df.index, dtype="object"))
    imei_normalized = normalized_identifier("imei")
    imei_applicable = asset_types.astype("string").str.casefold().isin({"smartphone", "tablet"})
    duplicate_imei = imei_applicable & imei_normalized.notna() & imei_normalized.duplicated(keep=False)
    audited_df["ITAM Duplicate IMEI"] = duplicate_imei.to_numpy()
    for index in audited_df.index[duplicate_imei]:
        position = row_positions.loc[index]
        add_finding(position, "DUPLICATE_IMEI", "High", "Duplicate IMEI.")

    core_columns = [
        ("asset_tag", "asset_tag"),
        ("serial_number", "serial_number"),
        ("model", "model"),
        ("purchase_year", "purchase_year"),
    ]
    missing_identity = []
    for index, row in audited_df.iterrows():
        missing_fields = [
            label for column, label in core_columns
            if column not in audited_df.columns or is_missing(row[column])
        ]
        missing_identity.append(bool(missing_fields))
        if missing_fields:
            position = row_positions.loc[index]
            add_finding(
                position,
                "MISSING_CORE_IDENTITY",
                "Medium",
                "Missing core identity: " + ", ".join(missing_fields) + ".",
            )
    audited_df["ITAM Missing Core Identity"] = missing_identity

    active_states = {"in use", "in store", "in repair"}
    state_values = audited_df.get("state", pd.Series(index=audited_df.index, dtype="object"))
    lifecycle_values = audited_df.get(
        "ITAM Lifecycle Status", pd.Series(index=audited_df.index, dtype="object")
    )
    state_mismatch = (
        lifecycle_values.astype("string").eq("Expired")
        & state_values.astype("string").str.strip().str.casefold().isin(active_states)
    )
    audited_df["ITAM State Review Required"] = state_mismatch.to_numpy()
    for index in audited_df.index[state_mismatch]:
        position = row_positions.loc[index]
        source_state = audited_df.at[index, "state"]
        add_finding(
            position,
            "LIFECYCLE_STATE_MISMATCH",
            "Medium",
            f"ITAM lifecycle is Expired but source State is {source_state}.",
        )

    warranty_status = audited_df.get(
        "Warranty Status", pd.Series(index=audited_df.index, dtype="object")
    ).astype("string")
    for index in audited_df.index[warranty_status.eq("Expired")]:
        add_finding(row_positions.loc[index], "EXPIRED_WARRANTY", "Low", "Warranty is expired.")
    for index in audited_df.index[warranty_status.eq("Expiring Soon")]:
        add_finding(row_positions.loc[index], "EXPIRING_WARRANTY", "Info", "Warranty is expiring soon.")

    replacement_candidate = lifecycle_values.astype("string").eq("Expired")
    audited_df["ITAM Replacement Candidate"] = replacement_candidate.to_numpy()

    finding_counts = []
    highest_severities = []
    review_required = []
    finding_text = []
    review_rules = {
        "DUPLICATE_ASSET_TAG", "DUPLICATE_SERIAL", "DUPLICATE_IMEI",
        "MISSING_CORE_IDENTITY", "LIFECYCLE_STATE_MISMATCH",
    }
    for findings in row_findings:
        finding_counts.append(len(findings))
        highest_severities.append(
            max((finding["severity"] for finding in findings), key=severity_priority.get, default="None")
        )
        review_required.append(any(finding["rule"] in review_rules for finding in findings))
        finding_text.append("; ".join(
            f"[{finding['severity']}] {finding['message'].rstrip('.') }"
            for finding in findings
        ))

    audited_df["ITAM Finding Count"] = finding_counts
    audited_df["ITAM Highest Severity"] = highest_severities
    audited_df["ITAM Review Required"] = review_required
    audited_df["ITAM Audit Findings"] = finding_text
    return audited_df


normalize_text = modular_normalize_text
find_column = modular_find_column
detect_asset_type = modular_detect_asset_type
detect_asset_type_from_data = modular_detect_asset_type_from_data
CANONICAL_COLUMNS = MODULAR_CANONICAL_COLUMNS
CANONICAL_SOURCE_MAP = MODULAR_CANONICAL_SOURCE_MAP
CANONICAL_REQUIRED_COLUMNS = MODULAR_CANONICAL_REQUIRED_COLUMNS
validate_source_columns = modular_validate_source_columns
build_canonical_dataframe = modular_build_canonical_dataframe
get_model_column = modular_get_model_column
get_type_column = modular_get_type_column
apply_literal_search = modular_apply_literal_search
detect_header_row = modular_detect_header_row
calculate_asset_age = modular_calculate_asset_age
get_warranty_status = modular_get_warranty_status
run_itam_audit = modular_run_itam_audit

# ============================================================================
# DATA VALIDATION
# ============================================================================

def validate_data(df, asset_type, model_col):
    """Validate data and return list of issues"""
    issues = []

    asset_tag_col = "asset_tag" if "asset_tag" in df.columns else find_column(df, ["asset tag", "assettag"])
    serial_col = "serial_number" if "serial_number" in df.columns else find_column(df, ["serial number", "serialnumber"])
    user_col = "user" if "user" in df.columns else find_column(df, ["user"])
    email_col = "email" if "email" in df.columns else find_column(df, ["email"])
    dept_col = "department" if "department" in df.columns else find_column(df, ["department", "user department"])
    location_col = "location" if "location" in df.columns else find_column(df, ["location"])

    def normalized_identifier(series):
        normalized = series.astype("string").str.strip().str.casefold()
        return normalized.where(normalized.notna() & normalized.ne(""))

    # Preserve displayed identifiers, but normalize only the comparison values.
    if asset_tag_col:
        comparison = normalized_identifier(df[asset_tag_col])
        duplicates = df[comparison.notna() & comparison.duplicated(keep=False)]
        if not duplicates.empty:
            display_cols = [c for c in [asset_tag_col, model_col, serial_col, user_col] if c]
            issues.append({
                "type": "Duplicate Asset Tags",
                "count": comparison[comparison.duplicated(keep=False)].nunique(),
                "details": f"Found {comparison[comparison.duplicated(keep=False)].nunique()} duplicate asset tags",
                "severity": "high",
                "data": duplicates[display_cols].sort_values(asset_tag_col)
            })

    if serial_col:
        comparison = normalized_identifier(df[serial_col])
        duplicates = df[comparison.notna() & comparison.duplicated(keep=False)]
        if not duplicates.empty:
            display_cols = [c for c in [serial_col, model_col, asset_tag_col, user_col] if c]
            issues.append({
                "type": "Duplicate Serial Numbers",
                "count": comparison[comparison.duplicated(keep=False)].nunique(),
                "details": f"Found {comparison[comparison.duplicated(keep=False)].nunique()} duplicate serial numbers",
                "severity": "high",
                "data": duplicates[display_cols].sort_values(serial_col)
            })

    # Check missing data
    if user_col:
        missing_users = df[df[user_col].isna() | (df[user_col] == "")]
        if not missing_users.empty:
            display_cols = [c for c in [asset_tag_col, model_col, serial_col, dept_col] if c]
            issues.append({
                "type": "Missing User Assignment",
                "count": len(missing_users),
                "details": f"{len(missing_users)} assets without assigned users",
                "severity": "medium",
                "data": missing_users[display_cols]
            })

    # Check invalid emails
    if email_col:
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        invalid_emails = df[df[email_col].notna() & ~df[email_col].astype(str).str.match(email_pattern)]
        if not invalid_emails.empty:
            display_cols = [c for c in [user_col, email_col, asset_tag_col, model_col] if c]
            issues.append({
                "type": "Invalid Email Format",
                "count": len(invalid_emails),
                "details": f"{len(invalid_emails)} invalid email addresses",
                "severity": "low",
                "data": invalid_emails[display_cols]
            })

    return issues

def show_validation_issues(issues):
    """Display validation issues"""
    if not issues:
        st.success("No data validation issues found")
        return

    st.warning(f"Found {len(issues)} validation issue(s)")

    for issue in issues:
        severity_class = f"severity-{issue['severity']}"
        severity_label = issue['severity'].upper()

        with st.expander(f"{issue['type']} ({issue['count']})", expanded=False):
            st.markdown(f'<span class="severity-badge {severity_class}">{severity_label}</span> {issue["details"]}',
                       unsafe_allow_html=True)
            if "data" in issue and not issue["data"].empty:
                st.dataframe(issue["data"], use_container_width=True, hide_index=True)

# ============================================================================
# DISPLAY FUNCTIONS
# ============================================================================

def show_summary_cards(df, df_expired=None):
    """Display summary metric cards"""
    total_assets = len(df)
    expired_assets = len(df_expired) if df_expired is not None else 0
    within_lifecycle = (df["Asset Age"].notna() & (df["Asset Age"] <= 5)).sum() if "Asset Age" in df else 0
    replacement_rate = (expired_assets / total_assets * 100) if total_assets > 0 else 0

    col1, col2, col3, col4 = st.columns(4)

    cards = [
        (col1, "TOTAL ASSETS", total_assets, "card-primary"),
        (col2, "WITHIN LIFECYCLE", within_lifecycle, "card-success"),
        (col3, "EXPIRED ASSETS", expired_assets, "card-warning"),
        (col4, "REPLACEMENT RATE", f"{replacement_rate:.1f}%", "card-info")
    ]

    for col, label, value, card_class in cards:
        with col:
            st.markdown(f"""
                <div class="metric-card {card_class}">
                    <div class="metric-label">{label}</div>
                    <h2>{value}</h2>
                </div>
            """, unsafe_allow_html=True)

def show_type_cards(df, type_col, asset_type):
    """Display asset type cards"""
    if not type_col:
        return

    st.markdown(f'<div class="section-header">{escape(str(type_col))} Statistics</div>', unsafe_allow_html=True)

    type_counts = df[type_col].value_counts().sort_values(ascending=False)
    cols_per_row = min(4, len(type_counts))
    cols = st.columns(cols_per_row)

    for idx, (wtype, count) in enumerate(type_counts.items()):
        with cols[idx % cols_per_row]:
            st.markdown(f"""
                <div class="type-card card-primary">
                    <div class="type-label">{escape(str(wtype))}</div>
                    <div class="type-count">{count}</div>
                </div>
            """, unsafe_allow_html=True)

        if (idx + 1) % cols_per_row == 0 and (idx + 1) < len(type_counts):
            cols = st.columns(cols_per_row)

def show_warranty_summary(df, model_col):
    """Display warranty status summary"""
    if "Warranty Status" not in df.columns:
        return

    status_counts = df["Warranty Status"].value_counts()
    col1, col2, col3 = st.columns(3)

    statuses = [
        (col1, "EXPIRED WARRANTY", status_counts.get("Expired", 0), "card-danger"),
        (col2, "EXPIRING SOON (90 DAYS)", status_counts.get("Expiring Soon", 0), "card-warning"),
        (col3, "ACTIVE WARRANTY", status_counts.get("Active", 0), "card-success")
    ]

    for col, label, count, card_class in statuses:
        with col:
            st.markdown(f"""
                <div class="metric-card {card_class}">
                    <div class="metric-label">{label}</div>
                    <h2>{count}</h2>
                </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    serial_col = find_column(df, ["serial number", "serialnumber"])
    user_col = find_column(df, ["user"])
    dept_col = find_column(df, ["department", "user department"])
    location_col = find_column(df, ["location"])
    warranty_col = find_column(df, ["warranty expiry", "warrantyexpiry"])

    for status, label in [("Expired", "Expired Warranty Assets"),
                          ("Expiring Soon", "Expiring Soon Assets"),
                          ("Active", "Active Warranty Assets")]:
        status_df = df[df["Warranty Status"] == status]
        if not status_df.empty:
            with st.expander(f"{label} ({len(status_df)})", expanded=False):
                display_cols = [c for c in [model_col, serial_col, user_col, dept_col, location_col, warranty_col] if c]
                st.dataframe(status_df[display_cols], use_container_width=True, hide_index=True)

def show_asset_age_summary(df):
    """Display asset age analysis"""
    if "Asset Age" not in df.columns or df["Asset Age"].sum() == 0:
        return

    df_temp = df.copy()
    conditions = [
        df_temp["Asset Age"] <= 1,
        (df_temp["Asset Age"] > 1) & (df_temp["Asset Age"] <= 3),
        (df_temp["Asset Age"] > 3) & (df_temp["Asset Age"] <= 5),
        df_temp["Asset Age"] > 5
    ]
    choices = ["New (0-1 year)", "Active (1-3 years)", "Aging (3-5 years)", "Old (5+ years)"]
    df_temp["Age Category"] = pd.Series(pd.NA, dtype="object")
    for condition, choice in zip(conditions, choices):
        df_temp.loc[condition, "Age Category"] = choice

    age_counts = df_temp["Age Category"].value_counts()
    avg_age = df_temp[df_temp["Asset Age"] > 0]["Asset Age"].mean()

    col1, col2, col3, col4, col5 = st.columns(5)

    metrics = [
        (col1, "AVERAGE AGE", f"{avg_age:.1f}", "YEARS", "card-info"),
        (col2, "NEW (0-1YR)", age_counts.get("New (0-1 year)", 0), "", "card-success"),
        (col3, "ACTIVE (1-3YR)", age_counts.get("Active (1-3 years)", 0), "", "card-primary"),
        (col4, "AGING (3-5YR)", age_counts.get("Aging (3-5 years)", 0), "", "card-warning"),
        (col5, "OLD (5+YR)", age_counts.get("Old (5+ years)", 0), "", "card-danger")
    ]

    for col, label, value, extra, card_class in metrics:
        with col:
            extra_html = f'<div class="metric-label">{extra}</div>' if extra else ''
            st.markdown(f"""
                <div class="metric-card {card_class}">
                    <div class="metric-label">{label}</div>
                    <h2>{value}</h2>
                    {extra_html}
                </div>
            """, unsafe_allow_html=True)

def show_category_metrics_with_region(df, model_col, asset_type):
    """Display unit breakdown and regional analysis"""
    if not model_col:
        st.warning("Model column not found")
        return

    region_col = find_column(df, ["place"] if asset_type == "Workstation" else ["site", "user site", "usersite"])
    region_label = "Place" if asset_type == "Workstation" else "Site"

    col_left, col_right = st.columns([1, 1])

    with col_left:
        model_counts = df[model_col].value_counts().sort_values(ascending=False)
        st.markdown(f'<div class="section-header">Unit Breakdown by {escape(str(model_col))}</div>', unsafe_allow_html=True)

        model_df = pd.DataFrame({
            model_col: model_counts.index,
            "Total Units": model_counts.values
        })

        st.dataframe(model_df, use_container_width=True, hide_index=True)

    with col_right:
        if region_col and region_col in df.columns:
            st.markdown(f'<div class="section-header">Regional Breakdown by {escape(str(region_label))}</div>', unsafe_allow_html=True)

            pivot_data = df.groupby([region_col, model_col]).size().unstack(fill_value=0)
            pivot_data["Total"] = pivot_data.sum(axis=1)
            pivot_data.loc["Grand Total"] = pivot_data.sum()
            pivot_data = pivot_data.reset_index().rename(columns={region_col: "Region"})

            st.dataframe(pivot_data, use_container_width=True, hide_index=True)
        else:
            st.info(f"{region_label} column not found in Excel file")

# ============================================================================
# CHART FUNCTIONS
# ============================================================================

@st.cache_data
def create_pie_chart(df, model_col):
    """Create pie chart for asset distribution"""
    if not model_col:
        return None

    model_counts = df[model_col].value_counts()

    fig = px.pie(
        values=model_counts.values,
        names=model_counts.index,
        title=f"Asset Distribution by {model_col}",
        hole=0.4,
        color_discrete_sequence=['#0066B3', '#0080C9', '#00A3E0', '#28A745', '#FFC107', '#DC3545']
    )

    fig.update_traces(
        textposition='inside',
        textinfo='percent+label',
        hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
    )

    fig.update_layout(
        showlegend=True,
        height=400,
        margin=dict(t=50, b=0, l=0, r=0),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Poppins, sans-serif", color="#2C3E50")
    )
    return fig

@st.cache_data
def create_dimension_chart(df, dimension_col, dimension_label):
    """Create a bar chart for a named asset dimension."""
    if not dimension_col:
        return None

    dimension_counts = df[dimension_col].value_counts().head(10)

    fig = px.bar(
        x=dimension_counts.values,
        y=dimension_counts.index,
        orientation='h',
        title=f"Top 10 {dimension_label} by Asset Count",
        labels={'x': 'Asset Count', 'y': dimension_label},
        color_discrete_sequence=['#0066B3']
    )

    fig.update_layout(
        showlegend=False,
        height=400,
        margin=dict(t=50, b=50, l=0, r=0),
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        font=dict(family="Poppins, sans-serif", color="#2C3E50")
    )
    return fig


create_department_chart = lambda df, dept_col: create_dimension_chart(df, dept_col, "Department")

# ============================================================================
# SIDEBAR CONTROLS
# ============================================================================

def sidebar_controls(df, asset_type, model_col, type_col):
    """Create sidebar filter controls"""
    st.sidebar.markdown('<div class="sidebar-section">Asset Filters</div>', unsafe_allow_html=True)

    filters = {}

    # Model Filter
    if model_col:
        filters[model_col] = st.sidebar.multiselect(
            f"Filter by {model_col}",
            df[model_col].unique(),
            key="filter_model"
        )

    # Type Filter
    if type_col:
        filters[type_col] = st.sidebar.multiselect(
            f"Filter by {type_col}",
            df[type_col].unique(),
            key="filter_type"
        )

    # Site Filter
    site_col = "site" if "site" in df.columns else find_column(df, ["site", "user site", "usersite"])
    if site_col:
        filters[site_col] = st.sidebar.multiselect(
            f"Filter by {site_col}",
            df[site_col].unique(),
            key="filter_site"
        )

    # Location Filter
    location_col = "location" if "location" in df.columns else find_column(df, ["location"])
    if location_col:
        filters[location_col] = st.sidebar.multiselect(
            f"Filter by {location_col}",
            df[location_col].unique(),
            key="filter_location"
        )

    # Department Filter
    dept_col = "department" if "department" in df.columns else find_column(df, ["department", "user department"])
    if dept_col:
        filters[dept_col] = st.sidebar.multiselect(
            f"Filter by {dept_col}",
            df[dept_col].unique(),
            key="filter_department"
        )

    # Workstation-specific filters
    if asset_type == "Workstation":
        status_col = "workstation_status" if "workstation_status" in df.columns else find_column(df, ["workstation status", "workstationstatus"])
        if status_col:
            filters[status_col] = st.sidebar.multiselect(
                f"Filter by {status_col}",
                df[status_col].unique(),
                key="filter_status"
            )

        place_col = "place" if "place" in df.columns else find_column(df, ["place"])
        if place_col:
            filters[place_col] = st.sidebar.multiselect(
                f"Filter by {place_col}",
                df[place_col].unique(),
                key="filter_place"
            )

        state_col = "state" if "state" in df.columns else find_column(df, ["state"])
        if state_col:
            filters[state_col] = st.sidebar.multiselect(
                f"Filter by {state_col}",
                df[state_col].unique(),
                key="filter_state"
            )
    else:
        # Mobile-specific filters
        programme_col = "programme" if "programme" in df.columns else find_column(df, ["programme", "program"])
        if programme_col:
            filters[programme_col] = st.sidebar.multiselect(
                f"Filter by {programme_col}",
                df[programme_col].unique(),
                key="filter_programme"
            )

        state_col = "state" if "state" in df.columns else find_column(df, ["state"])
        if state_col:
            filters[state_col] = st.sidebar.multiselect(
                f"Filter by {state_col}",
                df[state_col].unique(),
                key="filter_state_mobile"
            )

    st.sidebar.markdown('<div class="sidebar-section">Replacement Planning</div>', unsafe_allow_html=True)
    expired_models = st.sidebar.multiselect(
        "Mark for Replacement",
        options=df[model_col].unique() if model_col else [],
        help="Select assets that need replacement"
    )

    st.sidebar.markdown('<div class="sidebar-section">Search</div>', unsafe_allow_html=True)
    search_query = st.sidebar.text_input("Search all fields", placeholder="Enter search term...")

    # Apply filters
    filtered_df = df.copy()
    for col, selected_values in filters.items():
        if selected_values:
            filtered_df = filtered_df[filtered_df[col].isin(selected_values)]

    if search_query:
        filtered_df = apply_literal_search(filtered_df, search_query)

    expired_df = filtered_df[
        filtered_df["ITAM Lifecycle Status"].eq("Expired")
    ].copy() if "ITAM Lifecycle Status" in filtered_df else pd.DataFrame(columns=filtered_df.columns)
    selected_replacement_df = filtered_df[
        filtered_df[model_col].isin(expired_models)
    ].copy() if expired_models and model_col else pd.DataFrame(columns=filtered_df.columns)

    return filtered_df, expired_df, selected_replacement_df

# ============================================================================
# FILE OPERATIONS
# ============================================================================

def export_to_excel(df, filename="asset_data.xlsx"):
    """Export a sanitized copy so source text cannot become Excel formulas."""
    export_df = df.copy()
    for column in export_df.select_dtypes(include=["object", "string"]).columns:
        export_df[column] = export_df[column].map(
            lambda value: "'" + value if isinstance(value, str) and value.startswith(("=", "+", "-", "@")) else value
        )
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df.to_excel(writer, index=False, sheet_name='Assets')
    output.seek(0)
    return output

def create_sample_workstation_file():
    """Create sample workstation Excel file"""
    sample_data = {
        'Asset Tag': ['WS001', 'WS002', 'WS003', 'WS004', 'WS005'],
        'Model': ['Dell Latitude 5420', 'HP EliteBook 840', 'Lenovo ThinkPad X1', 'Dell Optiplex 7090', 'HP ProBook 450'],
        'Workstation Type': ['Laptop', 'Laptop', 'Laptop', 'Desktop', 'Laptop'],
        'Serial Number': ['SN12345', 'SN12346', 'SN12347', 'SN12348', 'SN12349'],
        'User': ['John Doe', 'Jane Smith', 'Bob Wilson', 'Alice Brown', 'Charlie Davis'],
        'User Email': ['john.doe@company.com', 'jane.smith@company.com', 'bob.wilson@company.com', 'alice.brown@company.com', 'charlie.davis@company.com'],
        'Department': ['IT', 'Finance', 'HR', 'Operations', 'Marketing'],
        'Location': ['HQ Building A', 'HQ Building B', 'Branch Office', 'HQ Building A', 'Remote'],
        'Site': ['Headquarters', 'Headquarters', 'Branch', 'Headquarters', 'Remote'],
        'Year Of Purchase': [2022, 2021, 2023, 2020, 2022],
        'Warranty Expiry': ['2025-12-31', '2024-11-30', '2026-06-30', '2023-10-31', '2025-08-15'],
        'Place': ['Malaysia', 'Malaysia', 'Singapore', 'Malaysia', 'Malaysia'],
        'Workstation Status': ['Active', 'Active', 'Active', 'Retired', 'Active'],
        'State': ['In Use', 'In Use', 'In Use', 'Storage', 'In Use']
    }

    df = pd.DataFrame(sample_data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Workstation Assets')
    output.seek(0)
    return output

def create_sample_mobile_file():
    """Create sample mobile Excel file"""
    sample_data = {
        'Asset Tag': ['MB001', 'MB002', 'MB003', 'MB004', 'MB005'],
        'Product': ['iPhone 13 Pro', 'Samsung Galaxy S21', 'iPad Air', 'iPhone 12', 'Samsung Tab S8'],
        'Product Type': ['IT Smartphones', 'IT Smartphones', 'IT Tablets', 'IT Smartphones', 'IT Tablets'],
        'Serial Number': ['SNM12345', 'SNM12346', 'SNM12347', 'SNM12348', 'SNM12349'],
        'User': ['John Doe', 'Jane Smith', 'Bob Wilson', 'Alice Brown', 'Charlie Davis'],
        'User Email': ['john.doe@company.com', 'jane.smith@company.com', 'bob.wilson@company.com', 'alice.brown@company.com', 'charlie.davis@company.com'],
        'Department': ['IT', 'Sales', 'Operations', 'Finance', 'HR'],
        'Location': ['HQ Building A', 'Field', 'HQ Building B', 'HQ Building A', 'Branch Office'],
        'Site': ['Headquarters', 'Field', 'Headquarters', 'Headquarters', 'Branch'],
        'Year Of Purchase': [2022, 2021, 2023, 2021, 2022],
        'Programme': ['Enterprise Mobility', 'Sales Force', 'Operations', 'Finance', 'HR Management'],
        'State': ['In Use', 'In Use', 'In Use', 'In Use', 'In Use']
    }

    df = pd.DataFrame(sample_data)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Mobile Assets')
    output.seek(0)
    return output

# ============================================================================
if False and __name__ == '__main__':
    # MAIN APPLICATION
    # ============================================================================

    inject_professional_css()

    st.markdown("""
        <div style='text-align: center; padding: 20px 0;'>
            <h1>Asset Management Dashboard System</h1>
            <p style='color: #7B8794; font-size: 1rem; font-weight: 400;'>Professional Asset Tracking & Analytics Platform</p>
        </div>
    """, unsafe_allow_html=True)

    uploaded_file = st.file_uploader("Upload Excel File (.xlsx)", type=["xlsx"])

    if uploaded_file is not None:
        try:
            # Validate file format
            uploaded_file.seek(0)
            file_bytes = uploaded_file.read()

            if not file_bytes.startswith(b'PK'):
                st.error("File Format Error")
                st.warning("The uploaded file is not a valid Excel (.xlsx) file.")

                with st.expander("Troubleshooting Guide - Click to Expand", expanded=True):
                    st.markdown("""
                    ### Common Causes & Solutions:

                    #### 1. File Permission Restrictions (Most Common)
                    Your Excel file may have permission restrictions (Internal Use, Confidential, etc.)

                    **Solution:**
                    - Open file in Microsoft Excel
                    - Click **File** → **Info** → **Protect Workbook**
                    - Remove all restrictions/permissions
                    - **Save As** → Choose **Excel Workbook (*.xlsx)**
                    - Upload the new unrestricted file

                    ---

                    #### 2. Wrong File Format
                    File might be `.xls` (old format) renamed to `.xlsx`

                    **Solution:**
                    - Open in Excel
                    - **File** → **Save As**
                    - Select format: **Excel Workbook (*.xlsx)**
                    - Save with new name

                    ---

                    #### 3. Corrupted File
                    File may be damaged during transfer

                    **Solution:**
                    - Open file in Excel (Excel may auto-repair)
                    - If warning appears, click **Yes** to repair
                    - **Save As** new file
                    - Try uploading new file

                    ---

                    #### 4. Password Protected
                    File has password protection

                    **Solution:**
                    - Open in Excel
                    - **File** → **Info** → **Protect Workbook**
                    - Remove password
                    - Save and retry

                    ---

                    #### 5. CSV Saved as .xlsx
                    CSV file with extension changed to .xlsx

                    **Solution:**
                    - Open file in Excel
                    - **Save As** → **Excel Workbook (*.xlsx)**

                    ---

                    #### 6. Incomplete Download
                    File not fully downloaded from email/cloud

                    **Solution:**
                    - Download file again
                    - Verify file size matches original
                    - Try uploading again
                    """)

                    st.info("Quick Fix: Use the sample templates below, then copy your data into them.")

                    col_sample1, col_sample2 = st.columns(2)
                    with col_sample1:
                        sample_ws = create_sample_workstation_file()
                        st.download_button(
                            label="Download Workstation Template",
                            data=sample_ws,
                            file_name="workstation_template.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    with col_sample2:
                        sample_mb = create_sample_mobile_file()
                        st.download_button(
                            label="Download Mobile Template",
                            data=sample_mb,
                            file_name="mobile_template.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                st.stop()

            # Read Excel file
            uploaded_file.seek(0)
            xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
            sheet_names = xls.sheet_names
            selected_sheet = st.sidebar.selectbox("Select Sheet", sheet_names)

            # Detect header row
            uploaded_file.seek(0)
            header_row = detect_header_row(uploaded_file, selected_sheet)
            if header_row is None:
                st.warning(
                    "Automatic header detection could not find a confident company export header. "
                    "Enable Manual Header Row Selection and choose the table header row."
                )

            st.sidebar.markdown("---")
            st.sidebar.markdown('<div class="sidebar-section">Header Settings</div>', unsafe_allow_html=True)
            use_manual = st.sidebar.checkbox("Manual Header Row Selection", value=False)
            if use_manual:
                header_row = st.sidebar.number_input(
                    "Header Row (0-based)",
                    min_value=0,
                    max_value=20,
                    value=header_row if header_row is not None else 0,
                )
                st.sidebar.success(f"Using row {header_row} as header")
            elif header_row is None:
                st.stop()

            # Load data
            uploaded_file.seek(0)
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row, engine='openpyxl')

            df.columns = [str(c).strip() for c in df.columns]
            df = df.loc[:, ~df.columns.duplicated(keep='first')]

            # Detect asset type
            asset_type = detect_asset_type_from_data(df)
            if asset_type == "Unknown":
                st.error("Could not confidently detect this export. Expected 'Workstation Type' or Product Type values of 'IT Smartphones' or 'IT Tablets'.")
                st.stop()

            missing_required = validate_source_columns(df, asset_type)
            if missing_required:
                st.error(
                    f"This {asset_type.lower()} export is missing required columns: "
                    + ", ".join(missing_required)
                )
                st.info("Analysis stopped safely. Check the export headers and select the correct header row.")
                st.stop()

            df = build_canonical_dataframe(df, asset_type)
            st.sidebar.success(f"Detected: **{asset_type}** Assets")

            # Show columns
            with st.sidebar.expander("Excel Columns Found", expanded=False):
                st.write(f"**Total columns:** {len(df.columns)}")
                for idx, col in enumerate(df.columns, 1):
                    st.text(f"{idx}. {col}")

            # Get key columns
            model_col = get_model_column(df, asset_type)
            type_col = get_type_column(df, asset_type)

            if not model_col:
                st.error("Model column not found in Excel file.")
                st.info("Ensure Excel has 'Model' (Workstation) or 'Product' (Mobile) column")
                st.stop()

            # Process data
            df = calculate_asset_age(df)

            expired_warranty_df = None
            if asset_type == "Workstation":
                df, expired_warranty_df = get_warranty_status(df)
            df = run_itam_audit(df)

            # Data validation
            st.markdown("---")
            with st.expander("Data Validation Report", expanded=False):
                issues = validate_data(df, asset_type, model_col)
                show_validation_issues(issues)

            audit_review_count = int(df["ITAM Review Required"].sum())
            audit_high_count = int(df["ITAM Highest Severity"].eq("High").sum())
            replacement_candidate_count = int(df["ITAM Replacement Candidate"].sum())
            st.markdown('<div class="section-header">Audit Summary</div>', unsafe_allow_html=True)
            audit_col1, audit_col2, audit_col3 = st.columns(3)
            audit_metrics = [
                (audit_col1, "ASSETS REQUIRING REVIEW", audit_review_count, "card-warning"),
                (audit_col2, "HIGH SEVERITY ASSETS", audit_high_count, "card-danger"),
                (audit_col3, "REPLACEMENT CANDIDATES", replacement_candidate_count, "card-info"),
            ]
            for column, label, value, card_class in audit_metrics:
                with column:
                    st.markdown(f"""
                        <div class="metric-card {card_class}">
                            <div class="metric-label">{label}</div>
                            <h2>{value}</h2>
                        </div>
                    """, unsafe_allow_html=True)

            # Sidebar controls
            df_filtered, df_expired, df_selected_replacement = sidebar_controls(df, asset_type, model_col, type_col)

            # Export section
            st.sidebar.markdown("---")
            st.sidebar.markdown('<div class="sidebar-section">Export Data</div>', unsafe_allow_html=True)

            if asset_type == "Workstation":
                col_exp1, col_exp2, col_exp3 = st.sidebar.columns(3)
            else:
                col_exp1, col_exp2 = st.sidebar.columns(2)

            with col_exp1:
                excel_data = export_to_excel(df_filtered)
                st.download_button(
                    label="All",
                    data=excel_data,
                    file_name=f"{asset_type.lower()}_assets_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Export all filtered data"
                )

            with col_exp2:
                if df_expired is not None and not df_expired.empty:
                    excel_expired = export_to_excel(df_expired)
                    st.download_button(
                        label="Expired",
                        data=excel_expired,
                        file_name=f"{asset_type.lower()}_expired_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Export expired assets only"
                    )

            if asset_type == "Workstation":
                with col_exp3:
                    if expired_warranty_df is not None and not expired_warranty_df.empty:
                        excel_warranty = export_to_excel(expired_warranty_df)
                        st.download_button(
                            label="Warranty",
                            data=excel_warranty,
                            file_name=f"warranty_expired_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="Export expired warranties"
                        )

            # Help section
            st.sidebar.markdown("---")
            st.sidebar.markdown('<div class="sidebar-section">Help & Support</div>', unsafe_allow_html=True)

            with st.sidebar.expander("Troubleshooting"):
                st.markdown("""
                **Common Issues:**

                **Model column not found**
                - Ensure Excel has 'Model' (Workstation) or 'Product' (Mobile) column

                **Error reading file**
                - Save file as .xlsx format
                - Remove password protection

                **Wrong asset type detected**
                - Check column names match expected format

                **Data not showing correctly**
                - Verify header row is correct
                - Check for merged cells in Excel
                """)

            with st.sidebar.expander("Contact Support"):
                st.markdown("""
                **Need Help?**

                Email: khalis.abdrahim@gmail.com

                **Response Time:**
                Mon-Fri: Within 24 hours
                Weekend: Within 48 hours
                """)

            st.sidebar.markdown("---")
            st.sidebar.markdown("""
                <div style='text-align: center; color: #7B8794; font-size: 0.85em;'>
                    <strong>AssetLens</strong><br/>
                    Version 2.4.0<br/>
                    <br/>
                    &copy; 2025 All rights reserved.<br/>
                    Developed by <strong>MKAR</strong><br/>
                </div>
            """, unsafe_allow_html=True)

            # Dashboard Summary
            st.markdown('<div class="section-header">Dashboard Summary</div>', unsafe_allow_html=True)
            show_summary_cards(df_filtered, df_expired)

            # Type Statistics
            if type_col:
                st.markdown("---")
                show_type_cards(df_filtered, type_col, asset_type)

            # Warranty Status
            if asset_type == "Workstation" and "Warranty Status" in df_filtered.columns:
                st.markdown("---")
                st.markdown('<div class="section-header">Warranty Status</div>', unsafe_allow_html=True)
                show_warranty_summary(df_filtered, model_col)

            # Asset Age Analysis
            if "Asset Age" in df_filtered.columns:
                st.markdown("---")
                st.markdown('<div class="section-header">Asset Age Analysis</div>', unsafe_allow_html=True)
                show_asset_age_summary(df_filtered)

            # Category Metrics
            st.markdown("---")
            show_category_metrics_with_region(df_filtered, model_col, asset_type)

            # Visual Analytics
            st.markdown("---")
            st.markdown('<div class="section-header">Visual Analytics</div>', unsafe_allow_html=True)

            col_chart1, col_chart2 = st.columns(2)

            with col_chart1:
                pie_fig = create_pie_chart(df_filtered, model_col)
                if pie_fig:
                    st.plotly_chart(pie_fig, use_container_width=True)

            with col_chart2:
                dept_col = find_column(df_filtered, ["department", "user department"])
                dept_fig = create_dimension_chart(df_filtered, dept_col, "Department")
                if dept_fig:
                    st.plotly_chart(dept_fig, use_container_width=True)
                else:
                    st.info("Department data not available")

            location_col = find_column(df_filtered, ["location"])
            loc_fig = create_dimension_chart(df_filtered, location_col, "Location")
            if loc_fig:
                st.plotly_chart(loc_fig, use_container_width=True)

            # Replacement Assets
            if not df_selected_replacement.empty:
                st.markdown("---")
                st.markdown('<div class="section-header">Assets Selected for Replacement Review</div>', unsafe_allow_html=True)
                st.dataframe(df_selected_replacement, use_container_width=True, hide_index=True)

            # Asset Details
            st.markdown("---")
            st.markdown('<div class="section-header">Asset Details</div>', unsafe_allow_html=True)

            year_col = find_column(df_filtered, ["year of purchase", "yearofpurchase"])
            display_columns = [col for col in df_filtered.columns if col != year_col]

            st.info(f"Displaying {len(display_columns)} columns from Excel file")
            st.dataframe(df_filtered[display_columns], use_container_width=True, height=600)

        except Exception as e:
            st.error(f"Error reading Excel file: {str(e)}")
            st.warning("**Troubleshooting Tips:**")
            st.markdown("""
            1. Ensure file format is .xlsx (Excel)
            2. File must have header row with clear column names
            3. Try opening file in Excel and save again
            4. Remove password protection if any
            5. Ensure file is not corrupted
            """)

    else:
        # Language selection
        if 'language' not in st.session_state:
            st.session_state.language = 'EN'

        col_lang1, col_lang2, col_space = st.columns([1, 1, 8])
        with col_lang1:
            if st.button("English", use_container_width=True,
                         type="primary" if st.session_state.language == 'EN' else "secondary"):
                st.session_state.language = 'EN'
                st.rerun()
        with col_lang2:
            if st.button("Bahasa", use_container_width=True,
                         type="primary" if st.session_state.language == 'MY' else "secondary"):
                st.session_state.language = 'MY'
                st.rerun()

        st.markdown("---")

        if st.session_state.language == 'EN':
            st.info("Please upload your Excel file to get started.")

            st.markdown("### How to Use This Dashboard")
            st.markdown("""
            This dashboard reads **original column names** directly from your Excel file.

            #### Key Features
            - Auto-detection of asset type (Workstation or Mobile)
            - All original columns displayed
            - Regional breakdown by Place/Site
            - Clean and professional UI
            - Smart filtering and search

            #### Required Columns

            **Workstation Assets:**
            - `Model` (Required)
            - `Workstation Type`, `Warranty Expiry`, `Place` (Optional)

            **Mobile Assets:**
            - `Product` (Required)
            - `Product Type`, `Programme`, `Site` (Optional)
            """)

            col_sample1, col_sample2, col_space2 = st.columns([2, 2, 6])
            with col_sample1:
                sample_ws = create_sample_workstation_file()
                st.download_button(
                    label="Workstation Sample",
                    data=sample_ws,
                    file_name="sample_workstation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col_sample2:
                sample_mb = create_sample_mobile_file()
                st.download_button(
                    label="Mobile Sample",
                    data=sample_mb,
                    file_name="sample_mobile.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            st.success("""
            **Your Data Security**
            - Files are NOT stored on any server
            - Processing happens in memory only
            - Data stays completely private
            """)


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
