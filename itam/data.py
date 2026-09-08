import re

import pandas as pd


def normalize_text(text):
    """Normalize text for column matching."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def find_column(df, search_terms):
    """Find a column by one or more search terms."""
    if isinstance(search_terms, str):
        search_terms = [search_terms]

    normalized_cols = {normalize_text(col): col for col in df.columns}

    for term in search_terms:
        normalized_term = normalize_text(term)
        for norm_col, orig_col in normalized_cols.items():
            if normalized_term in norm_col:
                return orig_col
    return None


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
    """Get the canonical model column."""
    return "model" if "model" in df.columns else None


def get_type_column(df, asset_type):
    """Get the source asset subtype column."""
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

            if matched_type and len(matched_identity) >= 3:
                score = len(matched_identity) * 2 + len(matched_type)
                candidates.append((score, i))

        if candidates:
            return max(candidates, key=lambda candidate: (candidate[0], -candidate[1]))[1]
        return None
    except Exception:
        return None