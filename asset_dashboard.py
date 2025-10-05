import streamlit as st
import pandas as pd
import re
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go

# ============================================================================
# PAGE CONFIGURATION
# ============================================================================
st.set_page_config(
    page_title="Asset Management Dashboard System",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================================
# CONSTANTS
# ============================================================================
WORKSTATION_IDENTIFIERS = ["workstation", "model", "warranty", "place"]
MOBILE_IDENTIFIERS = ["product", "programme", "program"]

# ============================================================================
# AIR SELANGOR THEME CSS
# ============================================================================
def inject_professional_css():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
        
        /* Air Selangor Color Palette */
        :root {
            --primary-blue: #0066B3;
            --secondary-blue: #0080C9;
            --light-blue: #E6F3FF;
            --accent-blue: #00A3E0;
            --text-primary: #2C3E50;
            --text-secondary: #7B8794;
            --background: #F5F7FA;
            --border: #E0E6ED;
            --success: #28A745;
            --warning: #FFC107;
            --danger: #DC3545;
        }
        
        .main {
            background: var(--background);
            font-family: 'Poppins', sans-serif;
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
        
        /* Sidebar Styling */
        [data-testid="stSidebar"] {
            background: #FFFFFF;
            border-right: 1px solid var(--border);
        }
        
        [data-testid="stSidebar"] * {
            color: var(--text-primary) !important;
        }
        
        [data-testid="stSidebar"] .stMarkdown {
            font-size: 0.9rem;
        }
        
        [data-testid="stSidebar"] label {
            font-weight: 500 !important;
            color: var(--text-primary) !important;
            font-size: 0.875rem !important;
        }
        
        /* Metric Cards - Soft Blue Background */
        .metric-card {
            background: linear-gradient(135deg, #E8F4FC 0%, #D6EDFA 100%);
            border-radius: 12px;
            padding: 24px;
            color: #1A4D7A;
            text-align: center;
            font-weight: 500;
            margin-bottom: 16px;
            box-shadow: 0 2px 12px rgba(0, 102, 179, 0.1);
            transition: all 0.3s ease;
            border: none;
        }
        
        .metric-card:hover {
            box-shadow: 0 4px 20px rgba(0, 102, 179, 0.18);
            transform: translateY(-3px);
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
        
        /* Card Color Variants - Soft Backgrounds */
        .card-primary { 
            background: linear-gradient(135deg, #E8F4FC 0%, #D6EDFA 100%);
        }
        .card-primary h2 { 
            color: #0066B3 !important; 
        }
        .card-primary .metric-label {
            color: #1A4D7A;
        }
        
        .card-success { 
            background: linear-gradient(135deg, #E8F8F0 0%, #D1F2E0 100%);
        }
        .card-success h2 { 
            color: #1B7A4C !important; 
        }
        .card-success .metric-label {
            color: #1B7A4C;
        }
        
        .card-warning { 
            background: linear-gradient(135deg, #FFF8E6 0%, #FFF0CC 100%);
        }
        .card-warning h2 { 
            color: #B8860B !important; 
        }
        .card-warning .metric-label {
            color: #8B6914;
        }
        
        .card-info { 
            background: linear-gradient(135deg, #E6F7FC 0%, #CCF0FA 100%);
        }
        .card-info h2 { 
            color: #007BA7 !important; 
        }
        .card-info .metric-label {
            color: #005F87;
        }
        
        .card-danger { 
            background: linear-gradient(135deg, #FFE8EB 0%, #FFD6DC 100%);
        }
        .card-danger h2 { 
            color: #B91C2E !important; 
        }
        .card-danger .metric-label {
            color: #8B1624;
        }
        
        /* Type Cards - Soft Blue Background */
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
        
        /* Section Headers */
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
        
        /* Sidebar Section Headers */
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
        
        /* Validation Severity Badges */
        .severity-badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 4px;
            font-weight: 600;
            font-size: 0.75rem;
            letter-spacing: 0.5px;
            margin-right: 8px;
        }
        
        .severity-high {
            background: #DC3545;
            color: white;
        }
        
        .severity-medium {
            background: #FFC107;
            color: #2C3E50;
        }
        
        .severity-low {
            background: #28A745;
            color: white;
        }
        
        /* Data Tables */
        .stDataFrame {
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 8px rgba(0, 102, 179, 0.06);
            border: 1px solid var(--border);
        }
        
        /* Expanders */
        .streamlit-expanderHeader {
            background: #FFFFFF;
            border: 1px solid var(--border);
            border-radius: 8px;
            font-weight: 600;
            color: var(--text-primary);
        }
        
        /* Buttons */
        .stButton>button {
            background: linear-gradient(135deg, var(--primary-blue) 0%, var(--secondary-blue) 100%);
            color: white;
            border: none;
            border-radius: 6px;
            padding: 10px 20px;
            font-weight: 500;
            transition: all 0.3s ease;
            box-shadow: 0 2px 6px rgba(0, 102, 179, 0.2);
        }
        
        .stButton>button:hover {
            box-shadow: 0 4px 12px rgba(0, 102, 179, 0.3);
            transform: translateY(-1px);
        }
        
        /* Download Buttons */
        .stDownloadButton>button {
            background: #FFFFFF;
            color: var(--primary-blue);
            border: 1px solid var(--primary-blue);
            border-radius: 6px;
            font-weight: 500;
            transition: all 0.3s ease;
            font-size: 0.85rem;
        }
        
        .stDownloadButton>button:hover {
            background: var(--primary-blue);
            color: white;
            box-shadow: 0 4px 12px rgba(0, 102, 179, 0.2);
        }
        
        /* Alerts */
        .stAlert {
            border-radius: 8px;
            border-left: 4px solid;
            box-shadow: 0 2px 8px rgba(0, 102, 179, 0.06);
        }
        
        /* File Uploader */
        [data-testid="stFileUploader"] {
            background: white;
            border-radius: 8px;
            padding: 24px;
            box-shadow: 0 2px 8px rgba(0, 102, 179, 0.06);
            border: 2px dashed var(--border);
        }
        
        /* Dividers */
        hr {
            margin: 32px 0;
            border: none;
            height: 1px;
            background: linear-gradient(90deg, transparent, var(--border), transparent);
        }
        
        /* Multi-select Tags */
        .stMultiSelect [data-baseweb="tag"] {
            background-color: var(--light-blue);
            color: var(--primary-blue);
            border-radius: 6px;
        }
        
        /* Info/Success/Warning Boxes */
        .stInfo { background: var(--light-blue); border-left: 4px solid var(--primary-blue); }
        .stSuccess { background: #D4EDDA; border-left: 4px solid var(--success); }
        .stWarning { background: #FFF3CD; border-left: 4px solid var(--warning); }
        
        /* Mobile Responsive */
        @media (max-width: 768px) {
            h1 { font-size: 1.6rem !important; }
            .metric-card h2 { font-size: 1.8rem; }
            .section-header { font-size: 1rem; padding: 12px 16px; }
        }
        </style>
    """, unsafe_allow_html=True)

# ============================================================================
# UTILITY FUNCTIONS
# ============================================================================

def normalize_text(text):
    """Normalize text for column matching"""
    return re.sub(r'[^a-z0-9]', '', str(text).lower())

def find_column(df, search_terms):
    """Find column by multiple search terms"""
    if isinstance(search_terms, str):
        search_terms = [search_terms]
    
    for col in df.columns:
        normalized_col = normalize_text(col)
        for term in search_terms:
            normalized_term = normalize_text(term)
            if normalized_term in normalized_col:
                return col
    return None

# ============================================================================
# ASSET TYPE DETECTION
# ============================================================================

def detect_asset_type(df_columns):
    """Auto-detect asset type from column names"""
    normalized = [normalize_text(col) for col in df_columns]
    
    workstation_score = sum(1 for identifier in WORKSTATION_IDENTIFIERS 
                           if any(identifier in norm for norm in normalized))
    mobile_score = sum(1 for identifier in MOBILE_IDENTIFIERS 
                      if any(identifier in norm for norm in normalized))
    
    return "Workstation" if workstation_score > mobile_score else "Mobile"

def get_model_column(df, asset_type):
    """Get model column based on asset type"""
    if asset_type == "Workstation":
        return find_column(df, ["model"])
    else:
        for col in df.columns:
            if normalize_text(col) == "product":
                return col
    return None

def get_type_column(df, asset_type):
    """Get type column based on asset type"""
    if asset_type == "Workstation":
        return find_column(df, ["workstation type", "workstationtype"])
    else:
        return find_column(df, ["product type", "producttype"])

def detect_header_row(excel_file, sheet_name):
    """Auto-detect header row in Excel file"""
    try:
        preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=15, engine='openpyxl')
        
        keywords = [
            "model", "serial", "user", "department", 
            "asset", "workstation", "location", "site",
            "computer", "employee", "email", "product",
            "mobile", "programme", "program"
        ]
        
        for i, row in preview.iterrows():
            values = row.astype(str).str.lower().str.strip().tolist()
            matches = sum(1 for v in values if any(keyword in v for keyword in keywords))
            
            if matches >= 3:
                return i
        
        return 0
    except:
        return 0

# ============================================================================
# DATA PROCESSING FUNCTIONS
# ============================================================================

def calculate_asset_age(df):
    """Calculate asset age from purchase year"""
    year_col = find_column(df, ["year of purchase", "yearofpurchase"])
    if year_col:
        current_year = pd.Timestamp.now().year
        try:
            df["Asset Age"] = current_year - pd.to_numeric(df[year_col], errors='coerce')
            df["Asset Age"] = df["Asset Age"].fillna(0).astype(int)
        except:
            df["Asset Age"] = 0
    else:
        df["Asset Age"] = 0
    return df

def get_warranty_status(df):
    """Calculate warranty status"""
    warranty_col = find_column(df, ["warranty expiry", "warrantyexpiry"])
    if not warranty_col:
        return df, None

    try:
        df_temp = df.copy()
        df_temp["Warranty Expiry Date"] = pd.to_datetime(df_temp[warranty_col], errors='coerce')

        today = pd.Timestamp.now()
        df_temp["Days to Expiry"] = (df_temp["Warranty Expiry Date"] - today).dt.days

        df_temp["Warranty Status"] = "Unknown"
        df_temp.loc[df_temp["Days to Expiry"] < 0, "Warranty Status"] = "Expired"
        df_temp.loc[(df_temp["Days to Expiry"] >= 0) & (df_temp["Days to Expiry"] <= 90), "Warranty Status"] = "Expiring Soon"
        df_temp.loc[df_temp["Days to Expiry"] > 90, "Warranty Status"] = "Active"

        expired_warranty_df = df_temp[df_temp["Warranty Status"] == "Expired"].copy()

        return df_temp, expired_warranty_df
    except:
        return df, None

# ============================================================================
# DATA VALIDATION
# ============================================================================

def validate_data(df, asset_type, model_col):
    """Validate data and return list of issues"""
    issues = []

    asset_tag_col = find_column(df, ["asset tag", "assettag"])
    serial_col = find_column(df, ["serial number", "serialnumber"])
    user_col = find_column(df, ["user"])
    email_col = find_column(df, ["email"])
    dept_col = find_column(df, ["department", "user department"])
    location_col = find_column(df, ["location"])

    # Check duplicate asset tags
    if asset_tag_col:
        duplicates = df[df[asset_tag_col].duplicated(keep=False) & df[asset_tag_col].notna()]
        if not duplicates.empty:
            dup_tags = duplicates[asset_tag_col].unique()
            display_cols = [c for c in [asset_tag_col, model_col, serial_col, user_col] if c]
            issues.append({
                "type": "Duplicate Asset Tags",
                "count": len(dup_tags),
                "details": f"Found {len(dup_tags)} duplicate asset tags",
                "severity": "high",
                "data": duplicates[display_cols].sort_values(asset_tag_col)
            })

    # Check duplicate serial numbers
    if serial_col:
        duplicates = df[df[serial_col].duplicated(keep=False) & df[serial_col].notna()]
        if not duplicates.empty:
            dup_serials = duplicates[serial_col].unique()
            display_cols = [c for c in [serial_col, model_col, asset_tag_col, user_col] if c]
            issues.append({
                "type": "Duplicate Serial Numbers",
                "count": len(dup_serials),
                "details": f"Found {len(dup_serials)} duplicate serial numbers",
                "severity": "high",
                "data": duplicates[display_cols].sort_values(serial_col)
            })

    # Check missing users
    if user_col:
        missing_users = df[df[user_col].isna() | (df[user_col] == "")]
        if not missing_users.empty:
            display_cols = [c for c in [asset_tag_col, model_col, serial_col, dept_col, location_col] if c]
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

    # Check missing department
    if dept_col:
        missing_dept = df[df[dept_col].isna() | (df[dept_col] == "")]
        if not missing_dept.empty:
            display_cols = [c for c in [asset_tag_col, model_col, user_col, location_col] if c]
            issues.append({
                "type": "Missing Department",
                "count": len(missing_dept),
                "details": f"{len(missing_dept)} assets without department",
                "severity": "medium",
                "data": missing_dept[display_cols]
            })

    # Check missing location
    if location_col:
        missing_loc = df[df[location_col].isna() | (df[location_col] == "")]
        if not missing_loc.empty:
            display_cols = [c for c in [asset_tag_col, model_col, user_col, dept_col] if c]
            issues.append({
                "type": "Missing Location",
                "count": len(missing_loc),
                "details": f"{len(missing_loc)} assets without location",
                "severity": "medium",
                "data": missing_loc[display_cols]
            })

    # Check expired warranties
    if asset_type == "Workstation":
        warranty_col = find_column(df, ["warranty expiry", "warrantyexpiry"])
        if warranty_col:
            try:
                df_temp = df.copy()
                df_temp[warranty_col] = pd.to_datetime(df_temp[warranty_col], errors='coerce')
                expired_warranty = df_temp[df_temp[warranty_col] < pd.Timestamp.now()]
                if not expired_warranty.empty:
                    display_cols = [c for c in [asset_tag_col, model_col, user_col, warranty_col] if c]
                    issues.append({
                        "type": "Expired Warranties",
                        "count": len(expired_warranty),
                        "details": f"{len(expired_warranty)} assets with expired warranties",
                        "severity": "medium",
                        "data": expired_warranty[display_cols]
                    })
            except:
                pass

    return issues

def show_validation_issues(issues):
    """Display validation issues with colored severity badges"""
    if not issues:
        st.success("No data validation issues found")
        return

    st.warning(f"Found {len(issues)} validation issue(s)")

    for issue in issues:
        # Create colored severity badge
        severity_class = f"severity-{issue['severity']}"
        severity_label = issue['severity'].upper()
        
        with st.expander(f"{issue['type']} ({issue['count']})", expanded=False):
            st.markdown(f'<span class="severity-badge {severity_class}">{severity_label}</span> {issue["details"]}', 
                       unsafe_allow_html=True)
            if "data" in issue and not issue["data"].empty:
                st.markdown("**Affected Assets:**")
                st.dataframe(issue["data"], use_container_width=True, hide_index=True)

# ============================================================================
# DISPLAY FUNCTIONS
# ============================================================================

def show_summary_cards(df, df_expired=None):
    """Display summary metric cards"""
    total_assets = len(df)
    expired_assets = len(df_expired) if df_expired is not None else 0
    active_assets = total_assets - expired_assets
    replacement_rate = (expired_assets / total_assets * 100) if total_assets > 0 else 0

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(f"""
            <div class="metric-card card-primary">
                <div class="metric-label">TOTAL ASSETS</div>
                <h2>{total_assets}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
            <div class="metric-card card-success">
                <div class="metric-label">ACTIVE ASSETS</div>
                <h2>{active_assets}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
            <div class="metric-card card-warning">
                <div class="metric-label">EXPIRED ASSETS</div>
                <h2>{expired_assets}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
            <div class="metric-card card-info">
                <div class="metric-label">REPLACEMENT RATE</div>
                <h2>{replacement_rate:.1f}%</h2>
            </div>
        """, unsafe_allow_html=True)

def show_type_cards(df, type_col, asset_type):
    """Display asset type cards"""
    if not type_col:
        return

    st.markdown(f'<div class="section-header">{type_col} Statistics</div>', unsafe_allow_html=True)

    type_counts = df[type_col].value_counts().sort_values(ascending=False)
    num_types = len(type_counts)
    cols_per_row = min(4, num_types)
    cols = st.columns(cols_per_row)

    for idx, (wtype, count) in enumerate(type_counts.items()):
        col_idx = idx % cols_per_row

        with cols[col_idx]:
            st.markdown(f"""
                <div class="type-card card-primary">
                    <div class="type-label">{wtype}</div>
                    <div class="type-count">{count}</div>
                </div>
            """, unsafe_allow_html=True)

        if (idx + 1) % cols_per_row == 0 and (idx + 1) < num_types:
            cols = st.columns(cols_per_row)

def show_warranty_summary(df, model_col):
    """Display warranty status summary"""
    if "Warranty Status" not in df.columns:
        return

    status_counts = df["Warranty Status"].value_counts()

    col1, col2, col3 = st.columns(3)

    with col1:
        expired = status_counts.get("Expired", 0)
        st.markdown(f"""
            <div class="metric-card card-danger">
                <div class="metric-label">EXPIRED WARRANTY</div>
                <h2>{expired}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        expiring = status_counts.get("Expiring Soon", 0)
        st.markdown(f"""
            <div class="metric-card card-warning">
                <div class="metric-label">EXPIRING SOON (90 DAYS)</div>
                <h2>{expiring}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col3:
        active = status_counts.get("Active", 0)
        st.markdown(f"""
            <div class="metric-card card-success">
                <div class="metric-label">ACTIVE WARRANTY</div>
                <h2>{active}</h2>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    serial_col = find_column(df, ["serial number", "serialnumber"])
    user_col = find_column(df, ["user"])
    dept_col = find_column(df, ["department", "user department"])
    location_col = find_column(df, ["location"])
    warranty_col = find_column(df, ["warranty expiry", "warrantyexpiry"])

    # Show expandable sections for each warranty status
    expired_df = df[df["Warranty Status"] == "Expired"]
    if not expired_df.empty:
        with st.expander(f"Expired Warranty Assets ({len(expired_df)})", expanded=False):
            display_cols = [c for c in [model_col, serial_col, user_col, dept_col, location_col, warranty_col] if c]
            st.dataframe(expired_df[display_cols], use_container_width=True, hide_index=True)

    expiring_df = df[df["Warranty Status"] == "Expiring Soon"]
    if not expiring_df.empty:
        with st.expander(f"Expiring Soon Assets ({len(expiring_df)})", expanded=False):
            display_cols = [c for c in [model_col, serial_col, user_col, dept_col, location_col, warranty_col] if c]
            st.dataframe(expiring_df[display_cols], use_container_width=True, hide_index=True)

    active_df = df[df["Warranty Status"] == "Active"]
    if not active_df.empty:
        with st.expander(f"Active Warranty Assets ({len(active_df)})", expanded=False):
            display_cols = [c for c in [model_col, serial_col, user_col, dept_col, location_col, warranty_col] if c]
            st.dataframe(active_df[display_cols], use_container_width=True, hide_index=True)

def show_asset_age_summary(df):
    """Display asset age analysis"""
    if "Asset Age" not in df.columns or df["Asset Age"].sum() == 0:
        return

    df_temp = df.copy()
    df_temp["Age Category"] = "Unknown"
    df_temp.loc[df_temp["Asset Age"] <= 1, "Age Category"] = "New (0-1 year)"
    df_temp.loc[(df_temp["Asset Age"] > 1) & (df_temp["Asset Age"] <= 3), "Age Category"] = "Active (1-3 years)"
    df_temp.loc[(df_temp["Asset Age"] > 3) & (df_temp["Asset Age"] <= 5), "Age Category"] = "Aging (3-5 years)"
    df_temp.loc[df_temp["Asset Age"] > 5, "Age Category"] = "Old (5+ years)"

    age_counts = df_temp["Age Category"].value_counts()
    avg_age = df_temp[df_temp["Asset Age"] > 0]["Asset Age"].mean()

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.markdown(f"""
            <div class="metric-card card-info">
                <div class="metric-label">AVERAGE AGE</div>
                <h2>{avg_age:.1f}</h2>
                <div class="metric-label">YEARS</div>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
            <div class="metric-card card-success">
                <div class="metric-label">NEW (0-1YR)</div>
                <h2>{age_counts.get("New (0-1 year)", 0)}</h2>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
            <div class="metric-card card-primary">
                <div class="metric-label">ACTIVE (1-3YR)</div>
                <h2>{age_counts.get("Active (1-3 years)", 0)}</h2>
            </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
            <div class="metric-card card-warning">
                <div class="metric-label">AGING (3-5YR)</div>
                <h2>{age_counts.get("Aging (3-5 years)", 0)}</h2>
            </div>
        """, unsafe_allow_html=True)
    
    with col5:
        st.markdown(f"""
            <div class="metric-card card-danger">
                <div class="metric-label">OLD (5+YR)</div>
                <h2>{age_counts.get("Old (5+ years)", 0)}</h2>
            </div>
        """, unsafe_allow_html=True)

def show_category_metrics_with_region(df, model_col, asset_type):
    """Display unit breakdown and regional analysis"""
    if not model_col:
        st.warning("Model column not found")
        return

    if asset_type == "Workstation":
        region_col = find_column(df, ["place"])
        region_label = "Place"
    else:
        region_col = find_column(df, ["site", "user site", "usersite"])
        region_label = "Site"
    
    col_left, col_right = st.columns([1, 1])
    
    with col_left:
        model_counts = df[model_col].value_counts().sort_values(ascending=False)
        st.markdown(f'<div class="section-header">Unit Breakdown by {model_col}</div>', unsafe_allow_html=True)
        
        model_df = pd.DataFrame({
            model_col: model_counts.index,
            "Total Units": model_counts.values
        })
        
        st.dataframe(
            model_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                model_col: st.column_config.TextColumn(model_col, width="large"),
                "Total Units": st.column_config.NumberColumn("Total Units", width="small")
            }
        )
    
    with col_right:
        if region_col and region_col in df.columns:
            st.markdown(f'<div class="section-header">Regional Breakdown by {region_label}</div>', unsafe_allow_html=True)
            
            unique_models = df[model_col].unique()
            unique_regions = df[region_col].unique()
            
            breakdown_data = []
            
            for region in sorted(unique_regions):
                row_data = {"Region": region}
                region_df = df[df[region_col] == region]
                
                for model in unique_models:
                    model_count = len(region_df[region_df[model_col] == model])
                    if model_count > 0:
                        row_data[model] = model_count
                
                row_data["Total"] = len(region_df)
                breakdown_data.append(row_data)
            
            grand_total_row = {"Region": "Grand Total"}
            for model in unique_models:
                model_total = len(df[df[model_col] == model])
                if model_total > 0:
                    grand_total_row[model] = model_total
            grand_total_row["Total"] = len(df)
            breakdown_data.append(grand_total_row)
            
            breakdown_df = pd.DataFrame(breakdown_data)
            
            breakdown_df = breakdown_df.fillna(0)
            for col in breakdown_df.columns:
                if col != "Region":
                    breakdown_df[col] = breakdown_df[col].astype(int)
            
            st.dataframe(
                breakdown_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Region": st.column_config.TextColumn("Region", width="medium", pinned=True),
                    "Total": st.column_config.NumberColumn("Grand Total", width="small")
                }
            )
        else:
            st.info(f"{region_label} column not found in Excel file")

# ============================================================================
# CHART FUNCTIONS
# ============================================================================

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

def create_department_chart(df, dept_col):
    """Create bar chart for department distribution"""
    if not dept_col:
        return None
    
    dept_counts = df[dept_col].value_counts().head(10)
    
    fig = px.bar(
        x=dept_counts.values,
        y=dept_counts.index,
        orientation='h',
        title=f"Top 10 {dept_col} by Asset Count",
        labels={'x': 'Asset Count', 'y': dept_col},
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

def create_location_chart(df, location_col):
    """Create bar chart for location distribution"""
    if not location_col:
        return None
    
    loc_counts = df[location_col].value_counts().head(10)
    
    fig = px.bar(
        x=loc_counts.values,
        y=loc_counts.index,
        orientation='h',
        title=f"Top 10 {location_col} by Asset Count",
        labels={'x': 'Asset Count', 'y': location_col},
        color_discrete_sequence=['#28A745']
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

# ============================================================================
# SIDEBAR CONTROLS
# ============================================================================

def sidebar_controls(df, asset_type, model_col, type_col):
    """Create sidebar filter controls"""
    st.sidebar.markdown('<div class="sidebar-section">Asset Filters</div>', unsafe_allow_html=True)

    # Model Filter
    model_filter = []
    if model_col:
        filter_label = f"{model_col}" if asset_type == "Workstation" else f"{model_col}"
        model_filter = st.sidebar.multiselect(
            f"Filter by {model_col}",
            df[model_col].unique(),
            help=f"Select one or more {model_col.lower()} to filter"
        )

    # Type Filter
    type_filter = []
    if type_col:
        type_label = f"{type_col}"
        type_filter = st.sidebar.multiselect(
            f"Filter by {type_col}",
            df[type_col].unique(),
            help=f"Select asset type to filter"
        )
    
    # Site Filter
    site_col = find_column(df, ["site", "user site", "usersite"])
    site_filter = []
    if site_col:
        site_filter = st.sidebar.multiselect(
            f"Filter by {site_col}",
            df[site_col].unique(),
            help="Select site location"
        )
    
    # Location Filter
    location_col = find_column(df, ["location"])
    location_filter = []
    if location_col:
        location_filter = st.sidebar.multiselect(
            f"Filter by {location_col}",
            df[location_col].unique(),
            help="Select physical location"
        )
    
    # Department Filter
    dept_col = find_column(df, ["department", "user department"])
    dept_filter = []
    if dept_col:
        dept_filter = st.sidebar.multiselect(
            f"Filter by {dept_col}",
            df[dept_col].unique(),
            help="Select department"
        )
    
    # Workstation Status Filter
    status_col = find_column(df, ["workstation status", "workstationstatus"])
    status_filter = []
    if asset_type == "Workstation" and status_col:
        status_filter = st.sidebar.multiselect(
            f"Filter by {status_col}",
            df[status_col].unique(),
            help="Select workstation status"
        )
    
    # State Filter
    state_col = find_column(df, ["state"])
    state_filter = []
    if state_col:
        state_filter = st.sidebar.multiselect(
            f"Filter by {state_col}",
            df[state_col].unique(),
            help="Select current state"
        )

    # Place Filter (Workstation)
    place_col = find_column(df, ["place"])
    place_filter = []
    if asset_type == "Workstation" and place_col:
        place_filter = st.sidebar.multiselect(
            f"Filter by {place_col}",
            df[place_col].unique(),
            help="Select place/country"
        )

    # Programme Filter (Mobile)
    programme_col = find_column(df, ["programme", "program"])
    programme_filter = []
    if asset_type == "Mobile" and programme_col:
        programme_filter = st.sidebar.multiselect(
            f"Filter by {programme_col}",
            df[programme_col].unique(),
            help="Select programme"
        )

    st.sidebar.markdown('<div class="sidebar-section">Replacement Planning</div>', unsafe_allow_html=True)
    
    # Expired/Replacement Filter
    expired_models = []
    if model_col:
        expired_models = st.sidebar.multiselect(
            f"Mark for Replacement",
            options=df[model_col].unique(),
            help="Select assets that need replacement"
        )

    # Search
    st.sidebar.markdown('<div class="sidebar-section">Search</div>', unsafe_allow_html=True)
    search_query = st.sidebar.text_input(
        "Search all fields",
        placeholder="Enter search term...",
        help="Search across all columns"
    )

    # Apply filters
    filtered_df = df.copy()

    if model_filter and model_col:
        filtered_df = filtered_df[filtered_df[model_col].isin(model_filter)]
    if type_filter and type_col:
        filtered_df = filtered_df[filtered_df[type_col].isin(type_filter)]
    if site_filter and site_col:
        filtered_df = filtered_df[filtered_df[site_col].isin(site_filter)]
    if location_filter and location_col:
        filtered_df = filtered_df[filtered_df[location_col].isin(location_filter)]
    if dept_filter and dept_col:
        filtered_df = filtered_df[filtered_df[dept_col].isin(dept_filter)]
    if status_filter and status_col:
        filtered_df = filtered_df[filtered_df[status_col].isin(status_filter)]
    if state_filter and state_col:
        filtered_df = filtered_df[filtered_df[state_col].isin(state_filter)]
    if place_filter and place_col:
        filtered_df = filtered_df[filtered_df[place_col].isin(place_filter)]
    if programme_filter and programme_col:
        filtered_df = filtered_df[filtered_df[programme_col].isin(programme_filter)]

    expired_df = None
    if expired_models and model_col:
        expired_df = filtered_df[filtered_df[model_col].isin(expired_models)]

    if search_query:
        filtered_df = filtered_df[filtered_df.apply(
            lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1
        )]

    return filtered_df, expired_df

# ============================================================================
# FILE OPERATIONS
# ============================================================================

def export_to_excel(df, filename="asset_data.xlsx"):
    """Export dataframe to Excel"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Assets')
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
        'Product Type': ['Phone', 'Phone', 'Tablet', 'Phone', 'Tablet'],
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
        # Read Excel file
        uploaded_file.seek(0)
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xls.sheet_names
        selected_sheet = st.sidebar.selectbox("Select Sheet", sheet_names)

        # Detect header row
        uploaded_file.seek(0)
        header_row = detect_header_row(uploaded_file, selected_sheet)
        
        st.sidebar.markdown("---")
        st.sidebar.markdown('<div class="sidebar-section">Header Settings</div>', unsafe_allow_html=True)
        use_manual = st.sidebar.checkbox("Manual Header Row Selection", value=False)
        if use_manual:
            header_row = st.sidebar.number_input("Header Row (0-based)", min_value=0, max_value=20, value=header_row)
            st.sidebar.success(f"Using row {header_row} as header")

        # Load data
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row, engine='openpyxl')
        
        df.columns = [str(c).strip() for c in df.columns]
        df = df.loc[:, ~df.columns.duplicated(keep='first')]
        
        # Detect asset type
        asset_type = detect_asset_type(df.columns)
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

        # Data validation
        st.markdown("---")
        with st.expander("Data Validation Report", expanded=False):
            issues = validate_data(df, asset_type, model_col)
            show_validation_issues(issues)

        # Sidebar controls
        df_filtered, df_expired = sidebar_controls(df, asset_type, model_col, type_col)

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
                <strong>Asset Management Dashboard System</strong><br/>
                Version 2.3.0<br/>
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
            dept_fig = create_department_chart(df_filtered, dept_col)
            if dept_fig:
                st.plotly_chart(dept_fig, use_container_width=True)
            else:
                st.info("Department data not available")

        location_col = find_column(df_filtered, ["location"])
        loc_fig = create_location_chart(df_filtered, location_col)
        if loc_fig:
            st.plotly_chart(loc_fig, use_container_width=True)

        # Replacement Assets
        if df_expired is not None and not df_expired.empty:
            st.markdown("---")
            st.markdown('<div class="section-header">Assets Marked for Replacement</div>', unsafe_allow_html=True)
            st.dataframe(df_expired, use_container_width=True, hide_index=True)

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
        
        st.info("""
        **Your Data Security**  
        - Files are NOT stored on any server
        - Processing happens locally in your browser
        - Data stays on your computer only
        """)
    
    else:
        st.info("Sila muat naik fail Excel anda untuk bermula.")
        
        st.markdown("### Cara Guna Dashboard Ini")
        st.markdown("""
        Dashboard ini membaca nama asal kolum terus dari fail Excel anda.

        #### Ciri-ciri Utama
        - Auto-detect jenis aset
        - Semua kolum asal dipaparkan
        - Pecahan mengikut Rantau
        - UI bersih dan profesional
        - Penapisan pintar

        #### Kolum yang Diperlukan
        
        **Aset Workstation:**
        - `Model` (Wajib)
        - `Workstation Type`, `Warranty Expiry`, `Place` (Opsyenal)

        **Aset Mobile:**
        - `Product` (Wajib)
        - `Product Type`, `Programme`, `Site` (Opsyenal)
        """)
        
        col_sample1, col_sample2, col_space2 = st.columns([2, 2, 6])
        with col_sample1:
            sample_ws = create_sample_workstation_file()
            st.download_button(
                label="Contoh Workstation",
                data=sample_ws,
                file_name="contoh_workstation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        with col_sample2:
            sample_mb = create_sample_mobile_file()
            st.download_button(
                label="Contoh Mobile",
                data=sample_mb,
                file_name="contoh_mobile.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        st.info("""
        **Keselamatan Data Anda**  
        - Fail TIDAK disimpan di server
        - Pemprosesan secara lokal
        - Data kekal di komputer anda
        """)