import streamlit as st
import pandas as pd
import re
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go

# ============ PAGE CONFIG ============
st.set_page_config(page_title="IT Asset Management Dashboard", layout="wide")
st.title("💻 IT Asset Management Dashboard")

# ============ COLUMN IDENTIFIERS (for detection only, NOT for renaming) ============
WORKSTATION_IDENTIFIERS = [
    "workstation", "model", "warranty", "place"
]

MOBILE_IDENTIFIERS = [
    "product", "programme", "program"
]

# ============ FUNCTIONS ============

def normalize_text(text):
    """Normalize text untuk matching - remove special chars and lowercase"""
    return re.sub(r'[^a-z0-9]', '', str(text).lower())

def detect_asset_type(df_columns):
    """Auto-detect asset type (Workstation or Mobile) based on headers"""
    normalized = [normalize_text(col) for col in df_columns]
    
    # Count matching identifiers for each type
    workstation_score = sum(1 for identifier in WORKSTATION_IDENTIFIERS 
                           if any(identifier in norm for norm in normalized))
    mobile_score = sum(1 for identifier in MOBILE_IDENTIFIERS 
                      if any(identifier in norm for norm in normalized))
    
    if workstation_score > mobile_score:
        return "Workstation"
    else:
        return "Mobile"

def find_column(df, search_terms):
    """Find column by searching for terms (case insensitive, flexible matching)"""
    if isinstance(search_terms, str):
        search_terms = [search_terms]
    
    for col in df.columns:
        normalized_col = normalize_text(col)
        for term in search_terms:
            normalized_term = normalize_text(term)
            if normalized_term in normalized_col:
                return col
    return None

def get_model_column(df, asset_type):
    """Get the model column based on asset type"""
    if asset_type == "Workstation":
        return find_column(df, ["model"])
    else:  # Mobile
        # For mobile, "Product" is the model (not "Product Type")
        for col in df.columns:
            normalized = normalize_text(col)
            if normalized == "product":  # Exact match untuk avoid "product type"
                return col
    return None

def get_type_column(df, asset_type):
    """Get the type column based on asset type"""
    if asset_type == "Workstation":
        return find_column(df, ["workstation type", "workstationtype"])
    else:  # Mobile
        return find_column(df, ["product type", "producttype"])

def calculate_asset_age(df):
    """Calculate asset age dari Year Of Purchase"""
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
    """Get warranty status and expired assets"""
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

def show_warranty_summary(df, model_col):
    """Show warranty status summary with expandable tables"""
    if "Warranty Status" not in df.columns:
        return

    status_counts = df["Warranty Status"].value_counts()

    col1, col2, col3 = st.columns(3)

    with col1:
        expired = status_counts.get("Expired", 0)
        st.markdown(f"""
            <div class="card red">
                ⚠️ <br/> Expired Warranty <br/><h2>{expired}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        expiring = status_counts.get("Expiring Soon", 0)
        st.markdown(f"""
            <div class="card orange">
                ⏰ <br/> Expiring Soon (90 days) <br/><h2>{expiring}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col3:
        active = status_counts.get("Active", 0)
        st.markdown(f"""
            <div class="card green">
                ✅ <br/> Active Warranty <br/><h2>{active}</h2>
            </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    serial_col = find_column(df, ["serial number", "serialnumber"])
    user_col = find_column(df, ["user"])
    dept_col = find_column(df, ["department", "user department"])
    location_col = find_column(df, ["location"])
    warranty_col = find_column(df, ["warranty expiry", "warrantyexpiry"])

    # Expired warranty list
    expired_df = df[df["Warranty Status"] == "Expired"]
    if not expired_df.empty:
        with st.expander(f"⚠️ Expired Warranty Assets ({len(expired_df)})", expanded=False):
            display_cols = [c for c in [model_col, serial_col, user_col, dept_col, location_col, warranty_col] if c]
            st.dataframe(expired_df[display_cols], use_container_width=True, hide_index=True)

    # Expiring soon list
    expiring_df = df[df["Warranty Status"] == "Expiring Soon"]
    if not expiring_df.empty:
        with st.expander(f"⏰ Expiring Soon Assets ({len(expiring_df)})", expanded=False):
            display_cols = [c for c in [model_col, serial_col, user_col, dept_col, location_col, warranty_col] if c]
            st.dataframe(expiring_df[display_cols], use_container_width=True, hide_index=True)

    # Active warranty list
    active_df = df[df["Warranty Status"] == "Active"]
    if not active_df.empty:
        with st.expander(f"✅ Active Warranty Assets ({len(active_df)})", expanded=False):
            display_cols = [c for c in [model_col, serial_col, user_col, dept_col, location_col, warranty_col] if c]
            st.dataframe(active_df[display_cols], use_container_width=True, hide_index=True)

def show_asset_age_summary(df):
    """Show asset age distribution"""
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
        st.metric("Average Age", f"{avg_age:.1f} years" if not pd.isna(avg_age) else "N/A")
    with col2:
        st.metric("New (0-1yr)", age_counts.get("New (0-1 year)", 0))
    with col3:
        st.metric("Active (1-3yr)", age_counts.get("Active (1-3 years)", 0))
    with col4:
        st.metric("Aging (3-5yr)", age_counts.get("Aging (3-5 years)", 0))
    with col5:
        st.metric("Old (5+yr)", age_counts.get("Old (5+ years)", 0))

def detect_header_row(excel_file, sheet_name):
    """Auto-detect header row dalam Excel"""
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

def validate_data(df, asset_type, model_col):
    """Validate data dan return issues"""
    issues = []

    asset_tag_col = find_column(df, ["asset tag", "assettag"])
    serial_col = find_column(df, ["serial number", "serialnumber"])
    user_col = find_column(df, ["user"])
    email_col = find_column(df, ["email"])
    dept_col = find_column(df, ["department", "user department"])
    location_col = find_column(df, ["location"])

    # Check duplicate Asset Tags
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

    # Check duplicate Serial Numbers
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

    # Check missing User info
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

    # Check invalid email format
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

    # Check missing Department
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

    # Check missing Location
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

    # Check expired warranties (only for Workstation)
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
    """Display validation issues"""
    if not issues:
        st.success("✅ No data validation issues found!")
        return

    st.warning(f"⚠️ Found {len(issues)} validation issue(s)")

    for issue in issues:
        severity_emoji = "🔴" if issue["severity"] == "high" else "🟡" if issue["severity"] == "medium" else "🟢"
        with st.expander(f"{severity_emoji} {issue['type']} ({issue['count']})", expanded=False):
            st.write(issue['details'])
            if "data" in issue and not issue["data"].empty:
                st.markdown("**Affected Assets:**")
                st.dataframe(issue["data"], use_container_width=True, hide_index=True)

def export_to_excel(df, filename="asset_data.xlsx"):
    """Export dataframe to Excel"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Assets')
    output.seek(0)
    return output

def create_sample_workstation_file():
    """Create sample Workstation Excel file"""
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
    """Create sample Mobile Excel file"""
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

def local_css():
    """Inject CSS untuk custom cards"""
    st.markdown("""
        <style>
        .card {
            border-radius: 12px;
            padding: 20px;
            color: white;
            text-align: center;
            font-weight: bold;
            margin-bottom: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .baby-blue {background-color: #89CFF0;}
        .green {background-color: #2ecc71;}
        .navy {background-color: #001F54;}
        .white {background-color: #FFFFFF; color: black;}
        .red {background-color: #e74c3c;}
        .orange {background-color: #f39c12;}
        .purple {background-color: #9b59b6;}
        
        .lang-toggle {
            display: flex;
            gap: 10px;
            margin-bottom: 20px;
        }
        .lang-btn {
            padding: 8px 16px;
            border-radius: 6px;
            border: 2px solid #ddd;
            background: white;
            cursor: pointer;
            font-weight: 500;
        }
        .lang-btn.active {
            background: #001F54;
            color: white;
            border-color: #001F54;
        }
        </style>
    """, unsafe_allow_html=True)

def sidebar_controls(df, asset_type, model_col, type_col):
    """Sidebar untuk filters dan controls"""
    st.sidebar.markdown("## ⚙️ Asset Controls")

    # Model/Product filter
    model_filter = []
    if model_col:
        if asset_type == "Workstation":
            filter_label = f"📦 {model_col}"
        else:
            filter_label = f"📦 {model_col} (Model)"
        model_filter = st.sidebar.multiselect(filter_label, df[model_col].unique())

    # Type filter
    type_filter = []
    if type_col:
        if asset_type == "Workstation":
            type_label = f"💻 {type_col}"
        else:
            type_label = f"📱 {type_col} (Category)"
        type_filter = st.sidebar.multiselect(type_label, df[type_col].unique())
    
    # Site filter
    site_col = find_column(df, ["site", "user site", "usersite"])
    site_filter = []
    if site_col:
        site_filter = st.sidebar.multiselect(f"🏢 {site_col}", df[site_col].unique())
    
    # Location filter
    location_col = find_column(df, ["location"])
    location_filter = []
    if location_col:
        location_filter = st.sidebar.multiselect(f"📍 {location_col}", df[location_col].unique())
    
    # Department filter
    dept_col = find_column(df, ["department", "user department"])
    dept_filter = []
    if dept_col:
        dept_filter = st.sidebar.multiselect(f"👥 {dept_col}", df[dept_col].unique())
    
    # Status filter (only for Workstation)
    status_col = find_column(df, ["workstation status", "workstationstatus"])
    status_filter = []
    if asset_type == "Workstation" and status_col:
        status_filter = st.sidebar.multiselect(f"📦 {status_col}", df[status_col].unique())
    
    # State filter
    state_col = find_column(df, ["state"])
    state_filter = []
    if state_col:
        state_filter = st.sidebar.multiselect(f"🔄 {state_col}", df[state_col].unique())

    # Place filter (only for Workstation)
    place_col = find_column(df, ["place"])
    place_filter = []
    if asset_type == "Workstation" and place_col:
        place_filter = st.sidebar.multiselect(f"🌍 {place_col}", df[place_col].unique())

    # Programme filter (only for Mobile)
    programme_col = find_column(df, ["programme", "program"])
    programme_filter = []
    if asset_type == "Mobile" and programme_col:
        programme_filter = st.sidebar.multiselect(f"📱 {programme_col}", df[programme_col].unique())

    # Expired assets selection
    expired_models = []
    if model_col:
        label = f"⚠️ Expired / Replacement (by {model_col})"
        expired_models = st.sidebar.multiselect(label, options=df[model_col].unique(),
                                                help="Pilih assets untuk replacement")

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

    # Extract expired assets
    expired_df = None
    if expired_models and model_col:
        expired_df = filtered_df[filtered_df[model_col].isin(expired_models)]

    # Search bar
    search_query = st.sidebar.text_input("🔎 Search (any field)")
    if search_query:
        filtered_df = filtered_df[filtered_df.apply(
            lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1
        )]

    return filtered_df, expired_df

def show_summary_cards(df, df_expired=None):
    """Display summary metrics cards"""
    total_assets = len(df)
    expired_assets = len(df_expired) if df_expired is not None else 0
    active_assets = total_assets - expired_assets
    replacement_rate = (expired_assets / total_assets * 100) if total_assets > 0 else 0

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(f"""
            <div class="card green">
                📦 <br/> Total Assets <br/><h2>{total_assets}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown(f"""
            <div class="card green">
                ✅ <br/> Active Assets <br/><h2>{active_assets}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
            <div class="card green">
                ⚠️ <br/> Expired Assets <br/><h2>{expired_assets}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
            <div class="card green">
                📊 <br/> Replacement Rate <br/><h2>{replacement_rate:.1f}%</h2>
            </div>
        """, unsafe_allow_html=True)

def show_category_metrics(df, model_col):
    """Display model breakdown metrics"""
    if not model_col:
        st.warning("Model column not found")
        return

    model_counts = df[model_col].value_counts().sort_values(ascending=False)

    st.markdown(f"### Unit Breakdown by {model_col}")

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

def show_type_cards(df, type_col, asset_type):
    """Display type breakdown as metric cards"""
    if not type_col:
        return

    if asset_type == "Workstation":
        title = f"### 💻 {type_col} Statistics"
    else:
        title = f"### 📱 {type_col} Statistics (Categories)"

    st.markdown(title)

    type_counts = df[type_col].value_counts().sort_values(ascending=False)
    num_types = len(type_counts)
    cols_per_row = min(4, num_types)
    cols = st.columns(cols_per_row)
    colors = ["baby-blue"]

    for idx, (wtype, count) in enumerate(type_counts.items()):
        col_idx = idx % cols_per_row
        color = colors[idx % len(colors)]

        with cols[col_idx]:
            emoji = "💻" if asset_type == "Workstation" else "📱"
            st.markdown(f"""
                <div class="card {color}">
                    {emoji} <br/> {wtype} <br/><h2>{count}</h2>
                </div>
            """, unsafe_allow_html=True)

        if (idx + 1) % cols_per_row == 0 and (idx + 1) < num_types:
            cols = st.columns(cols_per_row)

def create_pie_chart(df, model_col):
    """Create interactive pie chart"""
    if not model_col:
        return None

    model_counts = df[model_col].value_counts()

    fig = px.pie(
        values=model_counts.values,
        names=model_counts.index,
        title=f"Asset Distribution by {model_col}",
        hole=0.4,
        color_discrete_sequence=px.colors.qualitative.Set3
    )

    fig.update_traces(
        textposition='inside',
        textinfo='percent+label',
        hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
    )

    fig.update_layout(showlegend=True, height=400, margin=dict(t=50, b=0, l=0, r=0))
    return fig

def create_department_chart(df, dept_col):
    """Create bar chart untuk assets by department"""
    if not dept_col:
        return None
    
    dept_counts = df[dept_col].value_counts().head(10)
    
    fig = px.bar(
        x=dept_counts.values,
        y=dept_counts.index,
        orientation='h',
        title=f"Top 10 {dept_col} by Asset Count",
        labels={'x': 'Asset Count', 'y': dept_col},
        color=dept_counts.values,
        color_continuous_scale='Blues'
    )
    
    fig.update_layout(showlegend=False, height=400, margin=dict(t=50, b=50, l=0, r=0))
    return fig

def create_location_chart(df, location_col):
    """Create bar chart untuk assets by location"""
    if not location_col:
        return None
    
    loc_counts = df[location_col].value_counts().head(10)
    
    fig = px.bar(
        x=loc_counts.values,
        y=loc_counts.index,
        orientation='h',
        title=f"Top 10 {location_col} by Asset Count",
        labels={'x': 'Asset Count', 'y': location_col},
        color=loc_counts.values,
        color_continuous_scale='Greens'
    )
    
    fig.update_layout(showlegend=False, height=400, margin=dict(t=50, b=50, l=0, r=0))
    return fig

# ============ MAIN APP ============

uploaded_file = st.file_uploader("📂 Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        uploaded_file.seek(0)
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xls.sheet_names
        selected_sheet = st.sidebar.selectbox("📑 Pilih Sheet", sheet_names)

        # Auto detect header row
        uploaded_file.seek(0)
        header_row = detect_header_row(uploaded_file, selected_sheet)
        
        # Manual override option
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 🔧 Header Settings")
        use_manual = st.sidebar.checkbox("Manual Header Row Selection", value=False)
        if use_manual:
            header_row = st.sidebar.number_input("Header Row (0-based)", min_value=0, max_value=20, value=header_row)
            st.sidebar.success(f"✓ Using row {header_row} as header")

        # Read Excel - USE ORIGINAL COLUMN NAMES
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row, engine='openpyxl')
        
        # Clean column names - only strip whitespace, keep original names
        df.columns = [str(c).strip() for c in df.columns]
        
        # Remove duplicate columns if any (keep first occurrence)
        df = df.loc[:, ~df.columns.duplicated(keep='first')]
        
        # AUTO-DETECT ASSET TYPE (no renaming, just detection)
        asset_type = detect_asset_type(df.columns)
        
        # Display detected asset type
        if asset_type == "Workstation":
            st.sidebar.success(f"✅ Detected: **{asset_type}** Assets 💻")
        else:
            st.sidebar.success(f"✅ Detected: **{asset_type}** Assets 📱")
        
        # Show original columns found
        with st.sidebar.expander("📋 Excel Columns Found", expanded=False):
            st.write(f"**Total columns:** {len(df.columns)}")
            for idx, col in enumerate(df.columns, 1):
                st.text(f"{idx}. {col}")
        
        # Find key columns (without renaming)
        model_col = get_model_column(df, asset_type)
        type_col = get_type_column(df, asset_type)
        
        if not model_col:
            st.error(f"❌ Model column tidak dijumpai dalam file Excel.")
            if asset_type == "Workstation":
                st.info("📋 Pastikan Excel ada column 'Model'")
            else:
                st.info("📋 Pastikan Excel ada column 'Product' (untuk model device)")
            st.stop()
        
        # Calculate asset age
        df = calculate_asset_age(df)
        
        # Get warranty status (only for Workstation)
        expired_warranty_df = None
        if asset_type == "Workstation":
            df, expired_warranty_df = get_warranty_status(df)

        # Inject custom CSS
        local_css()

        # DATA VALIDATION SECTION
        st.markdown("---")
        with st.expander("🔍 Data Validation Report", expanded=False):
            issues = validate_data(df, asset_type, model_col)
            show_validation_issues(issues)

        # Sidebar filters and controls
        df_filtered, df_expired = sidebar_controls(df, asset_type, model_col, type_col)

        # EXPORT SECTION
        st.sidebar.markdown("---")
        st.sidebar.markdown("## 📥 Export Data")
        
        if asset_type == "Workstation":
            col_exp1, col_exp2, col_exp3 = st.sidebar.columns(3)
        else:
            col_exp1, col_exp2 = st.sidebar.columns(2)
        
        with col_exp1:
            excel_data = export_to_excel(df_filtered)
            st.download_button(
                label="📊 All",
                data=excel_data,
                file_name=f"{asset_type.lower()}_assets_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Export filtered data"
            )
        
        with col_exp2:
            if df_expired is not None and not df_expired.empty:
                excel_expired = export_to_excel(df_expired)
                st.download_button(
                    label="⚠️ Expired",
                    data=excel_expired,
                    file_name=f"{asset_type.lower()}_expired_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    help="Export expired assets"
                )
        
        if asset_type == "Workstation":
            with col_exp3:
                if expired_warranty_df is not None and not expired_warranty_df.empty:
                    excel_warranty = export_to_excel(expired_warranty_df)
                    st.download_button(
                        label="🔴 Warranty",
                        data=excel_warranty,
                        file_name=f"warranty_expired_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        help="Export expired warranty"
                    )
        
        # HELP & SUPPORT SECTION
        st.sidebar.markdown("---")
        st.sidebar.markdown("## 💡 Help & Support")
        
        with st.sidebar.expander("🆘 Troubleshooting", expanded=False):
            st.markdown("""
            **Common Issues:**
            
            🔴 **Model column not found**
            - Ensure Excel has 'Model' (Workstation) or 'Product' (Mobile) column
            
            🔴 **Error reading file**
            - Save file as .xlsx format
            - Remove password protection
            - Close file in Excel before upload
            
            🔴 **Wrong asset type detected**
            - Check column names match expected format
            - Use manual header row selection if needed
            
            🔴 **Data not showing correctly**
            - Verify header row is correct
            - Check for merged cells in Excel
            - Ensure data starts immediately after header
            """)
        
        with st.sidebar.expander("📧 Contact Support", expanded=False):
            st.markdown("""
            **Need Help?**
            
            📧 Email: khalis.abdrahim@airselangor.com  
            
            **Response Time:**  
            Mon-Fri: Within 24 hours  
            Weekend: Within 48 hours
            """)
        
        # FOOTER / VERSION INFO
        st.sidebar.markdown("---")
        st.sidebar.markdown("""
            <div style='text-align: center; color: #666; font-size: 0.85em;'>
                <strong>IT Asset Dashboard</strong><br/>
                Version 2.0.1<br/>
                Last Updated: Oct 2025<br/>
                <br/>
            </div>
        """, unsafe_allow_html=True)

        # Summary cards
        st.subheader("📊 Dashboard Summary")
        show_summary_cards(df_filtered, df_expired)

        # TYPE SECTION
        if type_col:
            st.markdown("---")
            show_type_cards(df_filtered, type_col, asset_type)

        # WARRANTY STATUS SECTION (only for Workstation)
        if asset_type == "Workstation" and "Warranty Status" in df_filtered.columns:
            st.markdown("---")
            st.subheader("🛡️ Warranty Status")
            show_warranty_summary(df_filtered, model_col)

        # ASSET AGE SECTION
        if "Asset Age" in df_filtered.columns:
            st.markdown("---")
            st.subheader("📅 Asset Age Analysis")
            show_asset_age_summary(df_filtered)

        # Category metrics
        st.markdown("---")
        show_category_metrics(df_filtered, model_col)

        # VISUAL CHARTS
        st.markdown("---")
        st.subheader("📊 Visual Analytics")
        
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

        # Expired assets section
        if df_expired is not None and not df_expired.empty:
            st.markdown("---")
            st.subheader("⚠️ Assets Marked for Replacement")
            st.dataframe(df_expired, use_container_width=True, hide_index=True)

        # Full asset details table - SHOW ALL ORIGINAL COLUMNS
        st.markdown("---")
        st.subheader("📋 Asset Details")
        
        # Calculate columns excluding Year of Purchase for display
        year_col = find_column(df_filtered, ["year of purchase", "yearofpurchase"])
        display_columns = [col for col in df_filtered.columns if col != year_col]
        
        st.info(f"📊 Displaying ALL {len(display_columns)} columns from Excel file (excluding Year Of Purchase used for calculations)")
        
        # Display with better height
        st.dataframe(df_filtered[display_columns], use_container_width=True, height=600)

    except Exception as e:
        st.error(f"❌ Error reading Excel file: {str(e)}")
        st.warning("💡 **Troubleshooting Tips:**")
        st.markdown("""
        1. **Pastikan file format .xlsx (Excel)**
        2. **File mesti ada header row dengan column names yang jelas**
        3. **Cuba buka file dalam Excel dan save semula**
        4. **Remove password protection jika ada**
        5. **Pastikan file tidak corrupt**
        
        ### Expected Columns:
        
        **Workstation Assets:**
        - Model (required)
        - Workstation Type
        - Serial Number
        - User, Department, Location
        - Warranty Expiry (optional)
        
        **Mobile Assets:**
        - Product (required - ini adalah model device)
        - Product Type (category: Tablet, Phone, etc)
        - Serial Number
        - User, Department, Location
        - Programme (optional)
        """)

else:
    # Initialize language selection in session state
    if 'language' not in st.session_state:
        st.session_state.language = 'EN'
    
    # Language toggle buttons
    col_lang1, col_lang2, col_space = st.columns([1, 1, 8])
    with col_lang1:
        if st.button("🇬🇧 English", use_container_width=True, 
                     type="primary" if st.session_state.language == 'EN' else "secondary"):
            st.session_state.language = 'EN'
            st.rerun()
    with col_lang2:
        if st.button("🇲🇾 Bahasa", use_container_width=True,
                     type="primary" if st.session_state.language == 'MY' else "secondary"):
            st.session_state.language = 'MY'
            st.rerun()
    
    st.markdown("---")
    
    # Content based on selected language
    if st.session_state.language == 'EN':
        st.info("📂 Please upload your Excel file to get started.")
        
        st.markdown("### 🚀 How to Use This Dashboard")
        st.markdown("""
        This dashboard reads **original column names** directly from your Excel file.  
        No complex mapping needed – what's in your Excel is what appears in the dashboard.

        #### 🔑 Key Features
        - **Auto-detection** of asset type (Workstation or Mobile)
        - **All original columns** are displayed (no data loss)
        - **Flexible** – works with various column name formats
        - **Smart filtering** and search capabilities

        #### 📋 Required Columns
        
        **For Workstation Assets:**
        - `Model` (Required)
        - `Workstation Type`, `Warranty Expiry`, `Place` (Optional)

        **For Mobile Assets:**
        - `Product` (Required – device model name)
        - `Product Type`, `Programme` (Optional)

        #### 📥 Download Sample Files
        """)
        
        col_sample1, col_sample2, col_space2 = st.columns([2, 2, 6])
        with col_sample1:
            sample_ws = create_sample_workstation_file()
            st.download_button(
                label="💻 Workstation Sample",
                data=sample_ws,
                file_name="sample_workstation_assets.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download sample Workstation Excel file",
                use_container_width=True
            )
        with col_sample2:
            sample_mb = create_sample_mobile_file()
            st.download_button(
                label="📱 Mobile Sample",
                data=sample_mb,
                file_name="sample_mobile_assets.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Download sample Mobile Excel file",
                use_container_width=True
            )
        
        st.info("""
        🔒 **Your Data Security**  
        - Files are **NOT stored** on any server
        - Processing happens **locally in your browser**
        - Data stays on **your computer only**
        
        ℹ️ **Note:** If you refresh the page or close the browser, you'll need to upload the file again.
        """)
    
    else:  # Bahasa Malaysia
        st.info("📂 Sila muat naik fail Excel anda untuk bermula.")
        
        st.markdown("### 🚀 Cara Guna Dashboard Ini")
        st.markdown("""
        Dashboard ini membaca **nama asal kolum** terus dari fail Excel anda.  
        Tiada proses mapping rumit – apa yang ada dalam Excel, itu yang muncul di dashboard.

        #### 🔑 Ciri-ciri Utama
        - **Auto-detect** jenis aset (Workstation atau Mobile)
        - **Semua kolum asal** dipaparkan (tiada data hilang)
        - **Fleksibel** – berfungsi dengan pelbagai format nama kolum
        - Kemudahan **penapisan** dan carian yang pintar

        #### 📋 Kolum yang Diperlukan
        
        **Untuk Aset Workstation:**
        - `Model` (Wajib)
        - `Workstation Type`, `Warranty Expiry`, `Place` (Opsyenal)

        **Untuk Aset Mobile:**
        - `Product` (Wajib – nama model peranti)
        - `Product Type`, `Programme` (Opsyenal)

        #### 📥 Muat Turun Fail Contoh
        """)
        
        col_sample1, col_sample2, col_space2 = st.columns([2, 2, 6])
        with col_sample1:
            sample_ws = create_sample_workstation_file()
            st.download_button(
                label="💻 Contoh Workstation",
                data=sample_ws,
                file_name="contoh_aset_workstation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Muat turun fail Excel contoh untuk Workstation",
                use_container_width=True
            )
        with col_sample2:
            sample_mb = create_sample_mobile_file()
            st.download_button(
                label="📱 Contoh Mobile",
                data=sample_mb,
                file_name="contoh_aset_mobile.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Muat turun fail Excel contoh untuk Mobile",
                use_container_width=True
            )
        
        st.info("""
        🔒 **Keselamatan Data Anda**  
        - Fail **TIDAK disimpan** di mana-mana server
        - Pemprosesan berlaku **secara lokal dalam browser anda**
        - Data kekal di **komputer anda sahaja**
        
        ℹ️ **Nota:** Jika anda refresh halaman atau tutup browser, anda perlu muat naik fail semula.
        """)
