import streamlit as st
import pandas as pd
import re
from io import BytesIO
import plotly.express as px
import plotly.graph_objects as go

# ============ PAGE CONFIG ============
st.set_page_config(page_title="IT Asset Management Dashboard", layout="wide")
st.title("💻 IT Asset Management Dashboard")

# ============ CATEGORY MAPPING ============
category_keywords = {
    "Desktop": ["optiplex"],
    "Laptop Dell": ["latitude"],
    "Laptop Acer": ["travelmate"],
    "Toughbook": ["rugged"],
    "iPad": ["ipad"]
}

# ============ REQUIRED COLUMNS ============
required_columns = {
    "computer": "Computer",
    "workstationtype": "Workstation Type",
    "serialnumber": "Serial Number",
    "model": "Model",
    "workstationstatus": "Workstation Status",
    "assettag": "Asset Tag",
    "servicetag": "Service Tag",
    "state": "State",
    "useremployeeid": "User Employee ID",
    "user": "User",
    "useremail": "User Email",
    "userjobtitle": "User Jobtitle",
    "department": "Department",
    "location": "Location",
    "site": "Site",
    "yearofpurchase": "Year Of Purchase",
    "warrantyexpiry": "Warranty Expiry",
    "place": "Place"
}

# ============ FUNCTIONS ============

def normalize_columns(columns):
    """Normalize column names untuk matching"""
    normalized = {}
    for col in columns:
        key = re.sub(r'[^a-z0-9]', '', col.lower())
        normalized[key] = col
    return normalized

def detect_header_row(excel_file, sheet_name):
    """Auto-detect header row dalam Excel"""
    preview = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, nrows=10, engine='openpyxl')
    for i, row in preview.iterrows():
        values = row.astype(str).str.lower().tolist()
        # Check for common header keywords
        keywords = ["model", "serial number", "user", "department", "asset tag", "workstation"]
        if any(keyword in v for v in values for keyword in keywords):
            return i
    return 0

def make_unique_columns(columns):
    """Buat column names yang unik"""
    seen = {}
    new_cols = []
    for col in columns:
        if col in seen:
            seen[col] += 1
            new_cols.append(f"{col}.{seen[col]}")
        else:
            seen[col] = 0
            new_cols.append(col)
    return new_cols

def validate_data(df):
    """Validate data dan return issues"""
    issues = []
    
    # Check duplicate Asset Tags
    if "Asset Tag" in df.columns:
        duplicates = df[df["Asset Tag"].duplicated(keep=False) & df["Asset Tag"].notna()]
        if not duplicates.empty:
            dup_tags = duplicates["Asset Tag"].unique()
            issues.append({
                "type": "Duplicate Asset Tags",
                "count": len(dup_tags),
                "details": f"Found {len(dup_tags)} duplicate asset tags",
                "severity": "high"
            })
    
    # Check duplicate Serial Numbers
    if "Serial Number" in df.columns:
        duplicates = df[df["Serial Number"].duplicated(keep=False) & df["Serial Number"].notna()]
        if not duplicates.empty:
            dup_serials = duplicates["Serial Number"].unique()
            issues.append({
                "type": "Duplicate Serial Numbers",
                "count": len(dup_serials),
                "details": f"Found {len(dup_serials)} duplicate serial numbers",
                "severity": "high"
            })
    
    # Check duplicate Service Tags
    if "Service Tag" in df.columns:
        duplicates = df[df["Service Tag"].duplicated(keep=False) & df["Service Tag"].notna()]
        if not duplicates.empty:
            dup_service = duplicates["Service Tag"].unique()
            issues.append({
                "type": "Duplicate Service Tags",
                "count": len(dup_service),
                "details": f"Found {len(dup_service)} duplicate service tags",
                "severity": "high"
            })
    
    # Check missing User info
    if "User" in df.columns:
        missing_users = df[df["User"].isna() | (df["User"] == "")].shape[0]
        if missing_users > 0:
            issues.append({
                "type": "Missing User Assignment",
                "count": missing_users,
                "details": f"{missing_users} assets without assigned users",
                "severity": "medium"
            })
    
    # Check invalid email format
    if "User Email" in df.columns:
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        invalid_emails = df[df["User Email"].notna() & 
                           ~df["User Email"].astype(str).str.match(email_pattern)]
        if not invalid_emails.empty:
            issues.append({
                "type": "Invalid Email Format",
                "count": len(invalid_emails),
                "details": f"{len(invalid_emails)} invalid email addresses",
                "severity": "low"
            })
    
    # Check missing Department
    if "Department" in df.columns:
        missing_dept = df[df["Department"].isna() | (df["Department"] == "")].shape[0]
        if missing_dept > 0:
            issues.append({
                "type": "Missing Department",
                "count": missing_dept,
                "details": f"{missing_dept} assets without department",
                "severity": "medium"
            })
    
    # Check missing Location
    if "Location" in df.columns:
        missing_loc = df[df["Location"].isna() | (df["Location"] == "")].shape[0]
        if missing_loc > 0:
            issues.append({
                "type": "Missing Location",
                "count": missing_loc,
                "details": f"{missing_loc} assets without location",
                "severity": "medium"
            })
    
    # Check missing Model
    if "Model" in df.columns:
        missing_model = df[df["Model"].isna() | (df["Model"] == "")].shape[0]
        if missing_model > 0:
            issues.append({
                "type": "Missing Model",
                "count": missing_model,
                "details": f"{missing_model} assets without model information",
                "severity": "high"
            })
    
    return issues

def show_validation_issues(issues):
    """Display validation issues"""
    if not issues:
        st.success("✅ No data validation issues found!")
        return
    
    st.warning(f"⚠️ Found {len(issues)} validation issue(s)")
    
    for issue in issues:
        severity_emoji = "🔴" if issue["severity"] == "high" else "🟡" if issue["severity"] == "medium" else "🟢"
        with st.expander(f"{severity_emoji} {issue['type']} ({issue['count']})"):
            st.write(issue['details'])

def export_to_excel(df, filename="asset_data.xlsx"):
    """Export dataframe to Excel"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Assets')
    output.seek(0)
    return output

def local_css():
    """Inject CSS untuk custom cards dan table styling"""
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
        </style>
    """, unsafe_allow_html=True)

def sidebar_controls(df):
    """Sidebar untuk filters dan controls"""
    st.sidebar.markdown("## ⚙️ Asset Controls")

    # Category filter
    category_filter = st.sidebar.multiselect(
        "📂 Category", 
        df["Category"].unique(), 
        default=df["Category"].unique()
    )
    
    # Workstation Type filter
    type_filter = []
    if "Workstation Type" in df.columns:
        type_filter = st.sidebar.multiselect(
            "💻 Workstation Type", 
            df["Workstation Type"].unique()
        )
    
    # Site filter
    site_filter = []
    if "Site" in df.columns:
        site_filter = st.sidebar.multiselect(
            "🏢 Site", 
            df["Site"].unique()
        )
    
    # Location filter
    location_filter = []
    if "Location" in df.columns:
        location_filter = st.sidebar.multiselect(
            "📍 Location", 
            df["Location"].unique()
        )
    
    # Department filter
    dept_filter = []
    if "Department" in df.columns:
        dept_filter = st.sidebar.multiselect(
            "👥 Department", 
            df["Department"].unique()
        )
    
    # Workstation Status filter
    status_filter = []
    if "Workstation Status" in df.columns:
        status_filter = st.sidebar.multiselect(
            "📦 Workstation Status", 
            df["Workstation Status"].unique()
        )
    
    # State filter
    state_filter = []
    if "State" in df.columns:
        state_filter = st.sidebar.multiselect(
            "🔄 State", 
            df["State"].unique()
        )

    # Expired assets selection
    expired_models = st.sidebar.multiselect(
        "⚠️ Expired / Replacement (by Model)",
        options=df["Model"].unique(),
        help="Pilih assets yang sudah expired untuk replacement"
    )

    # Apply filters
    filtered_df = df.copy()
    
    if category_filter:
        filtered_df = filtered_df[filtered_df["Category"].isin(category_filter)]
    
    if type_filter:
        filtered_df = filtered_df[filtered_df["Workstation Type"].isin(type_filter)]
    
    if site_filter:
        filtered_df = filtered_df[filtered_df["Site"].isin(site_filter)]
    
    if location_filter:
        filtered_df = filtered_df[filtered_df["Location"].isin(location_filter)]
    
    if dept_filter:
        filtered_df = filtered_df[filtered_df["Department"].isin(dept_filter)]
    
    if status_filter:
        filtered_df = filtered_df[filtered_df["Workstation Status"].isin(status_filter)]
    
    if state_filter:
        filtered_df = filtered_df[filtered_df["State"].isin(state_filter)]

    # Extract expired assets
    expired_df = None
    if expired_models:
        expired_df = filtered_df[filtered_df["Model"].isin(expired_models)]

    # Search bar
    search_query = st.sidebar.text_input("🔎 Search (User / Model / Serial Number)")
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
            <div class="card baby-blue">
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
            <div class="card navy">
                ⚠️ <br/> Expired Assets <br/><h2>{expired_assets}</h2>
            </div>
        """, unsafe_allow_html=True)

    with col4:
        st.markdown(f"""
            <div class="card white">
                📊 <br/> Replacement Rate <br/><h2>{replacement_rate:.1f}%</h2>
            </div>
        """, unsafe_allow_html=True)

def show_category_metrics(df):
    """Display category breakdown metrics"""
    category_counts = df["Category"].value_counts()
    
    col_metrics = st.columns(6)
    col_metrics[0].metric("Total Asset", len(df))
    
    categories = ["Desktop", "Laptop Dell", "Laptop Acer", "Toughbook", "iPad"]
    for i, cat in enumerate(categories, start=1):
        count = category_counts.get(cat, 0)
        col_metrics[i].metric(cat, count)

def create_pie_chart(df):
    """Create interactive pie chart untuk category distribution"""
    category_counts = df["Category"].value_counts()
    
    fig = px.pie(
        values=category_counts.values,
        names=category_counts.index,
        title="Asset Distribution by Category",
        hole=0.4,
        color_discrete_sequence=px.colors.qualitative.Set3
    )
    
    fig.update_traces(
        textposition='inside',
        textinfo='percent+label',
        hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>'
    )
    
    fig.update_layout(
        showlegend=True,
        height=400,
        margin=dict(t=50, b=0, l=0, r=0)
    )
    
    return fig

def create_department_chart(df):
    """Create bar chart untuk assets by department"""
    if "Department" not in df.columns:
        return None
    
    dept_counts = df["Department"].value_counts().head(10)
    
    fig = px.bar(
        x=dept_counts.values,
        y=dept_counts.index,
        orientation='h',
        title="Top 10 Departments by Asset Count",
        labels={'x': 'Asset Count', 'y': 'Department'},
        color=dept_counts.values,
        color_continuous_scale='Blues'
    )
    
    fig.update_layout(
        showlegend=False,
        height=400,
        margin=dict(t=50, b=50, l=0, r=0)
    )
    
    return fig

def create_location_chart(df):
    """Create bar chart untuk assets by location"""
    if "Location" not in df.columns:
        return None
    
    loc_counts = df["Location"].value_counts().head(10)
    
    fig = px.bar(
        x=loc_counts.values,
        y=loc_counts.index,
        orientation='h',
        title="Top 10 Locations by Asset Count",
        labels={'x': 'Asset Count', 'y': 'Location'},
        color=loc_counts.values,
        color_continuous_scale='Greens'
    )
    
    fig.update_layout(
        showlegend=False,
        height=400,
        margin=dict(t=50, b=50, l=0, r=0)
    )
    
    return fig

# ============ MAIN APP ============

uploaded_file = st.file_uploader("📂 Upload Excel File (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Reset file pointer
        uploaded_file.seek(0)
        
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xls.sheet_names
        selected_sheet = st.sidebar.selectbox("📑 Pilih Sheet", sheet_names)

        # Auto detect header row
        uploaded_file.seek(0)  # Reset again before reading
        header_row = detect_header_row(uploaded_file, selected_sheet)
        
        # Manual override option
        st.sidebar.markdown("---")
        st.sidebar.markdown("### 🔧 Header Settings")
        use_manual = st.sidebar.checkbox("Manual Header Row Selection", value=False)
        if use_manual:
            header_row = st.sidebar.number_input("Header Row (0-based)", min_value=0, max_value=20, value=header_row)
            st.sidebar.success(f"✓ Using row {header_row} as header")

        # Read Excel
        uploaded_file.seek(0)  # Reset again
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row, engine='openpyxl')
    
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {str(e)}")
        st.warning("💡 **Troubleshooting Tips:**")
        st.markdown("""
        1. **Cuba buka file dalam Excel** dan save semula sebagai .xlsx
        2. **Check file size** - pastikan tidak terlalu besar
        3. **Remove password protection** jika ada
        4. **Copy data ke Excel baru** dan save
        5. **Pastikan file tidak corrupt** - cuba buka dalam Excel dulu
        """)
        st.stop()

    # Make unique & normalized columns
    df.columns = make_unique_columns([str(c).strip() for c in df.columns])
    colmap = normalize_columns(df.columns)

    # Map to standard column names
    new_columns = {}
    for key, std_name in required_columns.items():
        for norm_col, original_col in colmap.items():
            if key in norm_col:
                new_columns[original_col] = std_name

    df = df.rename(columns=new_columns)

    # DEBUG: Show detected columns
    with st.expander("🔧 Debug: Column Detection", expanded=False):
        st.write("**Original Columns Found:**")
        st.write(df.columns.tolist())
        st.write("**Mapped Columns:**")
        st.write(new_columns)
        st.write("**Final Columns:**")
        st.write(df.columns.tolist())
        st.write("**Normalized Column Keys:**")
        st.write(list(colmap.keys()))

    # Check if Model exists
    if "Model" not in df.columns:
        st.error("❌ Column 'Model' tidak dijumpai.")
        st.warning("💡 Sila check 'Debug: Column Detection' di atas untuk lihat column names yang dijumpai.")
        st.info("Pastikan Excel file ada column dengan nama 'Model'.")
    else:
        # Assign categories based on Model
        df["Category"] = "Other"
        for category, keywords in category_keywords.items():
            mask = df["Model"].str.lower().str.contains("|".join(keywords), na=False)
            df.loc[mask, "Category"] = category

        # Inject custom CSS
        local_css()

        # DATA VALIDATION SECTION
        st.markdown("---")
        with st.expander("🔍 Data Validation Report", expanded=False):
            issues = validate_data(df)
            show_validation_issues(issues)

        # Sidebar filters and controls
        df_filtered, df_expired = sidebar_controls(df)

        # EXPORT SECTION
        st.sidebar.markdown("---")
        st.sidebar.markdown("## 📥 Export Data")
        
        col_exp1, col_exp2 = st.sidebar.columns(2)
        
        with col_exp1:
            # Export filtered data
            excel_data = export_to_excel(df_filtered)
            st.download_button(
                label="📊 Export Filtered",
                data=excel_data,
                file_name=f"assets_filtered_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        
        with col_exp2:
            # Export expired data
            if df_expired is not None and not df_expired.empty:
                excel_expired = export_to_excel(df_expired)
                st.download_button(
                    label="⚠️ Export Expired",
                    data=excel_expired,
                    file_name=f"assets_expired_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        # Summary cards
        st.subheader("📊 Dashboard Summary")
        show_summary_cards(df_filtered, df_expired)

        # Category metrics
        st.subheader("📈 Asset Breakdown by Category")
        show_category_metrics(df_filtered)

        # VISUAL CHARTS
        st.markdown("---")
        st.subheader("📊 Visual Analytics")
        
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            # Pie chart
            pie_fig = create_pie_chart(df_filtered)
            st.plotly_chart(pie_fig, use_container_width=True)
        
        with col_chart2:
            # Department chart
            dept_fig = create_department_chart(df_filtered)
            if dept_fig:
                st.plotly_chart(dept_fig, use_container_width=True)
            else:
                st.info("Department data not available")
        
        # Location chart (full width)
        loc_fig = create_location_chart(df_filtered)
        if loc_fig:
            st.plotly_chart(loc_fig, use_container_width=True)

        # Expired assets section
        if df_expired is not None and not df_expired.empty:
            st.markdown("---")
            st.subheader("⚠️ Assets Marked for Replacement")
            st.dataframe(df_expired, use_container_width=True)

        # Full asset details table
        st.markdown("---")
        st.subheader("📋 Asset Details")
        safe_columns = list(dict.fromkeys(list(new_columns.values()) + ["Category"]))
        st.dataframe(df_filtered[safe_columns], use_container_width=True)
    
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {str(e)}")
        st.warning("💡 **Troubleshooting Tips:**")
        st.markdown("""
        1. **Cuba buka file dalam Excel** dan save semula sebagai .xlsx
        2. **Check file size** - pastikan tidak terlalu besar
        3. **Remove password protection** jika ada
        4. **Copy data ke Excel baru** dan save
        5. **Pastikan file tidak corrupt** - cuba buka dalam Excel dulu
        """)
        st.stop()

else:
    st.info("📂 Sila upload fail Excel untuk mula.")