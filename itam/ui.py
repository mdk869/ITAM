import pandas as pd
import plotly.express as px
import streamlit as st

from itam.export import export_to_excel


DISPLAY_LABELS = {
    "asset_type": "Asset Type",
    "source_asset_subtype": "Asset Type Detail",
    "serial_number": "Serial Number",
    "asset_tag": "Asset Tag",
    "model": "Model / Product",
    "state": "Source State",
    "user": "User",
    "employee_id": "Employee ID",
    "email": "Email",
    "job_title": "Job Title",
    "department": "Department",
    "location": "Location",
    "site": "Site",
    "purchase_year": "Year Of Purchase",
    "warranty_expiry": "Warranty Expiry",
    "programme": "Programme",
    "place": "Place",
    "workstation_status": "Workstation Status",
    "imei": "IMEI",
    "sim_number": "SIM Number",
    "Asset Age": "Asset Age",
    "ITAM Lifecycle Status": "Lifecycle",
    "Warranty Status": "Warranty Status",
    "ITAM Finding Count": "Findings",
    "ITAM Highest Severity": "Severity",
    "ITAM Review Required": "Review Required",
    "ITAM Audit Findings": "Audit Notes",
    "ITAM State Review Required": "State Review",
    "ITAM Replacement Candidate": "Replacement Candidate",
    "ITAM Duplicate Asset Tag": "Duplicate Asset Tag",
    "ITAM Duplicate Serial": "Duplicate Serial",
    "ITAM Duplicate IMEI": "Duplicate IMEI",
    "ITAM Missing Core Identity": "Missing Core Identity",
}

EXPORT_PRESETS = {
    "Standard Asset View": [
        "asset_type", "asset_tag", "serial_number", "model", "state", "user",
        "department", "site", "location", "purchase_year", "warranty_expiry",
        "ITAM Lifecycle Status",
    ],
    "Audit Findings": [
        "asset_type", "asset_tag", "serial_number", "model", "state", "user",
        "site", "purchase_year", "Asset Age", "ITAM Lifecycle Status",
        "Warranty Status", "ITAM Finding Count", "ITAM Highest Severity",
        "ITAM Review Required", "ITAM Audit Findings",
    ],
    "Replacement Planning": [
        "asset_type", "asset_tag", "serial_number", "model", "site", "department",
        "purchase_year", "Asset Age", "state", "workstation_status", "programme",
        "ITAM Lifecycle Status", "ITAM Replacement Candidate",
    ],
    "Full Dataset": None,
}

SEARCH_COLUMNS = [
    "asset_tag", "serial_number", "model", "user", "employee_id", "email",
    "location", "site", "imei", "sim_number", "programme", "place",
]


def format_audit_note(value):
    """Return readable audit-note text without changing stored audit findings."""
    if pd.isna(value):
        return value
    text = str(value)
    if not text.strip():
        return text
    return text.replace("ITAM lifecycle", "Lifecycle").replace("source State", "Source State")


def _values(df, column):
    if column not in df.columns:
        return []
    return sorted(df[column].dropna().astype(str).replace("<NA>", "").loc[lambda values: values.ne("")].unique())


def _metric_row(metrics):
    columns = st.columns(len(metrics))
    for column, (label, value) in zip(columns, metrics):
        with column:
            normalized_label = label.casefold()
            accent = (
                "danger" if any(term in normalized_label for term in ["expired", "high severity"])
                else "warning" if any(term in normalized_label for term in ["review", "aging"])
                else "success" if any(term in normalized_label for term in ["within", "active", "new"])
                else "info"
            )
            st.markdown(
                f'<div class="al-kpi al-kpi-{accent}">'
                f'<div class="al-kpi-label">{label}</div>'
                f'<div class="al-kpi-value">{value}</div>'
                "</div>",
                unsafe_allow_html=True,
            )


def _filtered_by_selections(df, selections):
    filtered = df.copy()
    for column, selected in selections.items():
        if selected and column in filtered.columns:
            filtered = filtered[filtered[column].astype(str).isin(selected)]
    return filtered


def make_display_dataframe(df, columns=None):
    """Copy a frame and apply unique human-readable display labels."""
    selected_columns = list(df.columns) if columns is None else [
        column for column in columns if column in df.columns
    ]
    display_df = df.loc[:, selected_columns].copy()
    if "ITAM Audit Findings" in display_df.columns:
        display_df["ITAM Audit Findings"] = display_df["ITAM Audit Findings"].map(format_audit_note)
    labels = []
    used_labels = set()
    for column in display_df.columns:
        label = DISPLAY_LABELS.get(column, str(column))
        if label in used_labels and column in DISPLAY_LABELS:
            label = f"{label} (Canonical)"
        if label in used_labels:
            suffix = 2
            base_label = label
            while label in used_labels:
                label = f"{base_label} ({suffix})"
                suffix += 1
        labels.append(label)
        used_labels.add(label)
    display_df.columns = labels
    if not display_df.columns.is_unique:
        raise ValueError("Display column labels must be unique")
    return display_df


def resolve_export_columns(df, preset, custom_columns=None):
    """Resolve a preset to available internal columns without dropping data silently."""
    if preset == "Custom":
        requested = custom_columns or []
    elif preset == "Full Dataset":
        requested = list(df.columns)
    else:
        requested = EXPORT_PRESETS.get(preset, [])
    if preset == "Full Dataset":
        return requested
    return [
        column for column in requested
        if column in df.columns and df[column].notna().any()
    ]


def prepare_export_dataframe(df, columns):
    """Prepare selected export columns using the same unique display labels as tables."""
    if not columns:
        return None
    return make_display_dataframe(df, columns)


def _display_frame(df, columns=None):
    return make_display_dataframe(df, columns)


def _download(df, label, filename, key):
    st.download_button(
        label=label,
        data=export_to_excel(make_display_dataframe(df)),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
    )


def render_export_panel(df, scope_label, filename, key):
    """Render a compact, page-local custom Excel export workflow."""
    st.subheader("Export Data")
    st.caption(scope_label)
    preset_options = ["Standard Asset View", "Audit Findings", "Replacement Planning", "Full Dataset", "Custom"]
    preset = st.selectbox("Column Preset", preset_options, key=f"{key}-preset")
    available_display = make_display_dataframe(df)
    display_to_internal = dict(zip(available_display.columns, df.columns))
    default_internal = resolve_export_columns(df, preset)
    default_display = [
        label for label, column in display_to_internal.items() if column in default_internal
    ]
    columns_key = f"{key}-columns-select-{preset}"
    selected_display = st.multiselect(
        "Columns", list(available_display.columns), default=default_display,
        key=columns_key,
    )
    select_left, select_right = st.columns(2)
    with select_left:
        if st.button("Select All", key=f"{key}-all", use_container_width=True):
            st.session_state[columns_key] = list(available_display.columns)
            st.rerun()
    with select_right:
        if st.button("Clear", key=f"{key}-clear", use_container_width=True):
            st.session_state[columns_key] = []
            st.rerun()
    selected_internal = [display_to_internal[label] for label in selected_display]
    if not selected_internal:
        st.warning("Select at least one column to export.")
        return
    if df.empty:
        st.info("There are no rows in the current result to export.")
        return
    export_df = prepare_export_dataframe(df, selected_internal)
    if export_df is None or not export_df.columns.is_unique:
        st.error("The selected export columns could not be prepared safely.")
        return
    st.download_button(
        "Export Excel", export_to_excel(export_df), filename,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"{key}-download",
        use_container_width=True,
    )


def _bar_chart(df, column, title, key):
    if column not in df.columns:
        return
    counts = df[column].dropna().astype(str).value_counts().head(10).sort_values()
    if counts.empty:
        return
    figure = px.bar(
        x=counts.values, y=counts.index, orientation="h", title=title,
        labels={"x": "Assets", "y": title.removeprefix("Top ")},
    )
    figure.update_layout(height=360, margin=dict(t=55, b=20, l=10, r=10), showlegend=False)
    st.plotly_chart(figure, use_container_width=True, key=key)


def _lifecycle_metrics(df):
    labels = [
        ("NEW (0-1 YR)", "New"),
        ("ACTIVE (2-3 YR)", "Active"),
        ("AGING (4-5 YR)", "Aging"),
        ("EXPIRED (>5 YR)", "Expired"),
    ]
    _metric_row([(label, int(df["ITAM Lifecycle Status"].eq(value).sum())) for label, value in labels])


def render_overview(df, asset_type):
    st.title("Overview")
    st.caption(f"Management summary for {asset_type} assets")
    total = len(df)
    expired = int(df["ITAM Replacement Candidate"].sum())
    within = int(df["Asset Age"].between(0, 5).sum())
    _metric_row([
        ("Total Assets", total),
        ("Within Lifecycle", within),
        ("Expired Assets", expired),
        ("Replacement Rate", f"{expired / total * 100:.1f}%" if total else "0.0%"),
    ])

    st.subheader("Audit Summary")
    _metric_row([
        ("Assets Requiring Review", int(df["ITAM Review Required"].sum())),
        ("High Severity Assets", int(df["ITAM Highest Severity"].eq("High").sum())),
        ("Replacement Candidates", expired),
    ])

    st.subheader("Asset Portfolio")
    portfolio_left, portfolio_right = st.columns(2)
    with portfolio_left:
        _bar_chart(df, "model", "Top Models", "overview-models")
    with portfolio_right:
        _bar_chart(df, "asset_type", "Asset Type", "overview-types")

    st.subheader("Distribution")
    distribution_left, distribution_right = st.columns(2)
    with distribution_left:
        _bar_chart(df, "department", "Top Departments", "overview-departments")
    with distribution_right:
        _bar_chart(df, "location", "Top Locations", "overview-locations")


def _explorer_filters(df):
    selections = {}
    filter_columns = [
        ("asset_type", "Asset Type"), ("model", "Model / Product"),
        ("state", "Source State"), ("department", "Department"),
        ("site", "Site"), ("location", "Location"),
        ("purchase_year", "Year Of Purchase"), ("Warranty Status", "Warranty Status"),
        ("ITAM Lifecycle Status", DISPLAY_LABELS["ITAM Lifecycle Status"]), ("programme", "Programme"),
        ("workstation_status", "Workstation Status"), ("place", "Place"),
        ("job_title", "Job Title"),
    ]
    with st.expander("Filters", expanded=False):
        columns = st.columns(3)
        for index, (column, label) in enumerate(filter_columns):
            if column not in df.columns:
                continue
            with columns[index % 3]:
                selections[column] = st.multiselect(label, _values(df, column), key=f"explorer-{column}")
    return _filtered_by_selections(df, selections)


def render_asset_explorer(df, asset_type):
    st.title("Asset Explorer")
    st.caption("Find, filter, inspect and export individual asset records.")
    query = st.text_input("Search assets", placeholder="Asset Tag, Serial Number, Model, User, Site, IMEI, SIM Number...", key="asset-search")
    filtered = _explorer_filters(df)
    if query:
        filtered = filtered[filtered.astype(str).apply(lambda row: row.str.contains(query, case=False, regex=False, na=False).any(), axis=1)]
    st.write(f"{len(filtered):,} assets found")

    default_columns = [
        "asset_type", "asset_tag", "serial_number", "model", "state", "user",
        "department", "site", "location", "purchase_year", "warranty_expiry",
        "ITAM Lifecycle Status", "ITAM Review Required", "ITAM Highest Severity",
        "ITAM Audit Findings", "imei",
    ]
    st.dataframe(_display_frame(filtered, default_columns), use_container_width=True, height=520, hide_index=True)
    with st.expander("Inspect all available fields", expanded=False):
        st.dataframe(_display_frame(filtered), use_container_width=True, height=520, hide_index=True)
    with st.expander("Export Data", expanded=False):
        render_export_panel(
            filtered, "Current Filtered Results",
            f"assetlens_assets_{pd.Timestamp.now():%Y-%m-%d}.xlsx", "export-explorer",
        )


def render_lifecycle_warranty(df, asset_type):
    st.title("Lifecycle & Warranty")
    st.caption("Which assets are aging, expired, or approaching warranty expiry?")
    st.subheader("Lifecycle Summary")
    _lifecycle_metrics(df)
    average_age = df["Asset Age"].dropna().mean()
    st.metric("Average Asset Age", f"{average_age:.1f} years" if pd.notna(average_age) else "Unknown")

    lifecycle_left, lifecycle_right = st.columns(2)
    with lifecycle_left:
        _bar_chart(df, "purchase_year", "Purchase Year Analysis", "lifecycle-purchase-year")
    with lifecycle_right:
        _bar_chart(df, "Asset Age", "Asset Age Distribution", "lifecycle-age")

    if "Warranty Status" not in df.columns:
        st.info("Warranty data is not available in this export.")
        return
    st.subheader("Warranty Summary")
    _metric_row([
        ("Active Warranty", int(df["Warranty Status"].eq("Active").sum())),
        ("Expiring Soon", int(df["Warranty Status"].eq("Expiring Soon").sum())),
        ("Expired Warranty", int(df["Warranty Status"].eq("Expired").sum())),
        ("Unknown", int(df["Warranty Status"].eq("Unknown").sum())),
    ])
    for status, title in [("Expiring Soon", "Expiring Soon"), ("Expired", "Expired Warranty"), ("Expired", "Expired Lifecycle")]:
        subset = df[df["Warranty Status"].eq(status)] if title != "Expired Lifecycle" else df[df["ITAM Lifecycle Status"].eq("Expired")]
        if not subset.empty:
            with st.expander(f"{title} ({len(subset):,})", expanded=title == "Expiring Soon"):
                st.dataframe(_display_frame(subset, ["asset_tag", "serial_number", "model", "user", "site", "purchase_year", "Asset Age", "ITAM Lifecycle Status", "Warranty Status"]), use_container_width=True, hide_index=True)


def render_data_audit(df, asset_type):
    st.title("Data Audit")
    st.caption("Which source-system records need cross-checking?")
    findings = df[df["ITAM Finding Count"].gt(0)]
    _metric_row([
        ("Assets Requiring Review", int(df["ITAM Review Required"].sum())),
        ("High Severity Assets", int(df["ITAM Highest Severity"].eq("High").sum())),
        ("Assets With Findings", len(findings)),
        ("Clean Assets", int(df["ITAM Finding Count"].eq(0).sum())),
    ])
    audit_df = df.copy()
    with st.expander("Audit filters", expanded=False):
        columns = st.columns(3)
        severity = columns[0].multiselect(DISPLAY_LABELS["ITAM Highest Severity"], _values(df, "ITAM Highest Severity"), key="audit-severity")
        review = columns[1].selectbox("Review Required", ["All", "Yes", "No"], key="audit-review")
        asset_types = columns[2].multiselect("Asset Type", _values(df, "asset_type"), key="audit-type")
        site = st.multiselect("Site", _values(df, "site"), key="audit-site")
        finding_type = st.multiselect("Finding Type", ["Duplicate Asset Tag", "Duplicate Serial", "Duplicate IMEI", "Missing Core Identity", "Lifecycle / Source State mismatch", "Warranty"], key="audit-finding")
    if severity:
        audit_df = audit_df[audit_df["ITAM Highest Severity"].astype(str).isin(severity)]
    if review != "All":
        audit_df = audit_df[audit_df["ITAM Review Required"].eq(review == "Yes")]
    audit_df = _filtered_by_selections(audit_df, {"asset_type": asset_types, "site": site})
    if finding_type:
        terms = {"Duplicate Asset Tag": "Duplicate Asset Tag", "Duplicate Serial": "Duplicate Serial", "Duplicate IMEI": "Duplicate IMEI", "Missing Core Identity": "Missing core identity", "Lifecycle / Source State mismatch": "lifecycle is Expired", "Warranty": "Warranty"}
        audit_df = audit_df[audit_df["ITAM Audit Findings"].apply(lambda value: any(terms[item].casefold() in str(value).casefold() for item in finding_type))]

    audit_columns = ["asset_type", "asset_tag", "serial_number", "model", "state", "ITAM Lifecycle Status", "ITAM Highest Severity", "ITAM Review Required", "ITAM Finding Count", "ITAM Audit Findings", "user", "site", "purchase_year", "Asset Age", "Warranty Status"]
    st.dataframe(_display_frame(audit_df, audit_columns), use_container_width=True, height=520, hide_index=True)
    with st.expander("Export Data", expanded=False):
        render_export_panel(
            audit_df, "Current Audit Results",
            f"assetlens_audit_findings_{pd.Timestamp.now():%Y-%m-%d}.xlsx", "export-audit",
        )


def render_replacement_planning(df, asset_type):
    st.title("Replacement Planning")
    st.caption("Use age-based expired assets to support future replacement programme planning.")
    candidates = df[df["ITAM Replacement Candidate"].eq(True)].copy()
    total = len(df)
    average_age = candidates["Asset Age"].dropna().mean()
    oldest_age = candidates["Asset Age"].dropna().max()
    _metric_row([
        ("Replacement Candidates", len(candidates)),
        ("Percentage of Population", f"{len(candidates) / total * 100:.1f}%" if total else "0.0%"),
        ("Average Candidate Age", f"{average_age:.1f} years" if pd.notna(average_age) else "Unknown"),
        ("Oldest Candidate Age", f"{int(oldest_age)} years" if pd.notna(oldest_age) else "Unknown"),
    ])
    with st.expander("Planning filters", expanded=False):
        selections = {}
        columns = st.columns(3)
        for index, (column, label) in enumerate([( "model", "Model / Product"), ("site", "Site"), ("state", "Source State"), ("purchase_year", "Year Of Purchase"), ("programme", "Programme"), ("workstation_status", "Workstation Status")]):
            if column in candidates.columns:
                with columns[index % 3]:
                    selections[column] = st.multiselect(label, _values(candidates, column), key=f"replacement-{column}")
    candidates = _filtered_by_selections(candidates, selections if "selections" in locals() else {})
    chart_left, chart_right = st.columns(2)
    with chart_left:
        _bar_chart(candidates, "model", "Candidate Models", "replacement-models")
    with chart_right:
        _bar_chart(candidates, "site", "Candidate Sites", "replacement-sites")
    st.dataframe(_display_frame(candidates, ["asset_type", "asset_tag", "serial_number", "model", "site", "department", "purchase_year", "Asset Age", "state", "workstation_status", "programme", "ITAM Lifecycle Status"]), use_container_width=True, height=520, hide_index=True)
    with st.expander("Export Data", expanded=False):
        render_export_panel(
            candidates, "Current Replacement Candidates",
            f"assetlens_replacement_candidates_{pd.Timestamp.now():%Y-%m-%d}.xlsx", "export-replacement",
        )


def render_navigation(df, asset_type):
    page = st.sidebar.selectbox("Navigation", ["Overview", "Asset Explorer", "Lifecycle & Warranty", "Data Audit", "Replacement Planning"], key="navigation")
    renderers = {
        "Overview": render_overview,
        "Asset Explorer": render_asset_explorer,
        "Lifecycle & Warranty": render_lifecycle_warranty,
        "Data Audit": render_data_audit,
        "Replacement Planning": render_replacement_planning,
    }
    renderers[page](df, asset_type)
