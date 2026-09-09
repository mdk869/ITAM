import pandas as pd
import plotly.express as px
import streamlit as st
from html import escape

from itam.data import apply_literal_search
from itam.export import export_to_excel


def _active_theme_type():
    """Return 'dark' or 'light' for the user's active Streamlit theme (falls back to 'light')."""
    try:
        return st.context.theme.type or "light"
    except Exception:
        return "light"


def _themed_chart(figure):
    """Apply a transparent, theme-aware background/font so charts stay readable in dark mode."""
    is_dark = _active_theme_type() == "dark"
    figure.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color="#E7ECF2" if is_dark else "#172B4D",
    )
    grid_color = "rgba(231,236,242,0.15)" if is_dark else "rgba(15,39,68,0.08)"
    figure.update_xaxes(gridcolor=grid_color, zerolinecolor=grid_color)
    figure.update_yaxes(gridcolor=grid_color, zerolinecolor=grid_color)
    return figure


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
    "ITAM Planning Priority": "Planning Priority",
    "ITAM Planning Rank": "Planning Rank",
    "ITAM Duplicate Asset Tag": "Duplicate Asset Tag",
    "ITAM Duplicate Serial": "Duplicate Serial",
    "ITAM Duplicate IMEI": "Duplicate IMEI",
    "ITAM Missing Core Identity": "Missing Core Identity",
}

EXPORT_LABELS = {
    "employee_id": "Staff ID", "user": "Name", "job_title": "Designation",
    "email": "Email", "department": "Department", "place": "Place",
    "location": "Location", "state": "Asset State", "workstation_status": "Assignment Type",
    "asset_type": "Asset Type", "source_asset_subtype": "Asset Subtype", "model": "Model",
    "asset_tag": "Asset Tag", "serial_number": "Serial Number", "imei": "IMEI",
    "sim_number": "SIM Number", "site": "Site", "programme": "Programme",
    "purchase_year": "Year Of Purchase", "warranty_expiry": "Warranty Expiry",
    "Asset Age": "Asset Age", "ITAM Lifecycle Status": "Lifecycle",
    "Warranty Status": "Warranty Status", "ITAM Review Required": "Review Required",
    "ITAM Highest Severity": "Severity", "ITAM Audit Findings": "Audit Notes",
    "ITAM Finding Count": "Findings",
    "ITAM Duplicate Asset Tag": "Duplicate Asset Tag", "ITAM Duplicate Serial": "Duplicate Serial",
    "ITAM Duplicate IMEI": "Duplicate IMEI", "ITAM Missing Core Identity": "Missing Core Identity",
    "ITAM State Review Required": "State Review Required", "ITAM Replacement Candidate": "Replacement Candidate",
    "ITAM Planning Priority": "Planning Priority",
}

WORKSTATION_BASELINE = [
    "employee_id", "user", "job_title", "email", "department", "place", "location",
    "state", "workstation_status", "asset_type", "source_asset_subtype", "model",
    "asset_tag", "serial_number",
]
MOBILE_BASELINE = [
    "employee_id", "user", "job_title", "email", "department", "place", "location",
    "state", "asset_type", "model", "asset_tag", "serial_number", "imei", "sim_number",
]

# Raw source columns are intentionally absent from this export presentation contract.
EXPORT_FIELD_REGISTRY = [
    *[{"label": EXPORT_LABELS[column], "column": column, "datasets": {"Workstation"}, "mandatory": True} for column in WORKSTATION_BASELINE],
    *[{"label": EXPORT_LABELS[column], "column": column, "datasets": {"Smartphone", "Tablet"}, "mandatory": True} for column in MOBILE_BASELINE],
    *[{"label": EXPORT_LABELS[column], "column": column, "datasets": {"Workstation", "Smartphone", "Tablet"}} for column in ["site", "programme", "purchase_year", "warranty_expiry", "Asset Age", "ITAM Lifecycle Status", "Warranty Status", "ITAM Review Required", "ITAM Highest Severity", "ITAM Finding Count", "ITAM Audit Findings", "ITAM Duplicate Asset Tag", "ITAM Duplicate Serial", "ITAM Missing Core Identity", "ITAM State Review Required", "ITAM Replacement Candidate", "ITAM Planning Priority"]],
    {"label": "Audit Notes", "column": "ITAM Audit Findings", "datasets": {"Workstation", "Smartphone", "Tablet"}},
    {"label": "Duplicate IMEI", "column": "ITAM Duplicate IMEI", "datasets": {"Smartphone", "Tablet"}},
]

_merged_export_registry = {}
for _field in EXPORT_FIELD_REGISTRY:
    _existing = _merged_export_registry.get(_field["column"])
    if _existing is None:
        _merged_export_registry[_field["column"]] = _field.copy()
    else:
        _existing["datasets"] = _existing["datasets"] | _field["datasets"]
        _existing["mandatory"] = _existing.get("mandatory", False) or _field.get("mandatory", False)
EXPORT_FIELD_REGISTRY = list(_merged_export_registry.values())

PRESET_ADDITIONAL_FIELDS = {
    "Lifecycle & Warranty": ["purchase_year", "Asset Age", "ITAM Lifecycle Status", "warranty_expiry", "Warranty Status"],
    "Audit Findings": [
        "ITAM Review Required", "ITAM Highest Severity", "ITAM Audit Findings",
        "ITAM Duplicate Asset Tag", "ITAM Duplicate Serial", "ITAM Duplicate IMEI",
        "ITAM Missing Core Identity", "ITAM State Review Required",
    ],
    "Replacement Planning": [
        "purchase_year", "Asset Age", "ITAM Lifecycle Status",
        "ITAM Replacement Candidate", "ITAM Planning Priority",
    ],
}

AUDIT_CURATED_COLUMNS = [
    "ITAM Highest Severity", "ITAM Review Required", "asset_type", "asset_tag",
    "serial_number", "model", "state", "ITAM Lifecycle Status", "Finding Category",
    "ITAM Audit Findings", "site", "department", "imei",
]

EXPORT_PRESETS = {
    "Standard Asset View": [
        "asset_type", "asset_tag", "serial_number", "model", "state", "user",
        "department", "site", "location", "purchase_year", "warranty_expiry",
        "ITAM Lifecycle Status",
    ],
    "Audit Findings": [
        *AUDIT_CURATED_COLUMNS,
    ],
    "Replacement Planning": None,  # populated below from CANDIDATE_DETAIL_COLUMNS
    "Full Dataset": None,
}

SEARCH_COLUMNS = [
    "asset_tag", "serial_number", "model", "user", "employee_id", "email",
    "location", "site", "imei", "sim_number", "programme", "place",
]

# Default Candidate Detail table columns; also drives the Replacement Planning export preset.
CANDIDATE_DETAIL_COLUMNS = [
    "ITAM Planning Priority", "asset_type", "asset_tag", "serial_number", "model",
    "state", "workstation_status", "user", "department", "site", "location",
    "purchase_year", "Asset Age", "Warranty Status", "programme",
    "ITAM Lifecycle Status", "ITAM Review Required", "imei",
]
CANDIDATE_DETAIL_OPTIONAL_COLUMNS = {"workstation_status", "imei"}
EXPORT_PRESETS["Replacement Planning"] = CANDIDATE_DETAIL_COLUMNS

PLANNING_SEARCH_COLUMNS = [
    "asset_tag", "serial_number", "model", "user", "employee_id", "site",
    "location", "programme", "workstation_status", "imei",
]

PLANNING_FILTER_RESET_COLUMNS = [
    "model", "site", "department", "state", "workstation_status", "programme", "purchase_year",
]

AUDIT_SEARCH_COLUMNS = [
    "asset_tag", "serial_number", "model", "user", "employee_id", "site",
    "location", "department", "programme", "imei", "ITAM Audit Findings",
]
AUDIT_FILTER_RESET_COLUMNS = ["ITAM Highest Severity", "asset_type", "state", "site"]
AUDIT_SEVERITY_ORDER = ["Critical", "High", "Medium", "Low", "Info", "None"]
AUDIT_CATEGORY_ORDER = [
    "Duplicate Identity", "Missing Core Identity", "Lifecycle / State Review",
    "Warranty", "Other", "Clean",
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
    for column, metric in zip(columns, metrics):
        label, value = metric[0], metric[1]
        with column:
            if len(metric) > 2:
                accent = metric[2]
            else:
                normalized_label = label.casefold()
                accent = (
                    "danger" if any(term in normalized_label for term in ["expired", "high severity"])
                    else "warning" if any(term in normalized_label for term in ["review", "aging"])
                    else "success" if any(term in normalized_label for term in ["within", "active", "new"])
                    else "info"
                )
            st.markdown(
                f'<div class="al-kpi al-kpi-{accent}">'
                f'<div class="al-kpi-label">{escape(str(label))}</div>'
                f'<div class="al-kpi-value">{escape(str(value))}</div>'
                "</div>",
                unsafe_allow_html=True,
            )


def _filtered_by_selections(df, selections):
    filtered = df.copy()
    for column, selected in selections.items():
        if selected and column in filtered.columns:
            filtered = filtered[filtered[column].astype(str).isin(selected)]
    return filtered


def _meaningful_optional_columns(df, columns, optional_columns):
    """Drop optional columns that have no data at all, so e.g. IMEI never appears empty for Workstation."""
    return [
        column for column in columns
        if column not in optional_columns or (column in df.columns and df[column].notna().any())
    ]


def classify_finding_category(row):
    """Classify an audit row for display without changing locked audit fields."""
    if any(bool(row.get(column, False)) for column in [
        "ITAM Duplicate Asset Tag", "ITAM Duplicate Serial", "ITAM Duplicate IMEI",
    ]):
        return "Duplicate Identity"
    if bool(row.get("ITAM Missing Core Identity", False)):
        return "Missing Core Identity"
    if bool(row.get("ITAM State Review Required", False)):
        return "Lifecycle / State Review"
    if str(row.get("Warranty Status", "")).strip() in {"Expired", "Expiring Soon"}:
        return "Warranty"
    finding_count = row.get("ITAM Finding Count", 0)
    if pd.isna(finding_count):
        finding_count = 0
    if int(finding_count) == 0:
        return "Clean"
    return "Other"


def add_finding_categories(df):
    """Return a copy with a presentation-only Finding Category column."""
    result = df.copy()
    result["Finding Category"] = result.apply(classify_finding_category, axis=1)
    return result


def clean_asset_count(df):
    """Count assets with no findings using the authoritative finding count."""
    return int(df.get("ITAM Finding Count", pd.Series(dtype="int64")).eq(0).sum())


def audit_table_columns(df):
    """Choose readable audit columns, showing IMEI only when it contains data."""
    return _meaningful_optional_columns(df, AUDIT_CURATED_COLUMNS, {"imei"})


def make_display_dataframe(df, columns=None, label_map=None):
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
        label = (label_map or DISPLAY_LABELS).get(column, str(column))
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


def _export_asset_type(df):
    values = df.get("asset_type", pd.Series(dtype="object")).dropna().astype(str)
    return values.iloc[0] if not values.empty else "Workstation"


def _export_registry_fields(df, mandatory_only=False):
    asset_type = _export_asset_type(df)
    return [
        field for field in EXPORT_FIELD_REGISTRY
        if asset_type in field["datasets"] and (not mandatory_only or field.get("mandatory", False))
        and field["column"] in df.columns
    ]


def _export_columns_for_fields(df, fields, include_empty=False):
    columns = []
    for field in fields:
        column = field["column"]
        if column not in columns and (include_empty or df[column].notna().any()):
            columns.append(column)
    return columns


def resolve_export_columns(df, preset, custom_columns=None):
    """Resolve clean, dataset-aware export fields in operational order."""
    baseline = _export_columns_for_fields(df, _export_registry_fields(df, mandatory_only=True), include_empty=True)
    if preset in {"Standard Asset View", "Standard Asset Export"}:
        return baseline
    if preset == "Custom":
        allowed = {field["column"] for field in _export_registry_fields(df)}
        requested = [column for column in (custom_columns or []) if column in allowed]
    elif preset == "Full Dataset":
        return list(df.columns)
    else:
        requested = PRESET_ADDITIONAL_FIELDS.get(preset, [])
    additional = [column for column in requested if column in df.columns and df[column].notna().any()]
    return baseline + [column for column in additional if column not in baseline]


def prepare_export_dataframe(df, columns):
    """Prepare selected export columns with clean labels and no source aliases."""
    if not columns:
        return None
    return make_display_dataframe(df, columns, label_map=EXPORT_LABELS)


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


def render_export_panel(df, scope_label, filename, key, default_preset="Standard Asset View"):
    """Render a compact, page-local custom Excel export workflow."""
    st.subheader("Export Data")
    st.caption(scope_label)
    preset_options = ["Standard Asset Export", "Lifecycle & Warranty", "Audit Findings", "Replacement Planning", "Custom"]
    if default_preset == "Standard Asset View":
        default_preset = "Standard Asset Export"
    preset = st.selectbox(
        "Column Preset", preset_options, index=preset_options.index(default_preset),
        key=f"{key}-preset",
    )
    registry_fields = _export_registry_fields(df)
    display_to_internal = {field["label"]: field["column"] for field in registry_fields}
    available_display = list(display_to_internal)
    default_internal = resolve_export_columns(df, preset)
    default_display = [field["label"] for field in registry_fields if field["column"] in default_internal]
    columns_key = f"{key}-columns-select-{preset}"
    selected_display = st.multiselect(
        "Additional fields" if preset == "Custom" else "Columns", available_display,
        default=default_display if preset != "Custom" else [
            field["label"] for field in registry_fields
            if field["column"] in default_internal and not field.get("mandatory", False)
        ],
        key=columns_key,
    )
    select_left, select_right = st.columns(2)
    with select_left:
        if st.button("Select All", key=f"{key}-all", width="stretch"):
            st.session_state[columns_key] = available_display
            st.rerun()
    with select_right:
        if st.button("Clear", key=f"{key}-clear", width="stretch"):
            st.session_state[columns_key] = []
            st.rerun()
    selected_internal = [display_to_internal[label] for label in selected_display]
    if preset == "Custom":
        selected_internal = resolve_export_columns(df, "Custom", selected_internal)
    else:
        baseline = resolve_export_columns(df, "Standard Asset Export")
        selected_internal = baseline + [column for column in selected_internal if column not in baseline]
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
        width="stretch",
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
    st.plotly_chart(_themed_chart(figure), width="stretch", key=key)


def _lifecycle_metrics(df):
    labels = [
        ("NEW (0-1 YR)", "New"),
        ("ACTIVE (2-3 YR)", "Active"),
        ("AGING (4-5 YR)", "Aging"),
        ("EXPIRED (>5 YR)", "Expired"),
    ]
    _metric_row([(label, int(df["ITAM Lifecycle Status"].eq(value).sum())) for label, value in labels])


# ============================================================================
# OVERVIEW FILTERS (Slice 2C-1)
# ============================================================================

def _overview_filter_columns(asset_type):
    """Return list of (column, display_label) tuples for this asset_type."""
    # All datasets have these
    base_filters = [
        ("place", "Place"),
        ("site", "Site"),
        ("location", "Location"),
        ("department", "Department"),
    ]
    # Workstation-only filters
    if asset_type == "Workstation":
        return [
            ("source_asset_subtype", "Asset Subtype"),
            ("workstation_status", "Assignment Type"),
        ] + base_filters
    # Smartphone and Tablet: no Workstation-only filters
    return base_filters


def _init_overview_filters(asset_type):
    """Ensure all Overview filter session state keys exist."""
    for column, _ in _overview_filter_columns(asset_type):
        key = f"overview-{column}"
        if key not in st.session_state:
            st.session_state[key] = []


def _apply_overview_filters(df, asset_type):
    """Apply filters from session state and return filtered copy."""
    filtered = df.copy()
    for column, _ in _overview_filter_columns(asset_type):
        key = f"overview-{column}"
        selected = st.session_state.get(key, [])
        if selected and column in filtered.columns:
            filtered = filtered[filtered[column].astype(str).isin(selected)]
    return filtered


def _get_selected_filters(asset_type):
    """Get dict of {column: [selected_values]} for display purposes."""
    result = {}
    for column, _ in _overview_filter_columns(asset_type):
        key = f"overview-{column}"
        selected = st.session_state.get(key, [])
        if selected:
            result[column] = selected
    return result


def _format_context_indicator(total_unfiltered, filtered_df, selected_filters, asset_type):
    """Return concise context string e.g. 'Viewing: Pool • Kuala Selangor Region • 200 of 6,423 assets'."""
    filtered_count = len(filtered_df)
    if filtered_count == total_unfiltered and not selected_filters:
        return f"Viewing all {total_unfiltered:,} assets"
    
    # Build filter summary with user-friendly labels
    filter_parts = []
    label_map = {
        "source_asset_subtype": "Asset Subtype",
        "workstation_status": "Assignment Type",
        "place": "Place",
        "site": "Site",
        "location": "Location",
        "department": "Department",
    }
    for column, values in selected_filters.items():
        if values:
            label = label_map.get(column, column)
            # Show first value or count of multiple
            if len(values) == 1:
                filter_parts.append(str(values[0]))
            else:
                filter_parts.append(f"{label} ({len(values)})")
    
    filter_summary = " • ".join(filter_parts) if filter_parts else ""
    if filter_summary:
        return f"Viewing: {filter_summary} • {filtered_count:,} of {total_unfiltered:,} assets"
    return f"Viewing {filtered_count:,} of {total_unfiltered:,} assets"


def _clear_overview_filters():
    """Reset all Overview filter widget state."""
    for column, _ in _overview_filter_columns(st.session_state.get("_overview_asset_type", "Workstation")):
        st.session_state[f"overview-{column}"] = []


def _overview_missing_count(df, column):
    """Return the number of null or blank values in an Overview column."""
    if column not in df.columns:
        return 0
    values = df[column].astype("string")
    return int(values.isna().sum() + values.str.strip().eq("").sum())


def _display_group_value(value):
    """Map missing organisational values to a presentation-only group label."""
    if pd.isna(value):
        return "Not Assigned"
    text = str(value).strip()
    return text if text else "Not Assigned"


def _overview_distribution(df, column, top_n=None):
    """Return a blank-free categorical summary with counts and percentages."""
    summary_columns = ["Category", "Units", "% of Assets"]
    if column not in df.columns:
        return pd.DataFrame(columns=summary_columns)
    values = df[column].dropna().astype(str).str.strip()
    values = values[values.ne("")]
    if values.empty:
        return pd.DataFrame(columns=summary_columns)
    counts = values.value_counts()
    if top_n is not None:
        counts = counts.head(top_n)
    return pd.DataFrame({
        "Category": counts.index,
        "Units": counts.values,
        "% of Assets": (counts.values / len(df) * 100).round(1),
    })


def _overview_composition_summary(df, asset_type):
    """Summarize source subtype for Workstations and model for mobile assets."""
    column = "source_asset_subtype" if asset_type == "Workstation" else "model"
    return _overview_distribution(df, column)


def _overview_model_summary(df, top_n=10):
    """Return the leading models in the current Overview context."""
    return _overview_distribution(df, "model", top_n=top_n)


def _overview_assignment_summary(df):
    """Return exact source assignment categories without recoding values."""
    return _overview_distribution(df, "workstation_status")


def _overview_group_summary(df, group_column, asset_type):
    """Aggregate operational metrics by one level of the Place hierarchy."""
    columns = [group_column.title(), "Total", "Review", "Replacement"]
    if group_column not in df.columns:
        return pd.DataFrame(columns=columns)
    data = df.copy()
    if data.empty:
        return pd.DataFrame(columns=columns)
    data[group_column] = data[group_column].map(_display_group_value)
    data["_overview_review"] = data["ITAM Review Required"].eq(True)
    data["_overview_replacement"] = data["ITAM Replacement Candidate"].eq(True)
    grouped = data.groupby(group_column, dropna=False).agg(
        Total=(group_column, "size"),
        Review=("_overview_review", "sum"),
        Replacement=("_overview_replacement", "sum"),
    )
    if asset_type == "Workstation" and "workstation_status" in data.columns:
        statuses = data["workstation_status"].astype("string").fillna("").str.strip()
        for status in ["Personal", "Pool", "Counter"]:
            data[f"_overview_{status.lower()}"] = statuses.eq(status)
            grouped[status] = data.groupby(group_column)[f"_overview_{status.lower()}"].sum()
        grouped = grouped[["Total", "Personal", "Pool", "Counter", "Review", "Replacement"]]
    return grouped.reset_index().rename(columns={group_column: group_column.title()}).sort_values(
        "Total", ascending=False, kind="stable"
    ).reset_index(drop=True)


def _overview_place_summary(df, asset_type):
    """Return the compact Place Intelligence summary table."""
    return _overview_group_summary(df, "place", asset_type)


def _overview_place_subset(df, place):
    """Return a local Place drill-down subset without changing global filters."""
    if "place" not in df.columns or not place:
        return df.iloc[0:0].copy()
    display_values = df["place"].map(_display_group_value)
    return df.loc[display_values.eq(str(place))].copy()


def _overview_local_place_selection(selected_place, place_options):
    """Keep a local Place selection only while it remains in the filtered context."""
    return selected_place if selected_place in place_options else "Select a Place"


def _overview_ordered_status_summary(df, column, order, missing_label="Unknown"):
    """Summarize a derived status in semantic order without hiding missing values."""
    summary_columns = ["Category", "Assets", "% of Assets"]
    if column not in df.columns:
        return pd.DataFrame(columns=summary_columns)
    values = df[column].map(lambda value: missing_label if pd.isna(value) or not str(value).strip() else str(value).strip())
    counts = values.value_counts()
    categories = [category for category in order if category in counts]
    categories.extend(category for category in counts.index if category not in categories)
    return pd.DataFrame({
        "Category": categories,
        "Assets": [int(counts[category]) for category in categories],
        "% of Assets": [(counts[category] / len(df) * 100).round(1) for category in categories],
    })


def _overview_lifecycle_summary(df):
    """Return the authoritative lifecycle distribution in business order."""
    return _overview_ordered_status_summary(
        df, "ITAM Lifecycle Status", ["New", "Active", "Aging", "Expired", "Unknown"]
    )


def _overview_warranty_summary(df):
    """Return the authoritative warranty distribution in business order."""
    return _overview_ordered_status_summary(
        df, "Warranty Status", ["Active", "Expiring Soon", "Expired", "Unknown"]
    )


def _overview_severity_summary(df):
    """Return present audit severities in descending severity order."""
    return _overview_ordered_status_summary(
        df, "ITAM Highest Severity", AUDIT_SEVERITY_ORDER, missing_label="None"
    )


def _overview_finding_category_summary(df, top_n=5):
    """Return the leading affected-asset finding categories without raw audit text."""
    categorized = add_finding_categories(df)
    categories = categorized.loc[categorized["Finding Category"].ne("Clean"), "Finding Category"]
    counts = categories.value_counts().head(top_n)
    return pd.DataFrame({
        "Finding Category": counts.index,
        "Assets": counts.values,
        "% of Assets": (counts.values / len(df) * 100).round(1),
    })


def _overview_priority_summary(df):
    """Return the full derived planning-priority profile in locked order."""
    return _overview_ordered_status_summary(
        df,
        "ITAM Planning Priority",
        ["Priority Review", "Standard Planning", "Low Operational Priority", "Not Candidate"],
        missing_label="Not Candidate",
    )


def _overview_candidate_priority_counts(df):
    """Return candidate-only priority counts from the existing derived priority field."""
    summary = _overview_priority_summary(df)
    return summary.loc[summary["Category"].ne("Not Candidate")].set_index("Category")["Assets"]


def _overview_organisational_distribution(df, column):
    """Return a full organisational distribution using presentation-only Not Assigned values."""
    summary_columns = ["Category", "Units", "% of Assets"]
    if column not in df.columns:
        return pd.DataFrame(columns=summary_columns)
    values = df[column].map(_display_group_value)
    counts = values.value_counts()
    return pd.DataFrame({
        "Category": counts.index,
        "Units": counts.values,
        "% of Assets": (counts.values / len(df) * 100).round(1),
    })


def _render_overview_distribution(summary, title, key):
    """Render a compact themed distribution chart from an Overview summary."""
    if summary.empty:
        st.info(f"No {title.casefold()} values are available in the current context.")
        return
    chart_data = summary.sort_values("Units")
    figure = px.bar(
        chart_data, x="Units", y="Category", orientation="h", text="Units",
        title=title, custom_data=["% of Assets"],
        labels={"Category": "", "Units": "Assets"},
    )
    figure.update_traces(
        textposition="outside", cliponaxis=False,
        hovertemplate="%{y}<br>Units: %{x:,}<br>Share: %{customdata[0]}%<extra></extra>",
    )
    figure.update_layout(height=300, margin=dict(t=45, b=20, l=10, r=35), showlegend=False)
    st.plotly_chart(_themed_chart(figure), width="stretch", key=key)


def render_overview(df, asset_type):
    st.title("Overview")
    st.caption(f"Complete view of your IT asset inventory, lifecycle, audit status and replacement outlook.")
    
    # Track asset type for clear filters callback
    st.session_state["_overview_asset_type"] = asset_type
    
    # ========================================================================
    # GLOBAL CONTEXT FILTERS
    # ========================================================================
    total_unfiltered = len(df)
    filter_columns = _overview_filter_columns(asset_type)
    
    # Build filter context: apply upstream filters to determine downstream options
    # and sanitize stale selections BEFORE widget instantiation
    filter_context = df.copy()
    
    # Render section header with Clear Filters as secondary action
    header_cols = st.columns([1, 0.15], gap="large")
    with header_cols[0]:
        st.markdown("### Analysis Context")
    with header_cols[1]:
        st.button("Clear Filters", key="overview-clear-filters", on_click=_clear_overview_filters)
    
    # Render filters in balanced rows based on asset type
    if asset_type == "Workstation":
        # Workstation: 6 filters in 2 rows of 3
        # Row 1: Asset Subtype | Assignment Type | Place
        row1_cols = st.columns(3)
        row1_filters = filter_columns[:3]
        for col_idx, (column, label) in enumerate(row1_filters):
            with row1_cols[col_idx]:
                available_values = _values(filter_context, column) if column in filter_context.columns else []
                available_values_set = set(available_values)
                current_selection = st.session_state.get(f"overview-{column}", [])
                sanitized_selection = [v for v in current_selection if v in available_values_set]
                if sanitized_selection != current_selection:
                    st.session_state[f"overview-{column}"] = sanitized_selection
                st.multiselect(
                    label, available_values,
                    default=sanitized_selection,
                    key=f"overview-{column}",
                )
                selected = st.session_state.get(f"overview-{column}", [])
                if selected and column in filter_context.columns:
                    filter_context = filter_context[filter_context[column].astype(str).isin(selected)]
        
        # Row 2: Site | Location | Department
        row2_cols = st.columns(3)
        row2_filters = filter_columns[3:]
        for col_idx, (column, label) in enumerate(row2_filters):
            with row2_cols[col_idx]:
                available_values = _values(filter_context, column) if column in filter_context.columns else []
                available_values_set = set(available_values)
                current_selection = st.session_state.get(f"overview-{column}", [])
                sanitized_selection = [v for v in current_selection if v in available_values_set]
                if sanitized_selection != current_selection:
                    st.session_state[f"overview-{column}"] = sanitized_selection
                st.multiselect(
                    label, available_values,
                    default=sanitized_selection,
                    key=f"overview-{column}",
                )
                selected = st.session_state.get(f"overview-{column}", [])
                if selected and column in filter_context.columns:
                    filter_context = filter_context[filter_context[column].astype(str).isin(selected)]
    else:
        # Smartphone/Tablet: 4 filters in 1 balanced row
        filter_cols = st.columns(4)
        for col_idx, (column, label) in enumerate(filter_columns):
            with filter_cols[col_idx]:
                available_values = _values(filter_context, column) if column in filter_context.columns else []
                available_values_set = set(available_values)
                current_selection = st.session_state.get(f"overview-{column}", [])
                sanitized_selection = [v for v in current_selection if v in available_values_set]
                if sanitized_selection != current_selection:
                    st.session_state[f"overview-{column}"] = sanitized_selection
                st.multiselect(
                    label, available_values,
                    default=sanitized_selection,
                    key=f"overview-{column}",
                )
                selected = st.session_state.get(f"overview-{column}", [])
                if selected and column in filter_context.columns:
                    filter_context = filter_context[filter_context[column].astype(str).isin(selected)]
    
    # ========================================================================
    # APPLY FILTERS TO GET OVERVIEW DATAFRAME
    # ========================================================================
    overview_df = _apply_overview_filters(df, asset_type)
    selected_filters = _get_selected_filters(asset_type)
    
    # ========================================================================
    # CONTEXT INDICATOR
    # ========================================================================
    context_text = _format_context_indicator(total_unfiltered, overview_df, selected_filters, asset_type)
    st.markdown(f"**{context_text}**")
    
    # ========================================================================
    # MANAGEMENT SNAPSHOT (5 KPIs)
    # ========================================================================
    if overview_df.empty:
        st.info("No assets match the current filter selections. Adjust or clear filters to see results.")
        return
    
    st.markdown("### Management Snapshot")
    
    total = len(overview_df)
    within = int(overview_df["Asset Age"].between(0, 5).sum())
    expired = int(overview_df["ITAM Replacement Candidate"].sum())
    review_required = int(overview_df["ITAM Review Required"].sum())
    replacement_candidates = int(overview_df["ITAM Replacement Candidate"].sum())
    
    _metric_row([
        ("Total Assets", total),
        ("Within Lifecycle", within),
        ("Expired", expired),
        ("Review Required", review_required),
        ("Replacement Candidates", replacement_candidates),
    ])

    st.markdown("### Asset Composition & Model Intelligence")
    composition_col, model_col = st.columns(2)
    composition_column = "source_asset_subtype" if asset_type == "Workstation" else "model"
    composition_label = "Asset Subtype" if asset_type == "Workstation" else "Model"
    composition_summary = _overview_composition_summary(overview_df, asset_type)
    with composition_col:
        _render_overview_distribution(composition_summary, "Asset Composition", "overview-composition")
        missing_composition = _overview_missing_count(overview_df, composition_column)
        if missing_composition:
            st.caption(f"{missing_composition:,} assets with blank {composition_label.casefold()} values are omitted.")
    with model_col:
        st.markdown("#### Leading Models")
        model_summary = _overview_model_summary(overview_df, top_n=10).rename(columns={"Category": "Model"})
        if model_summary.empty:
            st.info("No model values are available in the current context.")
        else:
            st.dataframe(model_summary, width="stretch", height=300, hide_index=True)
        missing_models = _overview_missing_count(overview_df, "model")
        if missing_models:
            st.caption(f"{missing_models:,} assets with blank model values are omitted.")

    if asset_type == "Workstation":
        st.markdown("### Assignment Intelligence")
        assignment_summary = _overview_assignment_summary(overview_df).rename(columns={"Category": "Assignment Type"})
        if assignment_summary.empty:
            st.info("Assignment Type is not available for the current context.")
        else:
            st.dataframe(assignment_summary, width="stretch", hide_index=True)
        missing_assignments = _overview_missing_count(overview_df, "workstation_status")
        if missing_assignments:
            st.caption(f"{missing_assignments:,} assets with blank Assignment Type values are omitted.")

    st.markdown("### Place Intelligence")
    place_summary = _overview_place_summary(overview_df, asset_type)
    if place_summary.empty:
        st.info("No Place values are available in the current context.")
        return
    st.dataframe(place_summary, width="stretch", hide_index=True)

    place_key = "overview-place-drilldown"
    place_options = place_summary["Place"].tolist()
    current_place = st.session_state.get(place_key)
    sanitized_place = _overview_local_place_selection(current_place, place_options)
    if current_place is not None and sanitized_place != current_place:
        st.session_state[place_key] = sanitized_place
    selected_place = st.selectbox(
        "Place detail", ["Select a Place", *place_options],
        key=place_key,
    )
    if selected_place != "Select a Place":
        selected_place_df = _overview_place_subset(overview_df, selected_place)
        place_metrics = [
            ("Total Assets", len(selected_place_df)),
            ("Review Required", int(selected_place_df["ITAM Review Required"].eq(True).sum())),
            ("Replacement Candidates", int(selected_place_df["ITAM Replacement Candidate"].eq(True).sum())),
        ]
        if asset_type == "Workstation":
            statuses = selected_place_df["workstation_status"].astype("string").fillna("").str.strip()
            place_metrics[1:1] = [
                ("Personal", int(statuses.eq("Personal").sum())),
                ("Pool", int(statuses.eq("Pool").sum())),
                ("Counter", int(statuses.eq("Counter").sum())),
            ]
        st.markdown(f"#### {selected_place}")
        _metric_row(place_metrics)

        site_summary = _overview_group_summary(selected_place_df, "site", asset_type)
        st.markdown("#### Site Breakdown")
        if site_summary.empty:
            st.info("No Site values are available for the selected Place.")
        else:
            st.dataframe(site_summary, width="stretch", hide_index=True)

        location_summary = _overview_group_summary(selected_place_df, "location", asset_type)
        st.markdown("#### Location Breakdown")
        if location_summary.empty:
            st.info("No Location values are available for the selected Place.")
        else:
            st.dataframe(location_summary, width="stretch", hide_index=True)

    st.markdown("### Lifecycle & Warranty")
    lifecycle_col, warranty_col = st.columns(2)
    with lifecycle_col:
        st.markdown("#### Lifecycle Profile")
        st.dataframe(_overview_lifecycle_summary(overview_df), width="stretch", hide_index=True)
    with warranty_col:
        st.markdown("#### Warranty Profile")
        st.dataframe(_overview_warranty_summary(overview_df), width="stretch", hide_index=True)

    st.markdown("### Audit Health")
    severity_col, findings_col = st.columns(2)
    with severity_col:
        st.markdown("#### Severity Distribution")
        _render_overview_distribution(
            _overview_severity_summary(overview_df).rename(columns={"Assets": "Units"}),
            "Audit Severity", "overview-audit-severity",
        )
    with findings_col:
        st.markdown("#### Top Finding Categories")
        finding_summary = _overview_finding_category_summary(overview_df, top_n=5)
        if finding_summary.empty:
            st.info("No audit finding categories are present in the current context.")
        else:
            st.dataframe(finding_summary, width="stretch", hide_index=True)

    st.markdown("### Replacement Outlook")
    candidate_counts = _overview_candidate_priority_counts(overview_df)
    _metric_row([
        ("Replacement Candidates", int(overview_df["ITAM Replacement Candidate"].eq(True).sum())),
        ("Priority Review", int(candidate_counts.get("Priority Review", 0))),
        ("Standard Planning", int(candidate_counts.get("Standard Planning", 0))),
        ("Low Operational Priority", int(candidate_counts.get("Low Operational Priority", 0))),
    ])
    _render_overview_distribution(
        _overview_priority_summary(overview_df).rename(columns={"Assets": "Units"}),
        "Planning Priority Profile", "overview-planning-priority",
    )

    st.markdown("### Organisational Distribution")
    distribution_labels = {
        "Place": "place", "Site": "site", "Location": "location", "Department": "department",
    }
    distribution_key = "overview-organisational-distribution"
    selected_distribution = st.selectbox("Distribution By", list(distribution_labels), key=distribution_key)
    organisational_summary = _overview_organisational_distribution(
        overview_df, distribution_labels[selected_distribution]
    )
    top_distribution = organisational_summary.head(10)
    _render_overview_distribution(
        top_distribution, f"Top 10 {selected_distribution}", "overview-organisational-chart",
    )
    if len(organisational_summary) > len(top_distribution):
        st.caption(f"Showing the top 10 of {len(organisational_summary):,} {selected_distribution.casefold()} categories.")


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
    st.dataframe(_display_frame(filtered, default_columns), width="stretch", height=520, hide_index=True)
    with st.expander("Inspect all available fields", expanded=False):
        st.dataframe(_display_frame(filtered), width="stretch", height=520, hide_index=True)
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
                st.dataframe(_display_frame(subset, ["asset_tag", "serial_number", "model", "user", "site", "purchase_year", "Asset Age", "ITAM Lifecycle Status", "Warranty Status"]), width="stretch", hide_index=True)


def render_data_audit(df, asset_type):
    st.title("Data Audit")
    st.caption("Which source-system records need cross-checking?")
    audited_df = add_finding_categories(df)
    _metric_row([
        ("Review Required", int(audited_df["ITAM Review Required"].sum())),
        ("High Severity", int(audited_df["ITAM Highest Severity"].eq("High").sum())),
        ("Medium Severity", int(audited_df["ITAM Highest Severity"].eq("Medium").sum())),
        ("Low Severity", int(audited_df["ITAM Highest Severity"].eq("Low").sum())),
        ("Info", int(audited_df["ITAM Highest Severity"].eq("Info").sum())),
        ("Clean Assets", clean_asset_count(audited_df)),
    ])

    for title, column, order, key, height in [
        ("Severity Distribution", "ITAM Highest Severity", AUDIT_SEVERITY_ORDER, "audit-severity-distribution", 260),
        ("Finding Categories", "Finding Category", AUDIT_CATEGORY_ORDER, "audit-finding-categories", 300),
    ]:
        st.subheader(title)
        counts = audited_df[column].value_counts().reindex(order, fill_value=0)
        counts = counts[counts.gt(0)].sort_values()
        if not counts.empty:
            figure = px.bar(x=counts.values, y=counts.index, orientation="h", labels={"x": "Assets", "y": title})
            figure.update_layout(height=height, margin=dict(t=15, b=20, l=10, r=10), showlegend=False)
            st.plotly_chart(_themed_chart(figure), width="stretch", key=key)

    query = st.text_input("Search findings", placeholder="Asset Tag, Serial Number, Model, User, Site, IMEI, Audit Notes...", key="audit-search")
    with st.expander("Audit filters", expanded=False):
        clear_col, _ = st.columns([1, 5])
        with clear_col:
            st.write("")
            st.button("Clear audit filters", key="audit-clear-filters", on_click=_clear_audit_filters, width="stretch")
        columns = st.columns(3)
        with columns[0]:
            severity = st.multiselect("Severity", _values(df, "ITAM Highest Severity"), key="audit-severity")
            categories = st.multiselect("Finding Category", AUDIT_CATEGORY_ORDER, key="audit-category")
        with columns[1]:
            review = st.selectbox("Review Required", ["All", "Yes", "No"], key="audit-review")
            asset_types = st.multiselect("Asset Type", _values(df, "asset_type"), key="audit-type")
        with columns[2]:
            source_states = st.multiselect("Source State", _values(df, "state"), key="audit-state")
            site = st.multiselect("Site", _values(df, "site"), key="audit-site")
    if severity:
        audited_df = audited_df[audited_df["ITAM Highest Severity"].astype(str).isin(severity)]
    if review != "All":
        audited_df = audited_df[audited_df["ITAM Review Required"].eq(review == "Yes")]
    audited_df = _filtered_by_selections(audited_df, {"asset_type": asset_types, "state": source_states, "site": site, "Finding Category": categories})
    if query:
        search_columns = [column for column in AUDIT_SEARCH_COLUMNS if column in audited_df.columns]
        audited_df = apply_literal_search(audited_df, query, columns=search_columns)

    st.caption(f"Showing {len(audited_df):,} of {len(df):,} audited assets")
    if audited_df.empty:
        st.info("No audited assets match the current filters.")
    else:
        st.dataframe(_display_frame(audited_df, audit_table_columns(audited_df)), width="stretch", height=520, hide_index=True)
    with st.expander("Export Data", expanded=False):
        render_export_panel(
            audited_df, "Current Audit Results",
            f"assetlens_audit_findings_{pd.Timestamp.now():%Y-%m-%d}.xlsx", "export-audit",
            default_preset="Audit Findings",
        )


def _planning_priority_chart(candidates):
    order = ["Priority Review", "Standard Planning", "Low Operational Priority"]
    counts = candidates["ITAM Planning Priority"].value_counts().reindex(order, fill_value=0)
    total = int(counts.sum())
    if total == 0:
        st.info("There are no Replacement Candidates to segment yet.")
        return
    percentages = (counts / total * 100).round(1)
    figure = px.bar(
        x=counts.values, y=order, orientation="h",
        text=[f"{count:,} ({pct:.1f}%)" for count, pct in zip(counts.values, percentages.values)],
        labels={"x": "Replacement Candidates", "y": ""},
    )
    figure.update_traces(textposition="outside")
    figure.update_layout(height=300, margin=dict(t=20, b=20, l=10, r=10), showlegend=False, yaxis=dict(categoryorder="array", categoryarray=order[::-1]))
    st.plotly_chart(_themed_chart(figure), width="stretch", key="replacement-planning-priority")


def _candidate_age_chart(candidates, key="replacement-candidate-age-profile", title=None):
    age_labels = ["Age 6", "Age 7", "Age 8", "Age 9", "Age 10", "Age 11+"]

    def bucket(age):
        if pd.isna(age):
            return None
        age = int(age)
        return "Age 11+" if age >= 11 else f"Age {age}"

    buckets = candidates["Asset Age"].map(bucket)
    counts = buckets.value_counts().reindex(age_labels, fill_value=0)
    if int(counts.sum()) == 0:
        st.info("There are no Replacement Candidates to profile yet.")
        return
    figure = px.bar(
        x=age_labels, y=counts.values, title=title,
        labels={"x": "Asset Age", "y": "Replacement Candidates"},
        category_orders={"x": age_labels},
    )
    figure.update_layout(height=320, margin=dict(t=35 if title else 20, b=20, l=10, r=10), showlegend=False)
    st.plotly_chart(_themed_chart(figure), width="stretch", key=key)


def _top_n_counts(series, top_n=10):
    """Blank-filtered, descending value counts for a categorical breakdown (top N)."""
    if series is None:
        return pd.Series(dtype="int64")
    cleaned = series.dropna().astype(str).str.strip()
    cleaned = cleaned[cleaned.ne("") & cleaned.str.casefold().ne("na")]
    if cleaned.empty:
        return pd.Series(dtype="int64")
    return cleaned.value_counts().head(top_n)


def _chronological_counts(series):
    """Blank-filtered value counts ordered chronologically (ascending) for numeric-like values."""
    if series is None:
        return pd.Series(dtype="int64")
    cleaned = series.dropna().astype(str).str.strip()
    cleaned = cleaned[cleaned.ne("") & cleaned.str.casefold().ne("na")]
    if cleaned.empty:
        return pd.Series(dtype="int64")
    counts = cleaned.value_counts()
    try:
        order = sorted(counts.index, key=float)
    except ValueError:
        order = sorted(counts.index)
    return counts.reindex(order)


def _truncate_label(value, max_length=32):
    """Shorten a chart label so long names cannot break the layout."""
    text = str(value)
    return text if len(text) <= max_length else text[: max_length - 1].rstrip() + "…"


def _breakdown_bar_chart(df, column, title, key, top_n=10):
    """Render a compact top-N descending breakdown; returns False when there is no meaningful data."""
    if column not in df.columns:
        return False
    counts = _top_n_counts(df[column], top_n)
    if counts.empty:
        return False
    display_counts = counts.sort_values()
    labels = [_truncate_label(label) for label in display_counts.index]
    figure = px.bar(
        x=display_counts.values, y=labels, orientation="h", title=title,
        labels={"x": "Replacement Candidates", "y": ""},
    )
    figure.update_layout(height=340, margin=dict(t=45, b=20, l=10, r=10), showlegend=False)
    st.plotly_chart(_themed_chart(figure), width="stretch", key=key)
    return True


def _chronological_bar_chart(df, column, title, key):
    """Render a candidate breakdown in chronological order; returns False when there is no meaningful data."""
    if column not in df.columns:
        return False
    counts = _chronological_counts(df[column])
    if counts.empty:
        return False
    labels = list(counts.index)
    figure = px.bar(
        x=labels, y=counts.values, title=title,
        labels={"x": title, "y": "Replacement Candidates"},
        category_orders={"x": labels},
    )
    figure.update_layout(height=320, margin=dict(t=45, b=20, l=10, r=10), showlegend=False)
    st.plotly_chart(_themed_chart(figure), width="stretch", key=key)
    return True


def _clear_replacement_planning_filters():
    """Reset Replacement Planning widget state; must run as a button on_click callback.

    Streamlit callbacks execute before the widgets are re-instantiated on the
    next run, so it is safe to write these widget-backed keys here even though
    it would raise StreamlitWidgetAlreadyInstantiatedError if done inline
    after the widgets have already rendered in the current run.
    """
    st.session_state["replacement-segment"] = "All Replacement Candidates"
    for column in PLANNING_FILTER_RESET_COLUMNS:
        st.session_state[f"replacement-{column}"] = []
    st.session_state["replacement-search"] = ""


def _clear_audit_filters():
    """Reset Data Audit widget state from a pre-render callback."""
    st.session_state["audit-severity"] = []
    st.session_state["audit-review"] = "All"
    st.session_state["audit-category"] = []
    st.session_state["audit-type"] = []
    st.session_state["audit-state"] = []
    st.session_state["audit-site"] = []
    st.session_state["audit-search"] = ""


def render_replacement_planning(df, asset_type):
    st.title("Replacement Planning")
    st.caption("Use age-based expired assets to support future replacement programme planning.")
    candidates = df[df["ITAM Replacement Candidate"].eq(True)].copy()
    total = len(df)
    candidate_count = len(candidates)
    priority_counts = candidates["ITAM Planning Priority"].value_counts()

    st.subheader("Planning Summary")
    _metric_row([
        ("Replacement Candidates", candidate_count, "navy"),
        ("Priority Review", int(priority_counts.get("Priority Review", 0)), "warning"),
        ("Standard Planning", int(priority_counts.get("Standard Planning", 0)), "teal"),
        ("Low Operational Priority", int(priority_counts.get("Low Operational Priority", 0)), "muted"),
        ("Candidate Rate", f"{candidate_count / total * 100:.1f}%" if total else "0.0%", "info"),
    ])

    st.subheader("Planning Priority")
    _planning_priority_chart(candidates)

    st.subheader("Candidate Age Profile")
    _candidate_age_chart(candidates)

    st.subheader("Planning Filters")
    segment_options = ["All Replacement Candidates", "Priority Review", "Standard Planning", "Low Operational Priority"]
    search_query = st.text_input(
        "Search candidates",
        placeholder="Asset Tag, Serial Number, Model, User, Site, IMEI...",
        key="replacement-search",
    )
    with st.expander("Planning filters", expanded=False):
        segment_col, clear_col = st.columns([4, 1])
        with segment_col:
            segment = st.selectbox("Planning Priority Segment", segment_options, key="replacement-segment")
        with clear_col:
            st.write("")
            st.button(
                "Clear planning filters", key="replacement-clear-filters",
                width="stretch", on_click=_clear_replacement_planning_filters,
            )
        selections = {}
        columns = st.columns(3)
        filter_fields = [
            ("model", "Model / Product"), ("site", "Site"), ("department", "Department"),
            ("state", "Source State"), ("workstation_status", "Workstation Status"),
            ("programme", "Programme"), ("purchase_year", "Year Of Purchase"),
        ]
        for index, (column, label) in enumerate(filter_fields):
            if column in candidates.columns:
                with columns[index % 3]:
                    selections[column] = st.multiselect(label, _values(candidates, column), key=f"replacement-{column}")

    segmented = candidates if segment == "All Replacement Candidates" else candidates[candidates["ITAM Planning Priority"].eq(segment)]
    segmented = _filtered_by_selections(segmented, selections)
    if search_query:
        search_columns = [column for column in PLANNING_SEARCH_COLUMNS if column in segmented.columns]
        segmented = apply_literal_search(segmented, search_query, columns=search_columns or None)
    if "ITAM Planning Rank" in segmented.columns:
        segmented = segmented.sort_values(["ITAM Planning Rank", "Asset Age"], ascending=[True, False])

    st.subheader("Operational Distribution")
    operational_left, operational_right = st.columns(2)
    with operational_left:
        _breakdown_bar_chart(segmented, "state", "Source State", "replacement-breakdown-source-state")
    with operational_right:
        _breakdown_bar_chart(segmented, "workstation_status", "Workstation Status", "replacement-breakdown-workstation-status")

    st.subheader("Organizational Distribution")
    org_left, org_right = st.columns(2)
    with org_left:
        _breakdown_bar_chart(segmented, "site", "Site", "replacement-breakdown-site")
    with org_right:
        _breakdown_bar_chart(segmented, "department", "Department", "replacement-breakdown-department")
    _breakdown_bar_chart(segmented, "programme", "Programme", "replacement-breakdown-programme")

    st.subheader("Asset Profile")
    asset_left, asset_right = st.columns(2)
    with asset_left:
        _breakdown_bar_chart(segmented, "model", "Model / Product", "replacement-breakdown-model")
    with asset_right:
        _chronological_bar_chart(segmented, "purchase_year", "Year Of Purchase", "replacement-breakdown-purchase-year")
    if segment != "All Replacement Candidates":
        _candidate_age_chart(segmented, key="replacement-breakdown-asset-age", title=f"Asset Age — {segment}")

    st.subheader("Candidate Detail")
    summary_caption = f"Showing {len(segmented):,} of {candidate_count:,} replacement candidates"
    if segment != "All Replacement Candidates":
        summary_caption += f" — segment: {segment}"
    st.caption(summary_caption)
    if segmented.empty:
        st.info("No replacement candidates match the current filters.")
    else:
        detail_columns = _meaningful_optional_columns(df, CANDIDATE_DETAIL_COLUMNS, CANDIDATE_DETAIL_OPTIONAL_COLUMNS)
        st.dataframe(_display_frame(segmented, detail_columns), width="stretch", height=520, hide_index=True)
    with st.expander("Export Data", expanded=False):
        render_export_panel(
            segmented, "Current Replacement Candidates",
            f"assetlens_replacement_candidates_{pd.Timestamp.now():%Y-%m-%d}.xlsx", "export-replacement",
        )


NAVIGATION_ITEMS = [
    ("⌂", "Overview"),
    ("▣", "Asset Explorer"),
    ("◷", "Lifecycle & Warranty"),
    ("✓", "Data Audit"),
    ("▤", "Replacement Planning"),
]


def render_sidebar_dataset_info(asset_type=None, filename=None, record_count=None, container=None):
    """Render the compact dataset summary used by the global sidebar shell."""
    target = container.container() if container is not None else st.sidebar
    target.markdown('<div class="al-sidebar-section-label">Dataset</div>', unsafe_allow_html=True)
    if not asset_type or record_count is None:
        target.markdown('<div class="al-dataset-empty">No dataset loaded</div>', unsafe_allow_html=True)
        return
    safe_filename = escape(str(filename or "Uploaded workbook"))
    target.markdown(
        '<div class="al-dataset-panel">'
        f'<div class="al-dataset-row"><span>Dataset</span><strong>{escape(str(asset_type))} Assets</strong></div>'
        f'<div class="al-dataset-row"><span>Source</span><strong title="{safe_filename}">{safe_filename}</strong></div>'
        f'<div class="al-dataset-row"><span>Records</span><strong>{int(record_count):,}</strong></div>'
        '</div>',
        unsafe_allow_html=True,
    )


def _page_renderers():
    return {
        "Overview": render_overview,
        "Asset Explorer": render_asset_explorer,
        "Lifecycle & Warranty": render_lifecycle_warranty,
        "Data Audit": render_data_audit,
        "Replacement Planning": render_replacement_planning,
    }


def _set_navigation(page):
    st.session_state["navigation"] = page


def render_navigation(df, asset_type, container=None):
    """Render the five production pages as compact native sidebar buttons."""
    if container is not None:
        target = container.container()
    else:
        target = st.sidebar
    current_page = st.session_state.get("navigation", "Overview")
    if current_page not in {label for _, label in NAVIGATION_ITEMS}:
        current_page = "Overview"
    target.markdown('<div class="al-nav-section-label">Overview</div>', unsafe_allow_html=True)
    overview_icon = NAVIGATION_ITEMS[0][0]
    target.button(
        f"{overview_icon}  Overview", key="navigation-overview-button",
        type="primary" if current_page == "Overview" else "secondary",
        width="stretch", on_click=_set_navigation, args=("Overview",),
    )
    target.markdown('<div class="al-nav-section-label">Analysis</div>', unsafe_allow_html=True)
    analysis_pages = ["Asset Explorer", "Lifecycle & Warranty", "Data Audit", "Replacement Planning"]
    icons = {label: icon for icon, label in NAVIGATION_ITEMS}
    for page in analysis_pages:
        target.button(
            f"{icons[page]}  {page}", key=f"navigation-{page}",
            type="primary" if current_page == page else "secondary",
            width="stretch", on_click=_set_navigation, args=(page,),
        )


def render_selected_page(df, asset_type):
    """Render the page selected by the already-rendered global navigation."""
    page = st.session_state.get("navigation", "Overview")
    _page_renderers().get(page, render_overview)(df, asset_type)
