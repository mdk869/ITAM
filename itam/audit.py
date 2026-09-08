import pandas as pd

from itam.data import find_column


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


def get_warranty_status(df):
    """Calculate warranty status."""
    warranty_col = "warranty_expiry" if "warranty_expiry" in df.columns else find_column(df, ["warranty expiry", "warrantyexpiry"])
    if not warranty_col:
        return df, None

    df_temp = df.copy()
    df_temp["Warranty Expiry Date"] = pd.to_datetime(df_temp[warranty_col], errors="coerce")
    today = pd.Timestamp.now().normalize()
    df_temp["Days to Expiry"] = (df_temp["Warranty Expiry Date"] - today).dt.days

    df_temp["Warranty Status"] = "Unknown"
    df_temp.loc[df_temp["Days to Expiry"] < 0, "Warranty Status"] = "Expired"
    df_temp.loc[df_temp["Days to Expiry"].between(0, 90), "Warranty Status"] = "Expiring Soon"
    df_temp.loc[df_temp["Days to Expiry"] > 90, "Warranty Status"] = "Active"

    expired_warranty_df = df_temp[df_temp["Warranty Status"] == "Expired"].copy()
    return df_temp, expired_warranty_df


PLANNING_PRIORITY_RANK = {
    "Priority Review": 1,
    "Standard Planning": 2,
    "Low Operational Priority": 3,
    "Not Candidate": 4,
}


def classify_planning_priority(is_candidate, source_state):
    """Classify the locked Phase 6A replacement planning priority for one asset.

    This never mutates or re-derives the Replacement Candidate rule; it only
    reads the existing candidate flag and the raw source State for triage.
    """
    if not is_candidate:
        return "Not Candidate"
    normalized_state = str(source_state).strip().casefold() if pd.notna(source_state) else ""
    if normalized_state in {"in use", "in repair"}:
        return "Priority Review"
    if normalized_state == "in store":
        return "Standard Planning"
    if normalized_state in {"waiting to dispose", "disposed", "expired"}:
        return "Low Operational Priority"
    return "Standard Planning"


def derive_replacement_planning(df):
    """Add Planning Priority / Planning Rank without mutating source State values."""
    df = df.copy()
    is_candidate = df.get("ITAM Replacement Candidate", pd.Series(False, index=df.index)).fillna(False).astype(bool)
    source_state = df.get("state", pd.Series(pd.NA, index=df.index))
    priority = [classify_planning_priority(candidate, state) for candidate, state in zip(is_candidate, source_state)]
    df["ITAM Planning Priority"] = priority
    df["ITAM Planning Rank"] = pd.Series(priority, index=df.index).map(PLANNING_PRIORITY_RANK).astype("Int64")
    return df


def run_itam_audit(df):
    """Add row-level ITAM audit findings without changing source values."""
    audited_df = df.copy()
    row_findings = [[] for _ in audited_df.index]
    row_positions = pd.Series(range(len(audited_df)), index=audited_df.index)
    severity_priority = {"Critical": 5, "High": 4, "Medium": 3, "Low": 2, "Info": 1, "None": 0}

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
        row_findings[position].append({"rule": rule, "severity": severity, "message": message})

    def add_duplicate_flags(column, flag_column, rule, label):
        normalized = normalized_identifier(column)
        duplicate_mask = normalized.notna() & normalized.duplicated(keep=False)
        audited_df[flag_column] = duplicate_mask.to_numpy()
        for index in audited_df.index[duplicate_mask]:
            add_finding(row_positions.loc[index], rule, "High", f"Duplicate {label}.")

    add_duplicate_flags("asset_tag", "ITAM Duplicate Asset Tag", "DUPLICATE_ASSET_TAG", "Asset Tag")
    add_duplicate_flags("serial_number", "ITAM Duplicate Serial", "DUPLICATE_SERIAL", "Serial Number")

    asset_types = audited_df.get("asset_type", pd.Series(index=audited_df.index, dtype="object"))
    imei_normalized = normalized_identifier("imei")
    imei_applicable = asset_types.astype("string").str.casefold().isin({"smartphone", "tablet"})
    duplicate_imei = imei_applicable & imei_normalized.notna() & imei_normalized.duplicated(keep=False)
    audited_df["ITAM Duplicate IMEI"] = duplicate_imei.to_numpy()
    for index in audited_df.index[duplicate_imei]:
        add_finding(row_positions.loc[index], "DUPLICATE_IMEI", "High", "Duplicate IMEI.")

    core_columns = [("asset_tag", "asset_tag"), ("serial_number", "serial_number"), ("model", "model"), ("purchase_year", "purchase_year")]
    missing_identity = []
    for index, row in audited_df.iterrows():
        missing_fields = [label for column, label in core_columns if column not in audited_df.columns or is_missing(row[column])]
        missing_identity.append(bool(missing_fields))
        if missing_fields:
            add_finding(row_positions.loc[index], "MISSING_CORE_IDENTITY", "Medium", "Missing core identity: " + ", ".join(missing_fields) + ".")
    audited_df["ITAM Missing Core Identity"] = missing_identity

    active_states = {"in use", "in store", "in repair"}
    state_values = audited_df.get("state", pd.Series(index=audited_df.index, dtype="object"))
    lifecycle_values = audited_df.get("ITAM Lifecycle Status", pd.Series(index=audited_df.index, dtype="object"))
    state_mismatch = lifecycle_values.astype("string").eq("Expired") & state_values.astype("string").str.strip().str.casefold().isin(active_states)
    audited_df["ITAM State Review Required"] = state_mismatch.to_numpy()
    for index in audited_df.index[state_mismatch]:
        add_finding(row_positions.loc[index], "LIFECYCLE_STATE_MISMATCH", "Medium", f"ITAM lifecycle is Expired but source State is {audited_df.at[index, 'state']}.")

    warranty_status = audited_df.get("Warranty Status", pd.Series(index=audited_df.index, dtype="object")).astype("string")
    for index in audited_df.index[warranty_status.eq("Expired")]:
        add_finding(row_positions.loc[index], "EXPIRED_WARRANTY", "Low", "Warranty is expired.")
    for index in audited_df.index[warranty_status.eq("Expiring Soon")]:
        add_finding(row_positions.loc[index], "EXPIRING_WARRANTY", "Info", "Warranty is expiring soon.")

    audited_df["ITAM Replacement Candidate"] = lifecycle_values.astype("string").eq("Expired").to_numpy()
    audited_df = derive_replacement_planning(audited_df)
    review_rules = {"DUPLICATE_ASSET_TAG", "DUPLICATE_SERIAL", "DUPLICATE_IMEI", "MISSING_CORE_IDENTITY", "LIFECYCLE_STATE_MISMATCH"}
    finding_counts, highest_severities, review_required, finding_text = [], [], [], []
    for findings in row_findings:
        finding_counts.append(len(findings))
        highest_severities.append(max((finding["severity"] for finding in findings), key=severity_priority.get, default="None"))
        review_required.append(any(finding["rule"] in review_rules for finding in findings))
        finding_text.append("; ".join(f"[{finding['severity']}] {finding['message'].rstrip('.')}" for finding in findings))

    audited_df["ITAM Finding Count"] = finding_counts
    audited_df["ITAM Highest Severity"] = highest_severities
    audited_df["ITAM Review Required"] = review_required
    audited_df["ITAM Audit Findings"] = finding_text
    return audited_df
