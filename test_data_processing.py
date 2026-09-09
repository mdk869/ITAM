import unittest
from datetime import timedelta
from io import BytesIO

import pandas as pd

from asset_dashboard import escape
from itam.audit import calculate_asset_age, get_warranty_status
from itam.data import (
    CANONICAL_COLUMNS,
    apply_literal_search,
    build_canonical_dataframe,
    detect_asset_type,
    detect_asset_type_from_data,
    detect_header_row,
    validate_source_columns,
)
from itam.export import export_to_excel
from itam.ui import (
    CANDIDATE_DETAIL_COLUMNS,
    CANDIDATE_DETAIL_OPTIONAL_COLUMNS,
    DISPLAY_LABELS,
    AUDIT_SEARCH_COLUMNS,
    AUDIT_CURATED_COLUMNS,
    PLANNING_FILTER_RESET_COLUMNS,
    PLANNING_SEARCH_COLUMNS,
    _clear_audit_filters,
    _chronological_counts,
    _clear_replacement_planning_filters,
    add_finding_categories,
    audit_table_columns,
    clean_asset_count,
    _meaningful_optional_columns,
    _top_n_counts,
    _values,
    format_audit_note,
    make_display_dataframe,
    prepare_export_dataframe,
    resolve_export_columns,
)


def excel_with_rows(rows):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, index=False, header=False, sheet_name="Assets")
    output.seek(0)
    return output


class HeaderAndAssetDetectionTests(unittest.TestCase):
    def test_supported_headers_are_found_at_row_seven(self):
        metadata = [["metadata"]] * 7
        for header in [
            ["Workstation Type", "Serial Number", "Model", "Asset Tag", "State", "Year Of Purchase"],
            ["Product Type", "Serial Number", "Product", "AssetTag", "State", "Year Of Purchase"],
        ]:
            self.assertEqual(detect_header_row(excel_with_rows(metadata + [header]), "Assets"), 7)

    def test_header_normalization_and_unsupported_workbook(self):
        header = [[" Product-Type ", "Serial.Number", " Product ", "Asset Tag", "State", "Year Of Purchase"]]
        self.assertEqual(detect_header_row(excel_with_rows([["meta"]] * 7 + header), "Assets"), 7)
        self.assertIsNone(detect_header_row(excel_with_rows([["name", "color", "amount"]]), "Assets"))

    def test_asset_type_detection_is_exact(self):
        self.assertEqual(detect_asset_type(["Workstation Type"]), "Workstation")
        self.assertEqual(
            detect_asset_type_from_data(pd.DataFrame({"Product Type": ["IT Smartphones"]})),
            "Smartphone",
        )
        self.assertEqual(
            detect_asset_type_from_data(pd.DataFrame({"Product Type": ["IT Tablets"]})),
            "Tablet",
        )
        self.assertEqual(
            detect_asset_type_from_data(pd.DataFrame({"Product Type": ["Smartphone"]})),
            "Unknown",
        )


class CanonicalSchemaTests(unittest.TestCase):
    def workstation(self):
        return pd.DataFrame({
            "Workstation Type": ["Laptop"], "Model": ["Dell"], "Asset Tag": ["W1"],
            "Serial Number": ["S1"], "State": ["In Use"], "User Employee ID": ["E1"],
            "User Email": ["e@example.com"], "User Jobtitle": ["Engineer"],
            "Department": ["IT"], "Site": ["HQ"], "Year Of Purchase": [2025],
            "Warranty Expiry": ["2030-01-01"], "Workstation Status": ["Active"],
        })

    def mobile(self):
        return pd.DataFrame({
            "Product Type": ["IT Smartphones"], "Product": ["Phone"], "AssetTag": ["M1"],
            "Serial Number": ["S1"], "State": ["In Use"], "User -> Employee ID": ["E1"],
            "User -> Email": ["e@example.com"], "User -> Job Title": ["User"],
            "User -> Department": ["IT"], "User -> Site": ["HQ"],
            "Year Of Purchase": [2025], "Warranty Expiry Date": ["2030-01-01"],
            "No. IMEI": ["I1"], "No. Sim": ["SIM1"],
        })

    def test_workstation_mapping_preserves_raw_columns(self):
        source = self.workstation()
        result = build_canonical_dataframe(source, "Workstation")
        expected = {
            "asset_type": "Workstation", "model": "Dell", "asset_tag": "W1",
            "serial_number": "S1", "employee_id": "E1", "email": "e@example.com",
            "job_title": "Engineer", "department": "IT", "site": "HQ",
            "purchase_year": 2025, "warranty_expiry": "2030-01-01",
            "workstation_status": "Active", "source_asset_subtype": "Laptop",
        }
        for column, value in expected.items():
            self.assertEqual(result.loc[0, column], value)
        self.assertTrue(pd.isna(result.loc[0, "imei"]))
        self.assertTrue(pd.isna(result.loc[0, "sim_number"]))
        self.assertIn("Model", result.columns)

    def test_mobile_mapping_and_tablet_type(self):
        source = self.mobile()
        for asset_type in ["Smartphone", "Tablet"]:
            result = build_canonical_dataframe(source, asset_type)
            self.assertEqual(result.loc[0, "asset_type"], asset_type)
            self.assertEqual(result.loc[0, "model"], "Phone")
            self.assertEqual(result.loc[0, "asset_tag"], "M1")
            self.assertEqual(result.loc[0, "employee_id"], "E1")
            self.assertEqual(result.loc[0, "email"], "e@example.com")
            self.assertEqual(result.loc[0, "job_title"], "User")
            self.assertEqual(result.loc[0, "department"], "IT")
            self.assertEqual(result.loc[0, "site"], "HQ")
            self.assertEqual(result.loc[0, "warranty_expiry"], "2030-01-01")
            self.assertEqual(result.loc[0, "imei"], "I1")
            self.assertEqual(result.loc[0, "sim_number"], "SIM1")
            self.assertTrue(pd.isna(result.loc[0, "workstation_status"]))

    def test_required_columns_report_only_missing_required_headers(self):
        source = self.workstation().drop(columns=["Model"])
        self.assertEqual(validate_source_columns(source, "Workstation"), ["Model"])
        self.assertEqual(
            validate_source_columns(self.mobile().drop(columns=["AssetTag"]), "Smartphone"),
            ["AssetTag"],
        )


class LifecycleAndWarrantyTests(unittest.TestCase):
    def test_lifecycle_boundaries(self):
        current_year = pd.Timestamp.now().year
        years = [current_year - age for age in range(7)]
        result = calculate_asset_age(pd.DataFrame({"purchase_year": years}))
        self.assertEqual(result["ITAM Lifecycle Status"].tolist(), ["New", "New", "Active", "Active", "Aging", "Aging", "Expired"])

    def test_invalid_blank_and_future_years_are_unknown(self):
        year = pd.Timestamp.now().year
        result = calculate_asset_age(pd.DataFrame({"purchase_year": ["bad", "", year + 1]}))
        self.assertEqual(result["ITAM Lifecycle Status"].tolist(), ["Unknown", "Unknown", "Unknown"])

    def test_warranty_boundaries_and_invalid_values(self):
        today = pd.Timestamp.now().normalize()
        values = [today - timedelta(days=1), today, today + timedelta(days=90), today + timedelta(days=91), "", "invalid"]
        result, _ = get_warranty_status(pd.DataFrame({"warranty_expiry": values}))
        self.assertEqual(result["Warranty Status"].tolist(), ["Expired", "Expiring Soon", "Expiring Soon", "Active", "Unknown", "Unknown"])

    def test_warranty_missing_date_is_unknown(self):
        result, _ = get_warranty_status(pd.DataFrame({"warranty_expiry": [pd.NA]}))
        self.assertEqual(result.loc[0, "Warranty Status"], "Unknown")


class SearchExportAndEscapingTests(unittest.TestCase):
    def test_search_is_literal_and_case_insensitive(self):
        source = pd.DataFrame({"value": ["ABC[123]", "A+B", "SN.001", "other"]})
        for query in ["ABC[123]", "A+B", "SN.001"]:
            self.assertEqual(len(apply_literal_search(source, query)), 1)
        self.assertEqual(len(apply_literal_search(source, "abc[123]")), 1)

    def test_excel_formula_prefixes_are_neutralized_and_normal_values_unchanged(self):
        source = pd.DataFrame({"value": ["=SUM(A1:A2)", "+cmd", "-123", "@something", "normal"]})
        exported = pd.read_excel(export_to_excel(source), engine="openpyxl")
        self.assertEqual(exported["value"].tolist(), ["'=SUM(A1:A2)", "'+cmd", "'-123", "'@something", "normal"])

    def test_html_escaping_is_available_as_a_pure_helper(self):
        self.assertNotIn("<script>", escape("<script>alert(1)</script>"))

    def test_audit_search_is_literal_and_limited_to_audit_fields(self):
        source = pd.DataFrame({
            "asset_tag": ["ABC[123]", "OTHER"],
            "model": ["Dell", "A+B"],
            "ITAM Audit Findings": ["", "Needs review"],
            "unrelated": ["A+B", "ABC[123]"],
        })
        search_columns = [column for column in AUDIT_SEARCH_COLUMNS if column in source.columns]
        result = apply_literal_search(source, "ABC[123]", columns=search_columns)
        self.assertEqual(result.index.tolist(), [0])


class AuditPresentationTests(unittest.TestCase):
    def test_finding_categories_are_presentation_only_and_prioritized(self):
        source = pd.DataFrame({
            "ITAM Duplicate Serial": [True, False, False, False, False],
            "ITAM Missing Core Identity": [False, True, False, False, False],
            "ITAM State Review Required": [False, False, True, False, False],
            "Warranty Status": ["Active", "Active", "Active", "Expired", "Active"],
            "ITAM Finding Count": [1, 1, 1, 1, 0],
        })
        original_columns = list(source.columns)
        result = add_finding_categories(source)
        self.assertEqual(result["Finding Category"].tolist(), [
            "Duplicate Identity", "Missing Core Identity", "Lifecycle / State Review", "Warranty", "Clean",
        ])
        self.assertEqual(list(source.columns), original_columns)

    def test_clean_asset_count_uses_finding_count(self):
        self.assertEqual(clean_asset_count(pd.DataFrame({"ITAM Finding Count": [0, 1, 0]})), 2)

    def test_audit_filter_callback_restores_defaults(self):
        import unittest.mock
        state = {"audit-severity": ["High"], "audit-review": "Yes", "audit-category": ["Warranty"], "audit-type": ["Tablet"], "audit-state": ["In Use"], "audit-site": ["HQ"], "audit-search": "old"}
        with unittest.mock.patch("itam.ui.st.session_state", state):
            _clear_audit_filters()
        self.assertEqual(state, {"audit-severity": [], "audit-review": "All", "audit-category": [], "audit-type": [], "audit-state": [], "audit-site": [], "audit-search": ""})

    def test_imei_is_optional_in_curated_audit_columns(self):
        with_imei = pd.DataFrame({"imei": ["123"], "asset_tag": ["A"]})
        without_imei = pd.DataFrame({"imei": [pd.NA], "asset_tag": ["A"]})
        self.assertIn("imei", audit_table_columns(with_imei))
        self.assertNotIn("imei", audit_table_columns(without_imei))

    def test_audit_export_uses_curated_columns_and_readable_labels(self):
        source = pd.DataFrame({
            "ITAM Highest Severity": ["High"], "ITAM Review Required": [True],
            "asset_type": ["Workstation"], "asset_tag": ["A1"], "serial_number": ["S1"],
            "model": ["Dell"], "state": ["In Use"], "ITAM Lifecycle Status": ["Active"],
            "Finding Category": ["Duplicate Identity"],
            "ITAM Audit Findings": ["[High] Duplicate Serial"], "site": ["HQ"],
            "department": ["IT"], "imei": [pd.NA], "unrelated": ["hidden"],
        })
        columns = audit_table_columns(source)
        exported = prepare_export_dataframe(source, columns)
        self.assertEqual(columns, [column for column in AUDIT_CURATED_COLUMNS if column != "imei"])
        self.assertEqual(list(exported.columns), [
            "Severity", "Review Required", "Asset Type", "Asset Tag", "Serial Number",
            "Model / Product", "Source State", "Lifecycle", "Finding Category",
            "Audit Notes", "Site", "Department",
        ])
        self.assertNotIn("ITAM", " ".join(exported.columns))
        self.assertEqual(exported.loc[0, "Audit Notes"], "[High] Duplicate Serial")

    def test_empty_audit_export_dataframe_has_no_rows(self):
        source = pd.DataFrame(columns=["asset_tag", "ITAM Audit Findings"])
        exported = prepare_export_dataframe(source, audit_table_columns(source))
        self.assertIsNotNone(exported)
        self.assertTrue(exported.empty)


class DisplayColumnTests(unittest.TestCase):
    def test_audit_note_presentation_removes_internal_wording(self):
        note = "[Medium] ITAM lifecycle is Expired but source State is In Store"

        formatted = format_audit_note(note)

        self.assertEqual(formatted, "[Medium] Lifecycle is Expired but Source State is In Store")
        self.assertIn("[Medium]", formatted)
        self.assertIn("In Store", formatted)
        self.assertNotIn("ITAM lifecycle", formatted)

    def test_audit_note_presentation_handles_blank_and_null_values(self):
        self.assertEqual(format_audit_note(""), "")
        self.assertTrue(pd.isna(format_audit_note(pd.NA)))

    def test_display_labels_do_not_expose_internal_itam_prefix(self):
        self.assertTrue(all("itam" not in label.casefold() for label in DISPLAY_LABELS.values()))
        self.assertEqual(DISPLAY_LABELS["ITAM Lifecycle Status"], "Lifecycle")
        self.assertEqual(DISPLAY_LABELS["ITAM Audit Findings"], "Audit Notes")
        self.assertEqual(DISPLAY_LABELS["ITAM Duplicate Asset Tag"], "Duplicate Asset Tag")

    def test_workstation_display_labels_are_unique_and_preserve_collisions(self):
        source = CanonicalSchemaTests().workstation()
        source["Workstation Status"] = "Raw Status"
        processed = build_canonical_dataframe(source, "Workstation")
        original_columns = list(processed.columns)

        display = make_display_dataframe(processed)

        self.assertTrue(display.columns.is_unique)
        self.assertEqual(display["Workstation Status"].tolist(), ["Raw Status"])
        self.assertEqual(display["Workstation Status (Canonical)"].tolist(), ["Raw Status"])
        self.assertEqual(display["Year Of Purchase"].tolist(), [2025])
        self.assertEqual(display["Year Of Purchase (Canonical)"].tolist(), [2025])
        self.assertEqual(display["Warranty Expiry"].tolist(), ["2030-01-01"])
        self.assertEqual(display["Warranty Expiry (Canonical)"].tolist(), ["2030-01-01"])
        self.assertEqual(list(processed.columns), original_columns)

    def test_mobile_display_labels_are_unique_for_smartphone_and_tablet(self):
        source = CanonicalSchemaTests().mobile()
        for asset_type in ["Smartphone", "Tablet"]:
            processed = build_canonical_dataframe(source, asset_type)
            display = make_display_dataframe(processed)

            self.assertTrue(display.columns.is_unique)
            self.assertIn("Serial Number", display.columns)
            self.assertIn("Serial Number (Canonical)", display.columns)
            self.assertIn("Year Of Purchase", display.columns)
            self.assertIn("Year Of Purchase (Canonical)", display.columns)
            self.assertEqual(display["Serial Number"].tolist(), ["S1"])
            self.assertEqual(display["Serial Number (Canonical)"].tolist(), ["S1"])


class CustomExportTests(unittest.TestCase):
    def test_standard_preset_resolves_available_fields_in_order(self):
        source = CanonicalSchemaTests().mobile()
        processed = build_canonical_dataframe(source, "Smartphone")

        columns = resolve_export_columns(processed, "Standard Asset View")

        self.assertEqual(columns[:5], ["asset_type", "asset_tag", "serial_number", "model", "state"])
        self.assertNotIn("workstation_status", columns)

    def test_custom_export_preserves_selection_order_and_readable_labels(self):
        source = CanonicalSchemaTests().workstation()
        processed = build_canonical_dataframe(source, "Workstation")
        processed = calculate_asset_age(processed)
        selected = ["serial_number", "asset_tag", "ITAM Lifecycle Status"]

        exported = prepare_export_dataframe(processed, selected)

        self.assertEqual(list(exported.columns), ["Serial Number", "Asset Tag", "Lifecycle"])
        self.assertEqual(exported["Serial Number"].tolist(), ["S1"])

    def test_optional_preset_fields_and_empty_selection_are_safe(self):
        source = CanonicalSchemaTests().mobile()
        processed = build_canonical_dataframe(source, "Tablet")

        replacement_columns = resolve_export_columns(processed, "Replacement Planning")

        self.assertTrue(replacement_columns)
        self.assertNotIn("workstation_status", replacement_columns)
        self.assertEqual(resolve_export_columns(processed, "Custom", []), [])
        self.assertIsNone(prepare_export_dataframe(processed, []))


class BreakdownHelperTests(unittest.TestCase):
    def test_top_n_counts_is_descending_and_blank_filtered(self):
        series = pd.Series(["Dell", "Dell", "HP", "HP", "HP", "", "  ", "NA", None])

        counts = _top_n_counts(series)

        self.assertEqual(counts.index.tolist(), ["HP", "Dell"])
        self.assertEqual(counts.tolist(), [3, 2])

    def test_top_n_counts_respects_top_n_limit(self):
        series = pd.Series(["A", "B", "B", "C", "C", "C", "D", "D", "D", "D"])

        counts = _top_n_counts(series, top_n=2)

        self.assertEqual(counts.index.tolist(), ["D", "C"])

    def test_top_n_counts_empty_for_all_blank_series(self):
        series = pd.Series(["", "NA", None, "na"])

        counts = _top_n_counts(series)

        self.assertTrue(counts.empty)

    def test_chronological_counts_orders_numeric_values_ascending(self):
        series = pd.Series([2020, 2018, 2018, 2022, "", None])

        counts = _chronological_counts(series)

        self.assertEqual(counts.index.tolist(), ["2018", "2020", "2022"])
        self.assertEqual(counts.tolist(), [2, 1, 1])

    def test_chronological_counts_empty_for_blank_series(self):
        series = pd.Series([None, "", "NA"])

        self.assertTrue(_chronological_counts(series).empty)

    def test_export_preset_headings_are_clean_and_unique(self):
        source = CanonicalSchemaTests().workstation()
        processed = build_canonical_dataframe(source, "Workstation")
        processed = calculate_asset_age(processed)
        processed["ITAM Review Required"] = False
        processed["ITAM Highest Severity"] = "None"
        processed["ITAM Audit Findings"] = ""

        columns = resolve_export_columns(processed, "Audit Findings")
        exported = prepare_export_dataframe(processed, columns)

        self.assertTrue(exported.columns.is_unique)
        self.assertTrue(all("itam" not in label.casefold() for label in exported.columns))
        self.assertIn("Lifecycle", exported.columns)
        self.assertIn("Audit Notes", exported.columns)


class ReplacementPlanningFilterAndSearchTests(unittest.TestCase):
    def candidates(self):
        return pd.DataFrame({
            "asset_type": ["Workstation"] * 4,
            "asset_tag": ["W1", "W2", "W3", "W4"],
            "serial_number": ["SN.001", "SN002", "SN003", "SN004"],
            "model": ["Dell Latitude", "HP EliteBook", "Dell Latitude", "Lenovo"],
            "user": ["Alice", "Bob", "A+B Team", "Not Assigned"],
            "employee_id": ["E1", "E2", "E3", None],
            "department": ["IT", "Finance", "", "Not Assigned"],
            "site": ["HQ", "Branch", "HQ", None],
            "workstation_status": ["In Use", "In Store", "In Use", ""],
            "Asset Age": [7, 6, 9, 11],
            "ITAM Planning Rank": [1, 1, 1, 1],
        })

    def search_columns(self, df):
        """Mirror the production guard that only searches columns present in the current frame."""
        return [column for column in PLANNING_SEARCH_COLUMNS if column in df.columns]

    def test_candidate_search_is_literal_case_insensitive_and_regex_safe(self):
        source = self.candidates()
        columns = self.search_columns(source)
        for query in ["SN.001", "A+B Team", "sn.001"]:
            result = apply_literal_search(source, query, columns=columns)
            self.assertEqual(len(result), 1)
        # A literal "." must only match rows containing an actual dot, not act as a regex wildcard.
        dot_result = apply_literal_search(source, ".", columns=columns)
        self.assertEqual(dot_result["asset_tag"].tolist(), ["W1"])
        # Other regex metacharacters must not raise and must not be interpreted as regex.
        for pattern in ["*", "?", "[", "]", "(", ")"]:
            self.assertTrue(apply_literal_search(source, pattern, columns=columns).empty)
        # "+" is present literally in "A+B Team", so it must match that single row, not be treated as regex.
        plus_result = apply_literal_search(source, "+", columns=columns)
        self.assertEqual(plus_result["asset_tag"].tolist(), ["W3"])

    def test_candidate_search_matches_across_multiple_fields(self):
        source = self.candidates()
        columns = self.search_columns(source)
        self.assertEqual(len(apply_literal_search(source, "Branch", columns=columns)), 1)
        self.assertEqual(len(apply_literal_search(source, "In Use", columns=columns)), 2)
        self.assertEqual(len(apply_literal_search(source, "Dell Latitude", columns=columns)), 2)

    def test_blank_search_preserves_current_filtered_rows(self):
        source = self.candidates().iloc[[0, 2]]
        result = apply_literal_search(source, "", columns=self.search_columns(source))
        self.assertEqual(len(result), 2)
        self.assertEqual(result["asset_tag"].tolist(), ["W1", "W3"])

    def test_search_does_not_crash_on_null_values(self):
        source = self.candidates()
        result = apply_literal_search(source, "e1", columns=self.search_columns(source))
        self.assertEqual(result["asset_tag"].tolist(), ["W1"])

    def test_filter_option_values_drop_blanks_but_preserve_not_assigned(self):
        source = self.candidates()
        values = _values(source, "department")
        self.assertIn("Not Assigned", values)
        self.assertNotIn("", values)
        self.assertNotIn(None, values)

    def test_segment_filter_and_search_apply_cumulatively(self):
        source = self.candidates()
        segmented = source[source["site"].astype(str).eq("HQ")]
        result = apply_literal_search(segmented, "Dell", columns=self.search_columns(segmented))
        self.assertEqual(result["asset_tag"].tolist(), ["W1", "W3"])

    def test_planning_rank_and_age_sort_is_unchanged(self):
        source = self.candidates()
        sorted_df = source.sort_values(["ITAM Planning Rank", "Asset Age"], ascending=[True, False])
        self.assertEqual(sorted_df["asset_tag"].tolist(), ["W4", "W3", "W1", "W2"])

    def test_meaningful_optional_columns_drops_empty_optional_fields(self):
        smartphone_df = pd.DataFrame({"asset_tag": ["M1"], "imei": ["I1"], "workstation_status": [pd.NA]})
        visible = _meaningful_optional_columns(
            smartphone_df, ["asset_tag", "workstation_status", "imei"], CANDIDATE_DETAIL_OPTIONAL_COLUMNS,
        )
        self.assertEqual(visible, ["asset_tag", "imei"])

        workstation_df = pd.DataFrame({"asset_tag": ["W1"], "imei": [pd.NA], "workstation_status": ["In Use"]})
        visible = _meaningful_optional_columns(
            workstation_df, ["asset_tag", "workstation_status", "imei"], CANDIDATE_DETAIL_OPTIONAL_COLUMNS,
        )
        self.assertEqual(visible, ["asset_tag", "workstation_status"])

    def test_candidate_detail_columns_lead_with_planning_priority(self):
        self.assertEqual(CANDIDATE_DETAIL_COLUMNS[0], "ITAM Planning Priority")
        self.assertNotIn("ITAM Planning Rank", CANDIDATE_DETAIL_COLUMNS)
        self.assertNotIn("ITAM Finding Count", CANDIDATE_DETAIL_COLUMNS)

    def test_clear_replacement_planning_filters_resets_widget_defaults(self):
        import streamlit as st

        st.session_state["replacement-segment"] = "Priority Review"
        st.session_state["replacement-search"] = "abc"
        for column in PLANNING_FILTER_RESET_COLUMNS:
            st.session_state[f"replacement-{column}"] = ["some-value"]

        _clear_replacement_planning_filters()

        self.assertEqual(st.session_state["replacement-segment"], "All Replacement Candidates")
        self.assertEqual(st.session_state["replacement-search"], "")
        for column in PLANNING_FILTER_RESET_COLUMNS:
            self.assertEqual(st.session_state[f"replacement-{column}"], [])


class ExportAuditNoteTests(unittest.TestCase):
    def test_export_audit_notes_are_cleaned_without_mutating_source(self):
        source = pd.DataFrame({
            "ITAM Audit Findings": ["[Medium] ITAM lifecycle is Expired but source State is In Use"],
        })

        exported = prepare_export_dataframe(source, ["ITAM Audit Findings"])

        self.assertEqual(exported.loc[0, "Audit Notes"], "[Medium] Lifecycle is Expired but Source State is In Use")
        self.assertEqual(source.loc[0, "ITAM Audit Findings"], "[Medium] ITAM lifecycle is Expired but source State is In Use")


if __name__ == "__main__":
    unittest.main()