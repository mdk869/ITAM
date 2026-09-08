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


if __name__ == "__main__":
    unittest.main()