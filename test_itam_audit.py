import unittest

import pandas as pd

from asset_dashboard import run_itam_audit


class ItamAuditTests(unittest.TestCase):
    def audit(self, rows):
        return run_itam_audit(pd.DataFrame(rows))

    def test_clean_asset(self):
        result = self.audit({
            "asset_type": ["Workstation"], "asset_tag": ["WS001"],
            "serial_number": ["SN001"], "model": ["Dell"],
            "purchase_year": [2025], "state": ["In Use"],
            "ITAM Lifecycle Status": ["Active"], "Warranty Status": ["Active"],
        })
        self.assertEqual(result.loc[0, "ITAM Finding Count"], 0)
        self.assertEqual(result.loc[0, "ITAM Highest Severity"], "None")
        self.assertFalse(result.loc[0, "ITAM Review Required"])
        self.assertEqual(result.loc[0, "ITAM Audit Findings"], "")

    def test_duplicate_asset_tag_and_serial_are_normalized(self):
        result = self.audit({
            "asset_type": ["Workstation", "Workstation"],
            "asset_tag": ["ABC123", " abc123 "],
            "serial_number": ["SN001", " sn001 "],
            "model": ["Dell", "Dell"], "purchase_year": [2025, 2025],
            "state": ["In Use", "In Use"],
            "ITAM Lifecycle Status": ["Active", "Active"],
        })
        self.assertTrue(result["ITAM Duplicate Asset Tag"].all())
        self.assertTrue(result["ITAM Duplicate Serial"].all())
        self.assertTrue(result["ITAM Review Required"].all())
        self.assertTrue(result["ITAM Highest Severity"].eq("High").all())

    def test_blank_identifiers_are_not_duplicates_but_are_missing_identity(self):
        result = self.audit({
            "asset_type": ["Workstation", "Workstation"],
            "asset_tag": ["", "NA"], "serial_number": ["SN001", "SN002"],
            "model": ["Dell", "Dell"], "purchase_year": [2025, 2025],
            "state": ["In Use", "In Use"],
            "ITAM Lifecycle Status": ["Active", "Active"],
        })
        self.assertFalse(result["ITAM Duplicate Asset Tag"].any())
        self.assertTrue(result["ITAM Missing Core Identity"].all())
        self.assertTrue(result["ITAM Review Required"].all())

    def test_imei_duplicates_only_apply_to_mobile_assets(self):
        result = self.audit({
            "asset_type": ["Smartphone", "Tablet", "Workstation"],
            "asset_tag": ["M1", "M2", "W1"],
            "serial_number": ["S1", "S2", "S3"],
            "model": ["Phone", "Tablet", "PC"], "purchase_year": [2025] * 3,
            "imei": ["IMEI1", " imei1 ", "IMEI1"],
            "state": ["In Use"] * 3,
            "ITAM Lifecycle Status": ["Active"] * 3,
        })
        self.assertTrue(result.loc[0, "ITAM Duplicate IMEI"])
        self.assertTrue(result.loc[1, "ITAM Duplicate IMEI"])
        self.assertFalse(result.loc[2, "ITAM Duplicate IMEI"])

    def test_missing_identity_and_lifecycle_state_mismatch_are_medium(self):
        result = self.audit({
            "asset_type": ["Workstation"], "asset_tag": ["W1"],
            "serial_number": [""], "model": ["Dell"], "purchase_year": ["NA"],
            "state": ["In Use"], "ITAM Lifecycle Status": ["Expired"],
        })
        self.assertTrue(result.loc[0, "ITAM Missing Core Identity"])
        self.assertTrue(result.loc[0, "ITAM State Review Required"])
        self.assertEqual(result.loc[0, "ITAM Highest Severity"], "Medium")
        self.assertTrue(result.loc[0, "ITAM Review Required"])

    def test_disposal_state_does_not_mismatch_and_expired_is_replacement_candidate(self):
        result = self.audit({
            "asset_type": ["Workstation", "Workstation", "Workstation"],
            "asset_tag": ["W1", "W2", "W3"],
            "serial_number": ["S1", "S2", "S3"],
            "model": ["Dell"] * 3, "purchase_year": [2018] * 3,
            "state": ["Waiting To Dispose", "Disposed", "In Use"],
            "ITAM Lifecycle Status": ["Expired"] * 3,
        })
        self.assertFalse(result.loc[0, "ITAM State Review Required"])
        self.assertFalse(result.loc[1, "ITAM State Review Required"])
        self.assertTrue(result["ITAM Replacement Candidate"].all())
        self.assertTrue(result.loc[2, "ITAM State Review Required"])

    def test_warranty_findings_do_not_require_review(self):
        result = self.audit({
            "asset_type": ["Workstation", "Workstation"],
            "asset_tag": ["W1", "W2"], "serial_number": ["S1", "S2"],
            "model": ["Dell", "Dell"], "purchase_year": [2025, 2025],
            "state": ["In Use", "In Use"],
            "ITAM Lifecycle Status": ["Active", "Active"],
            "Warranty Status": ["Expired", "Expiring Soon"],
        })
        self.assertEqual(result.loc[0, "ITAM Highest Severity"], "Low")
        self.assertEqual(result.loc[0, "ITAM Finding Count"], 1)
        self.assertFalse(result["ITAM Review Required"].any())
        self.assertEqual(result.loc[1, "ITAM Highest Severity"], "Info")

    def test_waiting_to_dispose_and_disposed_are_lifecycle_exceptions(self):
        result = self.audit({
            "asset_type": ["Workstation", "Workstation"],
            "asset_tag": ["W1", "W2"], "serial_number": ["S1", "S2"],
            "model": ["Dell", "Dell"], "purchase_year": [2018, 2018],
            "state": ["Waiting To Dispose", "Disposed"],
            "ITAM Lifecycle Status": ["Expired", "Expired"],
        })
        self.assertFalse(result["ITAM State Review Required"].any())

    def test_highest_severity_uses_priority_order(self):
        result = self.audit({
            "asset_type": ["Workstation"], "asset_tag": ["W1"],
            "serial_number": ["S1"], "model": ["Dell"],
            "purchase_year": [2025], "state": ["In Use"],
            "ITAM Lifecycle Status": ["Expired"],
            "Warranty Status": ["Expired"],
        })
        self.assertEqual(result.loc[0, "ITAM Highest Severity"], "Medium")
        self.assertEqual(result.loc[0, "ITAM Finding Count"], 2)

    def test_not_assigned_is_not_generic_missing_identity(self):
        result = self.audit({
            "asset_type": ["Workstation"], "asset_tag": ["Not Assigned"],
            "serial_number": ["S1"], "model": ["Dell"],
            "purchase_year": [2025], "state": ["In Use"],
            "ITAM Lifecycle Status": ["Active"],
        })
        self.assertFalse(result.loc[0, "ITAM Missing Core Identity"])


if __name__ == "__main__":
    unittest.main()
