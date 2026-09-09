# AssetLens
IT Asset Audit & Lifecycle Intelligence

## Purpose

AssetLens is a read-only audit, analysis, and replacement-planning tool for official company asset exports. It processes uploaded Excel files in memory and adds lifecycle, warranty, data-quality, and planning analysis without changing source records.

The official company inventory system remains the authoritative source.

## Core Workflow

Official Inventory
→ Export
→ AssetLens
→ Audit / Analysis
→ Findings
→ Replacement Planning
→ Cross-check main system

## Current Features

- Overview
- Asset Explorer
- Lifecycle & Warranty
- Data Audit
- Replacement Planning
- Excel export

## Supported Datasets

AssetLens supports these official export types:

- Workstation
- Smartphone
- Tablet

The app scans the first rows of the selected worksheet for a recognizable header. It detects the asset type from the export-specific type columns and maps source fields into a consistent internal schema while preserving the source columns for analysis and export.

## Lifecycle Rules

- 0–1 years: New
- 2–3 years: Active
- 4–5 years: Aging
- More than 5 years: Expired
- Invalid or missing purchase year: Unknown

## Warranty Rules

- Expiry before today: Expired
- 0–90 days remaining: Expiring Soon
- More than 90 days remaining: Active
- Missing or invalid expiry: Unknown

## Data Audit

The audit identifies duplicate Asset Tag, duplicate Serial, duplicate IMEI for mobile assets, missing core identity, and lifecycle/source-state mismatches. Findings receive a severity, and identity or state issues are marked Review Required.

AssetLens does not modify source records.

## Replacement Planning

Assets with lifecycle status Expired are marked Replacement Candidate.

Planning priority is classified as:

- Priority Review
- Standard Planning
- Low Operational Priority

Replacement Candidate does not mean automatic replacement approval. It is an analysis and planning signal for follow-up against the authoritative inventory system and organizational process.

## Architecture

```text
asset_dashboard.py
→ bootstrap/upload/orchestration

itam/data.py
→ detection/canonicalization/search

itam/audit.py
→ lifecycle/warranty/audit/planning

itam/export.py
→ Excel export/sanitization

itam/ui.py
→ AssetLens pages/UI
```

## Installation

```bash
python -m venv .venv
```

Windows:

```powershell
.venv\Scripts\activate
```

```bash
pip install -r requirements.txt
streamlit run asset_dashboard.py
```

## Testing

```bash
python -m unittest discover -p "test_*.py"
python -m py_compile asset_dashboard.py itam/data.py itam/audit.py itam/export.py itam/ui.py
```

## Deployment

Deploy from GitHub to Streamlit Community Cloud. The production entry point is `asset_dashboard.py`.

Do not place credentials or secrets in the repository or upload them through the application.

## Security

- Search is literal and non-regex.
- User-controlled values are HTML-escaped at presentation boundaries.
- Excel formula-injection protection is applied during export.
- The source workflow is read-only.
- AssetLens itself does not persist uploaded datasets.

## Limitations

- Upload-based rather than a live API integration
- In-memory processing
- Audit quality depends on source data quality
- No database, authentication, or workflow engine
- No automatic source-system update
- No procurement approval workflow

## Maintenance

Use this change sequence:

```text
change
→ tests
→ compile
→ local Streamlit smoke
→ commit/push
→ Streamlit Cloud smoke test
```
