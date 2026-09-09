# AssetLens Engineering and Recovery Record

## Purpose

This document records the engineering baseline, cleanup history, known limitations, and release-readiness evidence for AssetLens. It is an operational record, not a substitute for real-data user acceptance testing.

## Phase 0-7 Milestone Summary

Phases 1-7 are complete and deployed. The delivered application supports official Workstation, Smartphone, and Tablet exports; canonical data normalization; lifecycle and warranty analysis; audit findings; replacement planning; filtered views; and Excel export.

## Phase 8 History

### Phase 8A - Deep Audit

Completed as an audit-only review. The baseline at that point was 65 passing tests, successful compilation and diff checks, and a mixed entry-point architecture. Findings included duplicate entry-point logic, legacy prototype paths, HTML-rendering risk, unused dependencies, documentation drift, and repository hygiene gaps. These findings were retained as historical evidence and resolved or downgraded below after cleanup.

### Phase 8B.1 - Safe Hardening

Completed. Presentation values use HTML escaping where raw HTML is intentionally rendered, Excel exports sanitize formula-like cell values, searches remain literal rather than regex-based, and the source workflow remains read-only.

Historical HTML injection finding: RESOLVED for the reviewed active rendering paths. Future raw HTML interpolation must preserve the same escaping contract.

### Phase 8B.2 - Canonical Duplication Removal

Completed. Canonical data, audit, export, and UI responsibilities are now owned by the `itam` package. The entry point delegates to those modules for detection, canonicalization, lifecycle, warranty, audit, and planning behavior. No active duplicate canonical business logic remains.

Historical duplicate canonical implementations finding: RESOLVED.

### Phase 8B.3 - Legacy Prototype Cleanup

Completed. The legacy prototype and `modular_*` bridge are absent, the entry point is reduced to 431 lines, and the active application uses the canonical module-backed workflow. `use_container_width` is absent from the application source.

Historical legacy prototype and deprecated widget-parameter findings: RESOLVED.

## Final Architecture

- `asset_dashboard.py`: bootstrap, upload handling, validation, orchestration, and page dispatch.
- `itam/data.py`: header and dataset detection, canonicalization, and literal search.
- `itam/audit.py`: lifecycle, warranty, audit findings, and replacement planning.
- `itam/export.py`: Excel export and formula-injection sanitization.
- `itam/ui.py`: AssetLens page renderers, display labels, filters, and presentation helpers.
- `test_data_processing.py`, `test_itam_audit.py`: focused standard-library unit tests.

The official company inventory system remains authoritative. AssetLens analyzes uploaded exports and does not update source records.

## Dependency Audit

| Dependency | Classification | Production use | Version strategy |
|---|---|---|---|
| `streamlit` | Direct production | App runtime and UI | Unpinned latest compatible release at install time |
| `pandas` | Direct production | Workbook data processing and analysis | Unpinned latest compatible release at install time |
| `plotly` | Direct production | Charts in the UI | Unpinned latest compatible release at install time |
| `openpyxl` | Direct production | `.xlsx` read/write engine | Unpinned latest compatible release at install time |
| `fuzzywuzzy` | Unused | None | Removed |
| `python-Levenshtein` | Unused | None | Removed |

No separate dev/test dependency is required by the current standard-library `unittest` suite. The minimal manifest is appropriate for Streamlit Community Cloud, but unpinned versions allow deployment-time drift; a future compatibility-tested pinning policy would improve reproducibility without being part of this stabilization phase. No broad upgrade was performed.

## Accepted Technical Debt

- The application remains upload-based and does not integrate with a live API.
- Processing is in memory and there is no database, authentication, or workflow engine.
- Source quality can limit audit quality; ambiguous or incomplete exports may produce Unknown values or require manual header selection.
- There is no automatic source-system update or procurement approval workflow.
- Streamlit reruns and the inline CSS/theme remain maintainability considerations; changing them is outside this stabilization phase.
- Date parsing and pandas behavior can vary with unusual user-provided Excel values.
- Unit tests focus on canonical and presentation contracts rather than full browser interaction.

## Compatibility Symbol Review

`test_data_processing.py` imports `escape` from `asset_dashboard.py`. This is a standard-library compatibility export used to test pure HTML escaping behavior. It is low risk and intentionally retained for this baseline. Migration to a canonical presentation helper may be considered later, but is not justified during final stabilization.

## Test Coverage Review

The existing 66 tests cover dataset and header detection, canonical mapping, lifecycle and warranty boundaries, normalized duplicate detection, literal search, export sanitization, display labels, audit categories, replacement planning, optional fields, and empty results. Unknown dataset handling is covered at the detection contract level.

Remaining gaps are primarily full Streamlit interaction, malformed workbook UX, repeated rerun state behavior, and real master-export UAT. They are documented gaps rather than reasons to add a larger test framework in this phase.

## Repository Hygiene

The repository is configured to ignore Python bytecode, `__pycache__/`, `.venv/`, dotenv files, and common tool caches. The hygiene check must confirm no tracked `__pycache__`, `.pyc`, temporary/debug files, secrets, sample master datasets, or local virtual environment. `.devcontainer/` and `config.toml` are legitimate repository configuration and are retained.

## Final Regression Baseline

- 66 tests: PASS.
- Required module compilation: PASS.
- `git diff --check`: PASS.
- Local Streamlit startup on port 8502: PASS with HTTP 200 and no Python traceback.
- No tracked runtime cache remains.

The exact final command results are recorded in the Phase 8C section below after execution.

## Phase 8C - Final Engineering Baseline

### Final Architecture

The final architecture is the module-backed workflow described above: the entry point handles bootstrap, upload, validation, and orchestration; `itam/data.py`, `itam/audit.py`, `itam/export.py`, and `itam/ui.py` own the canonical application behavior.

### Dependencies

Production dependencies are limited to Streamlit, Pandas, Plotly, and OpenPyXL. The two unused fuzzy-matching dependencies were removed. The current strategy is minimal, direct, and unpinned for Streamlit Community Cloud compatibility, with deployment drift accepted as technical debt.

### Validation Evidence

- Tests: `python -m unittest discover -p "test_*.py"` - PASS, 66 tests.
- Compile: `python -m py_compile asset_dashboard.py itam/data.py itam/audit.py itam/export.py itam/ui.py` - PASS.
- Diff check: `git diff --check` - PASS.
- Local smoke: `.venv\\Scripts\\streamlit.exe run asset_dashboard.py --server.headless true --server.port 8502` - PASS; HTTP 200; no Python traceback; server stopped after verification.

### Documentation Status

README now describes AssetLens, its supported datasets, current workflow, business rules, security boundaries, limitations, installation, testing, deployment, and maintenance process. This report preserves the Phase 8 audit history and records the final baseline.

### Repository Hygiene Status

The final hygiene review confirms ignored local environments and caches, no tracked Python bytecode, no secrets, no accidental sample master datasets, and no temporary debug artifacts. Existing project configuration files are retained.

### Real-Data UAT Checklist and Results

The latest real master exports were exercised on 2026-09-09:

- `IT_Asset ( Sep 8_ 2026 05_01 PM ).xlsx`
- `IT_Smartphones ( Sep 8_ 2026 05_01 PM ).xlsx`
- `IT_Tablets ( Sep 8_ 2026 05_00 PM ).xlsx`

Ingestion, page rendering, literal search, reset controls, canonical processing, count checks, display labels, optional fields, and export round-trips passed for all three datasets. Excel exports reopened successfully with matching filtered row counts and readable labels. Formula-injection sanitization also passed.

| Check | Workstation | Smartphone | Tablet |
|---|---|---|---|
| Upload workbook | PASS | PASS | PASS |
| Header auto-detection | PASS | PASS | PASS |
| Asset type detection | PASS | PASS | PASS |
| Overview | PASS | PASS | PASS |
| Asset Explorer | PASS | PASS | PASS |
| Lifecycle & Warranty | PASS | PASS | PASS |
| Data Audit | PASS | PASS | PASS |
| Replacement Planning | PASS | PASS | PASS |
| Search and reset | PASS | PASS | PASS |
| Filters and optional fields | PASS | PASS | PASS |
| Excel export | PASS | PASS | PASS |

Locked counts matched exactly. Workstation priority counts also matched exactly. The Tablet export emitted a Pandas date-format inference warning for mixed or nonstandard warranty values; invalid values remained Unknown and no count discrepancy resulted. Streamlit AppTest emitted expected bare-mode ScriptRunContext warnings. No Python traceback occurred.

### Release Recommendation

**READY FOR REAL-DATA UAT**

The engineering baseline is stable for controlled UAT. Production deployment is not declared complete; complete the checklist with the three real exports and perform a Streamlit Cloud smoke test before release approval.
