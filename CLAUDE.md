# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

Run the app:
```
streamlit run app.py
```

Install dependencies:
```
pip install -r requirements.txt
```

There is no test suite, linter, or build step configured in this project.

Database credentials are read from a `.env` file (via `python-dotenv`) by `db_connect.py`'s `REQUIRED_DB_VARS` — never hardcode credentials, and never commit `.env`.

## Architecture

This is a Streamlit app that pulls case-level data from a MySQL database and turns it into a "Casewise TAT" (turnaround time) report, delivered as an Excel workbook and/or a PPTX deck. It is a from-scratch reimplementation of the logic in the legacy `PPT_Report_Generator_S3` tool, adapted to a fixed DB schema instead of arbitrary uploaded CSVs, and with all AI/API-key functionality removed.

### Scope-first data loading (`app.py` → `data_loading.py` → `sql_queries.py`)

`app.py` forces the user to pick a scope (one or more specific clients, or an explicit "All Clients Combined" opt-in) and a `months_back` window *before* any query runs. This exists because an unscoped pull is a multi-million-row query — `data_loading.fetch_full_dataset` treats `company_ids=None` as a deliberate slow/batch choice, not a default.

`sql_queries.build_base_data_query()` builds a MySQL CTE query (`_QUERY_BASE_DATA_TEMPLATE`) against the case tables, joined through `ec_client`/`ec_master_company` for the client list. The query was ported from an original Postgres/Redshift-dialect query — when touching it, be aware MySQL differs from that dialect in several ways that have already caused bugs here:
- `GREATEST()`/`LEAST()` return `NULL` if *any* argument is `NULL` (Postgres/Redshift ignore NULLs). Every `GREATEST(...)` in the query wraps its arguments in `COALESCE(x, received_date)` for this reason — don't remove that wrapping.
- Identifiers with spaces/special characters must use backtick-quoting, not double quotes (`ANSI_QUOTES` is not assumed to be set, so a double-quoted identifier silently becomes a string literal instead of erroring).
- `--` comments require a following whitespace character.

The SQL query only builds the base CTEs (through `T2`); it deliberately stops before the holiday-adjusted TAT layers.

### TAT/ageing calculations happen in Python, not SQL (`tat_calculations.py`)

The original design computed TAT/ageing by counting holidays per-case via correlated SQL subqueries, which is a full-table-scan-per-row performance disaster at real data volumes. `tat_calculations.apply_tat_calculations(df, holidays)` replaces that with vectorized `numpy.searchsorted` lookups against precomputed sorted holiday-date arrays (`OffDaySets`, one per TAT type: Working / Calendar / WorkingPlusSaturday). It produces `due_date_recal`, `ageing`, `ageing_bucket`, `tat_status_recal`, `due_date_final`, `ageing_final`, `ageing_bucket_final`, `tat_status_final`, etc. `data_loading.fetch_full_dataset` always calls this after the raw SQL fetch — the SQL layer and this module are a single logical pipeline, not independent stages.

Note: `ec_master_holidays` has no `holiday_type=3` rows in this schema. The `insuff_ageing`/`insuff_ageing_bucket` calculation is therefore defined as elapsed calendar days minus type-1 and type-2 holidays, not a true type-3 lookup.

### Disk cache keyed by scope (`data_cache.py`)

Query results are cached to `cache/query_base_data_{key}.pkl`, where `key` is an MD5 hash of the sorted company-id selection plus `months_back` (`data_cache.make_key`). Different client/month selections get independent cache entries and independent freshness (`DEFAULT_TTL_SECONDS`); there is no single "the cache" to invalidate — clearing/refreshing is always scope-specific.

### Report generation (`report_builder.py` orchestrates everything downstream)

`report_builder.generate_client_report(df, client_name, ...)` is the single entry point used by `app.py` for both single-client and batch (ZIP) generation. It filters by date range, computes TAT/case-status/severity pivots via `data_processor.py` (`TATCalculator`, `CaseStatusAnalyzer`, `SeverityAnalyzer`, `AgeingSummary`, `PendingSummary`), generates insight text via `insights.TATInsightsGenerator`, and then produces:
- Charts (`visualizer.TATVisualizer`) and slides (`ppt_report.ProfessionalPPTGenerator`), used only by the PPTX output.
- The Excel workbook (`excel_report.ExcelReportGenerator`), which has no charts at all.

These two outputs are allowed to diverge in structure: the PPTX slides are an aggregate/presentation-oriented view (bucket-distribution charts, compact tables) built with "own logic," while the Excel workbook mirrors a specific reference dashboard's grid layout. Don't assume a change to one implies the same change to the other.

### Excel workbook is formula-driven against a fixed Raw Data layout (`excel_report.py`)

The workbook always has exactly three sheets: **Recal Summary**, **Flat Summary**, **Raw Data** — each summary sheet is fully self-contained (title block, TAT table, pending-TAT, monthly ageing, ageing contribution, status ageing, WIP-monthly, severity), not split across additional tabs.

Critically, the summary sheets contain **live Excel formulas** (`COUNTIFS`/`SUMIFS`/`IFERROR` division), not pre-computed Python values — they reference the Raw Data sheet directly by column letter via the `RAW_DATA_COLUMNS` → `RAW` letter map, so a cell like `=COUNTIFS('Raw Data'!$X:$X,"IT",...)` recalculates in Excel itself. When editing these formulas:
- `RAW_DATA_COLUMNS` defines a fixed 24-column order for the Raw Data sheet; every summary formula's column-letter reference depends on that order staying stable. Changing it requires updating every `SummarySheetBuilder` method that hardcodes a `RAW[...]` lookup.
- Formulas intentionally do **not** filter by client name (that filter criterion was removed from every `COUNTIFS`/`SUMIFS` call in both summary sheets) — a workbook always represents one already-scoped pull, so an extra client-name criterion was redundant and, in the "All Clients" combined case, actively wrong.
- `_write_header_block` writes the client name to row 1 and the title (a `CONCATENATE` formula referencing row 1) to row 3 — the row numbers are load-bearing; other code assumes this exact placement.
- Formula correctness cannot be verified from the formula strings alone — Excel doesn't evaluate `openpyxl`-written formulas until the file is opened. Verifying a change means opening the generated workbook in real Excel (e.g. via `win32com.client`), forcing a full recalculation, and reading back computed cell values.

### Streamlit session-state persistence for downloads (`app.py`)

`st.download_button` triggers a Streamlit rerun on click. Report outputs (chart bytes, dataframes, PPTX/XLSX bytes) are read into `st.session_state` immediately after generation — before the `tempfile.TemporaryDirectory()` they were written into goes out of scope — and the results-rendering section reads only from `st.session_state`, gated on `st.session_state.report_ready`/`report_mode`, independent of whether a generate button was just clicked. This is what allows both the PPTX and XLSX download buttons to coexist without either click wiping out the other's data.

### Environment caveat: WSL/UNC paths

This repo lives under a `\\wsl.localhost\...` UNC path. Streamlit's file-watcher does not reliably detect changes there, so after editing any `.py` file the running `streamlit run` process must be fully killed and restarted (not relied upon to hot-reload) before changes take effect.
