"""Casewise TAT Report Generator - Streamlit app.

Fetches case-wise TAT data from the checkpoint_live database (via
db_connect.py + sql_queries.QUERY_BASE_DATA) and generates the same
PPTX/XLSX TAT performance report previously produced by
C:\\Gen AI\\PPT_Report_Generator_S3, adapted to read from this database
instead of an uploaded CSV/S3 file. No external AI/API keys are used -
all narrative text is computed from the data itself.

Run with: streamlit run app.py
"""
import logging
import tempfile
import zipfile
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

import data_cache
import db_connect
from config import COL_CLIENT_EXTERNAL_ID, COL_COMPANY, COL_PROCESS
from data_loading import fetch_client_list, fetch_full_dataset
from data_processor import prepare_dataframe
from report_builder import generate_client_report

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

st.set_page_config(page_title="Casewise TAT Report Generator", layout="wide")
st.title("📊 Casewise TAT Report Generator")
st.caption("Generates the TAT performance PPTX/XLSX report directly from the checkpoint_live database.")
st.markdown("---")


# --------------------------------------------------------------------------
# Data loading
# --------------------------------------------------------------------------
# "All clients, several months" is genuinely 2-2.5 MILLION matching rows on
# this database (confirmed via EXPLAIN), so the app asks for a scope
# (specific client(s), or an explicit "all clients" opt-in) BEFORE hitting
# the big case tables at all. The client picker itself only queries small
# dimension tables (ec_client/ec_master_company - a few thousand rows), so
# it stays fast regardless of scope. Every scoped pull is cached to disk
# (data_cache.py, keyed by scope) and reused across app restarts.

CLIENT_LIST_TTL_SECONDS = 24 * 60 * 60  # client roster changes rarely
CACHE_TTL_SECONDS = data_cache.DEFAULT_TTL_SECONDS
MAX_MONTHS_BACK = 6


def safe_filename(name: str) -> str:
    return "".join(c if c.isalnum() else "_" for c in name)


for key, default in [
    ("client_list_df", None),
    ("client_list_error", None),
    ("df_raw", None),
    ("df_prepared", None),
    ("min_date", None),
    ("max_date", None),
    ("load_error", None),
    ("data_source", None),  # "cache" or "database"
    ("loaded_scope_label", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


def _apply_loaded_data(df_raw: pd.DataFrame) -> None:
    if df_raw.empty:
        st.session_state.load_error = "Query executed successfully but returned no rows for this scope."
        st.session_state.df_raw = None
        st.session_state.df_prepared = None
        return

    df_prepared, min_date, max_date = prepare_dataframe(df_raw)
    if df_prepared.empty:
        st.session_state.load_error = (
            f"Loaded {len(df_raw):,} rows, but none had a parseable value in the "
            "received_date column."
        )
        st.session_state.df_raw = None
        st.session_state.df_prepared = None
        return

    st.session_state.df_raw = df_raw
    st.session_state.df_prepared = df_prepared
    st.session_state.min_date = min_date
    st.session_state.max_date = max_date


st.markdown("### Step 1: Select Scope")

if st.session_state.client_list_df is None:
    with st.spinner("Loading client list..."):
        try:
            conn = db_connect.get_connection()
            try:
                st.session_state.client_list_df = fetch_client_list(conn)
            finally:
                conn.close()
        except db_connect.ConfigError as e:
            st.session_state.client_list_error = (
                f"Database configuration error: {e}\n\n"
                "Check that all required variables are set in the .env file "
                "(DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_NAME)."
            )
        except Exception as e:
            st.session_state.client_list_error = f"Failed to load client list: {e}"

if st.session_state.client_list_error:
    st.error(st.session_state.client_list_error)
    st.stop()

client_list_df = st.session_state.client_list_df
label_to_company_id = {}
for _, r in client_list_df.iterrows():
    ext_id = r.get("client_external_id")
    label = f"{r['company_name']} (External ID: {int(ext_id)})" if pd.notna(ext_id) else str(r["company_name"])
    label_to_company_id[label] = r["company_id"]

col_scope1, col_scope2 = st.columns(2)

with col_scope1:
    fetch_mode = st.radio(
        "Pull scope",
        ["Specific Client(s)", "All Clients Combined (large pull)"],
        horizontal=True,
        help="Scoping to specific clients keeps the database pull fast (seconds). "
             "All Clients Combined pulls millions of rows and can take a long time.",
    )
    selected_fetch_labels = []
    all_clients_confirmed = False
    if fetch_mode == "Specific Client(s)":
        selected_fetch_labels = st.multiselect(
            "Client(s) — search by name or external ID",
            sorted(label_to_company_id.keys()),
        )
    else:
        st.warning(
            "This pulls data for every client at once - confirmed to be roughly "
            "2-2.5 million rows for a 4-month window on this database. It can take "
            "several minutes or longer, depending on server load."
        )
        all_clients_confirmed = st.checkbox("I understand - fetch all clients anyway")

with col_scope2:
    months_back = st.number_input(
        "Months of history (received_date)",
        min_value=1, max_value=MAX_MONTHS_BACK, value=4, step=1,
    )

selected_company_ids = [label_to_company_id[label] for label in selected_fetch_labels]
scope_ready = bool(selected_company_ids) or (fetch_mode.startswith("All") and all_clients_confirmed)

col_load1, col_load2, col_status = st.columns([1, 1, 2])
with col_load1:
    load_cached_clicked = st.button("📂 Load Data (use cache)", type="primary", disabled=not scope_ready)
with col_load2:
    force_refresh_clicked = st.button("⚡ Force Refresh from Database", disabled=not scope_ready)

if not scope_ready:
    st.info("Select at least one client (or confirm the all-clients pull) to enable loading.")

if load_cached_clicked or force_refresh_clicked:
    cache_key = data_cache.make_key(selected_company_ids or None, months_back)
    scope_label = (
        f"{len(selected_company_ids)} client(s)" if selected_company_ids else "All Clients Combined"
    ) + f", last {months_back} month(s)"

    st.session_state.load_error = None
    # "Load Data" always serves whatever is cached for this scope, however old -
    # ordinary users should never trigger a database pull just because the TTL
    # expired. Only a truly missing cache (first-time scope) or an explicit
    # "Force Refresh" click goes to the database. Refresh the shared cache with
    # refresh_cache.py (or the Force Refresh button) instead.
    cached_df = None if force_refresh_clicked else data_cache.load(cache_key)

    if cached_df is not None:
        _apply_loaded_data(cached_df)
        st.session_state.data_source = "cache"
        st.session_state.loaded_scope_label = scope_label
    else:
        spinner_msg = (
            "Connecting to database and running query..."
            if selected_company_ids
            else "Connecting to database and running query across ALL clients - this can take a while..."
        )
        with st.spinner(spinner_msg):
            try:
                conn = db_connect.get_connection()
                try:
                    df_raw = fetch_full_dataset(conn, company_ids=selected_company_ids or None, months_back=months_back)
                finally:
                    conn.close()
            except db_connect.ConfigError as e:
                st.session_state.load_error = (
                    f"Database configuration error: {e}\n\n"
                    "Check that all required variables are set in the .env file "
                    "(DB_HOST, DB_PORT, DB_USER, DB_PASSWORD, DB_NAME)."
                )
                df_raw = None
            except Exception as e:
                st.session_state.load_error = f"Failed to load data from the database: {e}"
                df_raw = None

            if df_raw is not None:
                _apply_loaded_data(df_raw)
                if st.session_state.df_prepared is not None:
                    data_cache.save(df_raw, cache_key)
                    st.session_state.data_source = "database"
                    st.session_state.loaded_scope_label = scope_label

with col_status:
    if st.session_state.load_error:
        st.error(st.session_state.load_error)
    elif st.session_state.df_prepared is not None:
        df = st.session_state.df_prepared
        cache_age = data_cache.age_seconds(data_cache.make_key(selected_company_ids or None, months_back))
        source_note = (
            f"from cache ({int(cache_age // 60)} min old)"
            if st.session_state.data_source == "cache" and cache_age is not None
            else "freshly fetched from database"
        )
        st.success(
            f"✅ Loaded {len(df):,} rows ({source_note}) | Scope: {st.session_state.loaded_scope_label} | "
            f"{df[COL_COMPANY].nunique() if COL_COMPANY in df.columns else '?'} clients in result | "
            f"Date range: {st.session_state.min_date} to {st.session_state.max_date}"
        )
        if (
            st.session_state.data_source == "cache"
            and cache_age is not None
            and cache_age >= CACHE_TTL_SECONDS
        ):
            st.warning(
                f"This cache is {int(cache_age // 3600)}h{int((cache_age % 3600) // 60)}m old. "
                "Click **Force Refresh from Database**, or run `python refresh_cache.py` to "
                "update the shared cache for everyone."
            )
    else:
        st.info("Select a scope above, then click **Load Data**.")

if st.session_state.df_prepared is None:
    st.stop()

df = st.session_state.df_prepared

with st.expander("Preview raw data (first 20 rows)"):
    st.dataframe(df.head(20), use_container_width=True)

st.markdown("---")

# --------------------------------------------------------------------------
# Filters
# --------------------------------------------------------------------------

st.markdown("### Step 2: Report Options")
st.caption("These operate on the data already loaded above - no additional database query.")

col_dates, col_scope = st.columns(2)

with col_dates:
    date_range = st.date_input(
        "Date range (based on received_date)",
        value=(st.session_state.min_date, st.session_state.max_date),
        min_value=st.session_state.min_date,
        max_value=st.session_state.max_date,
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = end_date = date_range

with col_scope:
    label_to_client = {}
    if COL_COMPANY in df.columns:
        if COL_CLIENT_EXTERNAL_ID in df.columns:
            client_lookup = (
                df[[COL_COMPANY, COL_CLIENT_EXTERNAL_ID]]
                .dropna(subset=[COL_COMPANY])
                .drop_duplicates(subset=[COL_COMPANY])
            )
            for _, r in client_lookup.iterrows():
                ext_id = r[COL_CLIENT_EXTERNAL_ID]
                label = f"{r[COL_COMPANY]} (External ID: {int(ext_id)})" if pd.notna(ext_id) else str(r[COL_COMPANY])
                label_to_client[label] = r[COL_COMPANY]
        else:
            for name in df[COL_COMPANY].dropna().astype(str).unique():
                label_to_client[name] = name

    client_labels = sorted(label_to_client.keys())

    scope_mode = st.radio(
        "Report grouping",
        ["All Loaded Clients (Combined)", "Select Specific Client(s)"],
        horizontal=True,
        help="Choose how to group the already-loaded data into report(s) - "
             "one combined report, or one report per selected client.",
    )

    selected_clients = []
    if scope_mode == "Select Specific Client(s)":
        selected_labels = st.multiselect(
            "Client(s) — search by name or external ID",
            client_labels,
            help="Type either the client name or its external ID to filter this list.",
        )
        selected_clients = [label_to_client[label] for label in selected_labels]

col_opts1, col_opts2, col_opts3 = st.columns(3)
with col_opts1:
    generate_ppt = st.checkbox("Generate PowerPoint (.pptx)", value=True)
with col_opts2:
    generate_excel = st.checkbox("Generate Excel workbook (.xlsx)", value=True)
with col_opts3:
    if COL_PROCESS in df.columns:
        process_options = sorted(df[COL_PROCESS].dropna().astype(str).unique().tolist())
        process_filter = st.multiselect("Filter by process (optional)", process_options)
    else:
        process_filter = []

st.markdown("---")
st.markdown("### Step 3: Generate Report")

if not generate_ppt and not generate_excel:
    st.warning("Select at least one output format (PPTX or XLSX) above.")

generate_clicked = st.button("🚀 Generate Report", type="primary", disabled=not (generate_ppt or generate_excel))

# Results are stashed in session_state (as bytes, not file paths - the
# TemporaryDirectory they're generated in is deleted at the end of this
# script run) so that clicking one download_button - which triggers its own
# Streamlit rerun - doesn't wipe out the results section and make the OTHER
# download button disappear. Without this, only whichever file the user
# downloads first is ever retrievable.
for key, default in [("report_ready", False), ("report_mode", None)]:
    if key not in st.session_state:
        st.session_state[key] = default

if generate_clicked:
    if scope_mode == "Select Specific Client(s)" and not selected_clients:
        st.error("Please select at least one client, or switch to 'All Clients (Combined)'.")
        st.stop()

    df_scope = df.copy()
    if process_filter:
        df_scope = df_scope[df_scope[COL_PROCESS].astype(str).isin(process_filter)]

    if df_scope.empty:
        st.error("No rows match the selected process filter.")
        st.stop()

    st.session_state.report_ready = False

    with tempfile.TemporaryDirectory() as tmpdir:
        if scope_mode == "All Loaded Clients (Combined)" or len(selected_clients) <= 1:
            if selected_clients:
                client_name = selected_clients[0]
            else:
                # "Combined" scope with no explicit pick - if the loaded data
                # happens to resolve to exactly one real company, name the
                # report after that company instead of a generic label.
                distinct_companies = df_scope[COL_COMPANY].dropna().unique() if COL_COMPANY in df_scope.columns else []
                client_name = distinct_companies[0] if len(distinct_companies) == 1 else "All Clients"
            client_df = df_scope if not selected_clients else df_scope[df_scope[COL_COMPANY] == client_name]

            with st.spinner(f"Generating report for {client_name}..."):
                result = generate_client_report(
                    client_df, client_name, tmpdir, start_date, end_date,
                    generate_ppt=generate_ppt, generate_excel=generate_excel,
                )

            if not result["success"]:
                st.error(f"Report generation failed: {result['error']}")
            else:
                chart_images = {}
                for chart_key, chart_path in result["chart_paths"].items():
                    if chart_path:
                        with open(chart_path, "rb") as f:
                            chart_images[chart_key] = f.read()

                st.session_state.report_mode = "single"
                st.session_state.report_client_name = client_name
                st.session_state.report_insights = result["insights"] or {}
                st.session_state.report_df_final = result["df_final"]
                st.session_state.report_df_recal = result["df_recal"]
                st.session_state.report_pivot_df = result["pivot_df"]
                st.session_state.report_severity_pivot_df = result["severity_pivot_df"]
                st.session_state.report_chart_images = chart_images
                st.session_state.report_ppt_bytes = (
                    open(result["ppt_path"], "rb").read() if generate_ppt and result["ppt_path"] else None
                )
                st.session_state.report_excel_bytes = (
                    open(result["excel_path"], "rb").read() if generate_excel and result["excel_path"] else None
                )
                st.session_state.report_ready = True

        else:
            # Batch mode: multiple clients selected -> generate each, zip the outputs.
            progress_bar = st.progress(0)
            status_text = st.empty()
            results = []

            for idx, client_name in enumerate(selected_clients):
                status_text.text(f"Processing: {client_name}")
                client_df = df_scope[df_scope[COL_COMPANY] == client_name]

                if client_df.empty:
                    results.append({"success": False, "client": client_name, "error": "No data found for this client"})
                else:
                    result = generate_client_report(
                        client_df, client_name, tmpdir, start_date, end_date,
                        generate_ppt=generate_ppt, generate_excel=generate_excel,
                    )
                    results.append(result)

                progress_bar.progress((idx + 1) / len(selected_clients))

            status_text.text("Processing complete!")

            successful = [r for r in results if r.get("success")]
            failed = [r for r in results if not r.get("success")]

            all_report_paths = []
            for result in successful:
                if generate_ppt and result.get("ppt_path"):
                    all_report_paths.append(result["ppt_path"])
                if generate_excel and result.get("excel_path"):
                    all_report_paths.append(result["excel_path"])

            zip_bytes = None
            if all_report_paths:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zipf:
                    for path in all_report_paths:
                        zipf.write(path, arcname=path.split("/")[-1].split("\\")[-1])
                zip_bytes = zip_buffer.getvalue()

            st.session_state.report_mode = "batch"
            st.session_state.report_successful_count = len(successful)
            st.session_state.report_failed = [
                {"client": f["client"], "error": f.get("error", "Unknown error")} for f in failed
            ]
            st.session_state.report_zip_bytes = zip_bytes
            st.session_state.report_ready = True

if st.session_state.report_ready and st.session_state.report_mode == "single":
    client_name = st.session_state.report_client_name
    insights = st.session_state.report_insights

    st.success(f"Report generated for **{client_name}**")

    m1, m2, m3 = st.columns(3)
    m1.metric("Final IN TAT %", f"{insights.get('final_in_tat_pct', 0):.1f}%")
    m2.metric("Recal IN TAT %", f"{insights.get('recal_in_tat_pct', 0):.1f}%")
    m3.metric("Status", insights.get("performance_status", "N/A"))

    for bullet in insights.get("exec_summary_bullets", []):
        st.markdown(f"- {bullet}")

    tabs = st.tabs(["TAT Final", "TAT Recal", "Case Status", "Severity", "Charts"])
    with tabs[0]:
        st.dataframe(st.session_state.report_df_final, use_container_width=True)
    with tabs[1]:
        st.dataframe(st.session_state.report_df_recal, use_container_width=True)
    with tabs[2]:
        st.dataframe(st.session_state.report_pivot_df, use_container_width=True)
    with tabs[3]:
        st.dataframe(st.session_state.report_severity_pivot_df, use_container_width=True)
    with tabs[4]:
        for chart_key, chart_bytes in st.session_state.report_chart_images.items():
            st.image(chart_bytes, caption=chart_key, use_container_width=True)

    dl_col1, dl_col2 = st.columns(2)
    if st.session_state.report_ppt_bytes is not None:
        with dl_col1:
            st.download_button(
                "📥 Download PPTX",
                st.session_state.report_ppt_bytes,
                file_name=f"TAT_Report_{safe_filename(client_name)}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                key="dl_pptx",
            )
    if st.session_state.report_excel_bytes is not None:
        with dl_col2:
            st.download_button(
                "📥 Download XLSX",
                st.session_state.report_excel_bytes,
                file_name=f"TAT_Dashboard_{safe_filename(client_name)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_xlsx",
            )

elif st.session_state.report_ready and st.session_state.report_mode == "batch":
    c1, c2 = st.columns(2)
    c1.metric("✅ Successful", st.session_state.report_successful_count)
    c2.metric("❌ Failed", len(st.session_state.report_failed))

    if st.session_state.report_failed:
        st.error("Failed clients:")
        for fail in st.session_state.report_failed:
            st.write(f"- {fail['client']}: {fail['error']}")

    if st.session_state.report_zip_bytes is not None:
        st.download_button(
            "📥 Download All Reports (ZIP)",
            st.session_state.report_zip_bytes,
            file_name=f"TAT_Reports_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
            mime="application/zip",
            key="dl_zip",
        )
