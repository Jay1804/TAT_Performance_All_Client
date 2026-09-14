"""Single entry point for pulling + enriching the Casewise TAT dataset.

Runs the (now holiday-subquery-free) QUERY_BASE_DATA, fetches the small
holidays table once, and applies the vectorized TAT/Ageing calculations in
Python (see tat_calculations.py) instead of the removed per-row correlated
SQL subqueries.

"All clients, several months" is genuinely a multi-million-row pull on this
database (confirmed via EXPLAIN - MySQL has to examine ~2-2.5M matching
rows in ec_case_master alone before any joins), so callers should scope to
specific client(s) via company_ids wherever possible; company_ids=None runs
the full unscoped pull and should be treated as an explicit, slow/batch
choice, not a default.
"""
from typing import Callable, Iterable, List, Optional

import pandas as pd

import sql_queries
from data_processor import normalize_columns
from tat_calculations import apply_tat_calculations, load_holidays


def fetch_client_list(conn) -> pd.DataFrame:
    """Fast lookup of actual reporting clients (company_id, company_name,
    client_external_id) - queries only the small dimension tables, never
    touches the large case tables.
    """
    with conn.cursor() as cursor:
        cursor.execute(sql_queries.QUERY_CLIENT_LIST)
        rows = cursor.fetchall()
    return pd.DataFrame(rows)


def fetch_full_dataset(conn, company_ids: Optional[Iterable[int]] = None, months_back: int = 4) -> pd.DataFrame:
    query = sql_queries.build_base_data_query(company_ids=company_ids, months_back=months_back)

    with conn.cursor() as cursor:
        cursor.execute(query)
        rows = cursor.fetchall()

    df = normalize_columns(pd.DataFrame(rows))

    if df.empty:
        return df

    with conn.cursor() as cursor:
        holidays = load_holidays(cursor)

    return apply_tat_calculations(df, holidays)


def fetch_full_dataset_batched(
    get_conn: Callable[[], object],
    company_ids: List[int],
    months_back: int = 4,
    batch_size: int = 200,
    progress: Optional[Callable[[int, int, int], None]] = None,
) -> pd.DataFrame:
    """Same result as fetch_full_dataset(company_ids=None, ...), but built by
    looping over company_ids in small scoped batches instead of running one
    unscoped multi-million-row query.

    The live production database has been observed to drop the connection
    mid-query ("Lost connection to MySQL server during query") on the
    unscoped all-clients pull, most likely a server-side execution-time
    guard protecting the OLTP workload. Each batch reuses the same
    company_id-scoped query path that's already fast for a handful of
    clients, and opens a fresh connection per batch via get_conn (a callable
    like db_connect.get_connection, not a live connection - a single
    connection held across many sequential queries risks the same guard).
    """
    frames = []
    batches = [company_ids[i:i + batch_size] for i in range(0, len(company_ids), batch_size)]

    for i, batch in enumerate(batches):
        query = sql_queries.build_base_data_query(company_ids=batch, months_back=months_back)
        conn = get_conn()
        try:
            with conn.cursor() as cursor:
                cursor.execute(query)
                rows = cursor.fetchall()
        finally:
            conn.close()
        frames.append(normalize_columns(pd.DataFrame(rows)))
        if progress:
            progress(i + 1, len(batches), len(rows))

    df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    if df.empty:
        return df

    conn = get_conn()
    try:
        with conn.cursor() as cursor:
            holidays = load_holidays(cursor)
    finally:
        conn.close()

    return apply_tat_calculations(df, holidays)
