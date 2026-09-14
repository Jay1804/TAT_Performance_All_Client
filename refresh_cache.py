"""Manual cache-warming script: hits the database once for the "All Clients
Combined" scope and overwrites the shared on-disk cache (data_cache.py) that
app.py reads from.

Run this whenever you want fresh data available to every user, instead of
letting each Streamlit user trigger their own database pull:

    python refresh_cache.py
    python refresh_cache.py --months-back 6

After this completes, anyone using the app's "Load Data" button for the
All Clients Combined scope + the same months-back window gets served
straight from this cache - no database hit on their end.

The pull is done in small per-client batches (see
data_loading.fetch_full_dataset_batched) rather than one unscoped query -
the live production database has been observed to drop the connection
mid-query on the single-query "all clients" pull (2-2.5M rows), most likely
a server-side execution-time guard.
"""
import argparse
import time

import data_cache
import db_connect
from data_loading import fetch_client_list, fetch_full_dataset_batched


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--months-back", type=int, default=4,
        help="Months of received_date history to pull (default: 4, matching the app's default).",
    )
    parser.add_argument(
        "--batch-size", type=int, default=200,
        help="Clients per batch query (default: 200). Lower this if batches still time out.",
    )
    args = parser.parse_args()

    conn = db_connect.get_connection()
    try:
        client_df = fetch_client_list(conn)
    finally:
        conn.close()
    company_ids = client_df["company_id"].tolist()

    print(f"Pulling {len(company_ids)} clients, last {args.months_back} month(s), "
          f"in batches of {args.batch_size}...")
    start = time.time()

    def progress(done, total, n_rows):
        print(f"  batch {done}/{total}: {n_rows:,} rows")

    df = fetch_full_dataset_batched(
        db_connect.get_connection, company_ids,
        months_back=args.months_back, batch_size=args.batch_size, progress=progress,
    )

    elapsed = time.time() - start
    cache_key = data_cache.make_key(None, args.months_back)
    data_cache.save(df, cache_key)

    print(f"Done in {elapsed:.1f}s - cached {len(df):,} rows under key {cache_key}.")


if __name__ == "__main__":
    main()
