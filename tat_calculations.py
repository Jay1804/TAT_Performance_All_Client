"""Vectorized replacement for the holiday-dependent SQL layers removed from
QUERY_BASE_DATA (due_date_Recal, Ageing, Ageing_Bucket, TAT Status_Recal,
due_date_Final, opt_Final_due_date, Ageing_Final, Ageing_Bucket_Final,
TAT Status_Final, Insuff_Ageing, Insuff_Ageing_Bucket).

Why this exists: the original SQL ran ~10 correlated subqueries against the
holidays table per output row. MySQL re-evaluates each one per row ("Range
checked for each record") instead of once per query, which made even a
single-client, 10-day pull take minutes. The holidays table itself is tiny
(~3k rows), so counting holidays in a date range is done here with
np.searchsorted against a small sorted array per "off-day" category -
O(log n) per row instead of a per-row SQL round trip.

Off-day sets (which holiday rows count as "non-working" for due-date
extension and Ageing exclusion) mirror the ones the original SQL already
used correctly for due_date_Recal/due_date_Final, applied consistently to
Ageing too:
  - Working / workday            -> holiday_type in (1, 2)
  - Calendar / calendar          -> holiday_type == 2
  - WorkingPlusSaturday / workday-sat -> holiday_type == 2 OR weekday is Sunday

ec_master_holidays has no holiday_type == 3 (the value the original,
Redshift-era query - written against a differently-shaped calendar table -
filtered on for 'Working' Ageing), so "Ageing = elapsed calendar days minus
off-days" is used instead of the original literal "count of type=3 rows",
which would otherwise always be zero on this schema. Confirmed with the
report owner.
"""
from datetime import date, datetime
from typing import Optional

import numpy as np
import pandas as pd

WORKING_TYPES = {"Working", "workday"}
CALENDAR_TYPES = {"Calendar", "calendar"}
SATURDAY_TYPES = {"WorkingPlusSaturday", "workday-sat"}

AGEING_BUCKET_TIERS = [
    (3, 10, "3--10"),
    (11, 14, "11--14"),
    (15, 20, "15--20"),
    (21, 25, "21--25"),
    (26, 30, "26--30"),
]


def load_holidays(cursor) -> pd.DataFrame:
    """Fetch the full (small) holidays table. `cursor` is an open DB cursor."""
    cursor.execute("SELECT holiday_date, holiday_type FROM ec_master_holidays")
    rows = cursor.fetchall()
    df = pd.DataFrame(rows)
    df["holiday_date"] = pd.to_datetime(df["holiday_date"])
    df["is_sunday"] = df["holiday_date"].dt.dayofweek == 6
    return df


def _sorted_dates(holidays: pd.DataFrame, mask: pd.Series) -> np.ndarray:
    return np.sort(holidays.loc[mask, "holiday_date"].values.astype("datetime64[D]"))


class OffDaySets:
    """Precomputed sorted date arrays for each TAT-type's off-day definition."""

    def __init__(self, holidays: pd.DataFrame):
        is_sunday = holidays.get("is_sunday")
        if is_sunday is None:
            is_sunday = holidays["holiday_date"].dt.dayofweek == 6

        self.working = _sorted_dates(holidays, holidays["holiday_type"].isin([1, 2]))
        self.calendar = _sorted_dates(holidays, holidays["holiday_type"] == 2)
        self.saturday = _sorted_dates(holidays, (holidays["holiday_type"] == 2) | is_sunday)

    def for_tat_type(self, tat_type: str) -> np.ndarray:
        if tat_type in WORKING_TYPES:
            return self.working
        if tat_type in CALENDAR_TYPES:
            return self.calendar
        if tat_type in SATURDAY_TYPES:
            return self.saturday
        return np.array([], dtype="datetime64[D]")


def _count_in_range(sorted_dates: np.ndarray, starts: np.ndarray, ends: np.ndarray) -> np.ndarray:
    """For each row, count how many values in sorted_dates fall in [starts[i], ends[i]] inclusive."""
    lo = np.searchsorted(sorted_dates, starts, side="left")
    hi = np.searchsorted(sorted_dates, ends, side="right")
    return (hi - lo).astype("int64")


def _count_offdays_per_row(df: pd.DataFrame, offsets: OffDaySets, start_col: str, end_col: str) -> np.ndarray:
    """Count off-days per row, grouped by tat_type since each type uses a different date set."""
    starts = df[start_col].values.astype("datetime64[D]")
    ends = df[end_col].values.astype("datetime64[D]")
    counts = np.zeros(len(df), dtype="int64")

    for tat_type in df["client_tat_type"].dropna().unique():
        sorted_dates = offsets.for_tat_type(tat_type)
        mask = (df["client_tat_type"] == tat_type).values
        if not mask.any():
            continue
        if len(sorted_dates) == 0:
            counts[mask] = 0
            continue
        counts[mask] = _count_in_range(sorted_dates, starts[mask], ends[mask])

    return counts


def _bucket(values: pd.Series, first_tier_low: int = -3) -> pd.Series:
    """Mirrors the Ageing_Bucket/Ageing_Bucket_Final/Insuff_Ageing_Bucket CASE expressions.
    Insuff_Ageing_Bucket uses -50 as the first tier's lower bound in the original SQL;
    Ageing_Bucket/Ageing_Bucket_Final use -3 - hence the parameter.
    """
    result = pd.Series(pd.NA, index=values.index, dtype="object")
    result = result.mask(values.between(first_tier_low, 2), "0--2")
    for low, high, label in AGEING_BUCKET_TIERS:
        result = result.mask(values.between(low, high), label)
    result = result.mask(values > 30, "30 +")
    return result


def _due_date(df: pd.DataFrame, offsets: OffDaySets, base_col: str) -> pd.Series:
    """due_date = base_date + Process_TAT + count(off-days in [base_date, base_date + Process_TAT])."""
    base = df[base_col]
    naive_due = base + pd.to_timedelta(df["process_tat"], unit="D")

    tmp = pd.DataFrame({
        "client_tat_type": df["client_tat_type"],
        "_start": base,
        "_end": naive_due,
    })
    offday_count = _count_offdays_per_row(tmp, offsets, "_start", "_end")

    return base + pd.to_timedelta(df["process_tat"] + offday_count, unit="D")


def _ageing(df: pd.DataFrame, offsets: OffDaySets, base_col: str, today: pd.Timestamp) -> pd.Series:
    """Ageing = elapsed calendar days (base_date -> Final_Report_Sent or today) minus off-days in that range."""
    end = df["final_report_sent"].fillna(today)
    base = df[base_col]
    elapsed = (end - base).dt.days

    tmp = pd.DataFrame({
        "client_tat_type": df["client_tat_type"],
        "_start": base,
        "_end": end,
    })
    offday_count = _count_offdays_per_row(tmp, offsets, "_start", "_end")

    return elapsed - offday_count


def _tat_status(client_category: pd.Series, insuff_status: pd.Series, case_status: pd.Series,
                 ageing: pd.Series, process_tat: pd.Series) -> pd.Series:
    result = pd.Series("OT", index=ageing.index, dtype="object")
    result = result.mask(ageing <= process_tat, "IT")
    result = result.mask(case_status.isin(["Closed by Authbridge", "Closed by Client", "On Hold"]), "IT")
    result = result.mask((client_category == "RED") & (insuff_status == "Insuff"), "IT")
    return result


def apply_tat_calculations(df: pd.DataFrame, holidays: pd.DataFrame, today: Optional[date] = None) -> pd.DataFrame:
    """Add due_date_Recal, Ageing, Ageing_Bucket, TAT Status_Recal, due_date_Final,
    opt_Final_due_date, Ageing_Final, Ageing_Bucket_Final, TAT Status_Final,
    Insuff_Ageing, Insuff_Ageing_Bucket to a normalized (lower_snake_case) DataFrame.

    Expects columns produced by the simplified QUERY_BASE_DATA (through T2):
    start_date, received_date, process_tat, client_tat_type, final_report_sent,
    client_category, insuff_status, case_status, latest_check_insuff_raised_date.
    """
    df = df.copy()
    today_ts = pd.Timestamp(today or date.today())

    for col in ["start_date", "received_date", "final_report_sent", "latest_check_insuff_raised_date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col])

    offsets = OffDaySets(holidays)

    # --- Recal branch (keyed on start_date) ---
    df["due_date_recal"] = _due_date(df, offsets, "start_date")
    df["ageing"] = _ageing(df, offsets, "start_date", today_ts)
    df["ageing_bucket"] = _bucket(df["ageing"])
    df["tat_status_recal"] = _tat_status(
        df["client_category"], df["insuff_status"], df["case_status"], df["ageing"], df["process_tat"]
    )

    # --- Final branch (keyed on received_date) ---
    df["due_date_final"] = _due_date(df, offsets, "received_date")

    dow = df["due_date_final"].dt.dayofweek  # Monday=0 ... Sunday=6
    opt_due = df["due_date_final"].copy()
    is_working = df["client_tat_type"].isin(WORKING_TYPES)
    is_saturday_type = df["client_tat_type"].isin(SATURDAY_TYPES)
    opt_due = opt_due.mask(is_working & (dow == 5), df["due_date_final"] + pd.Timedelta(days=2))  # Saturday
    opt_due = opt_due.mask((is_working | is_saturday_type) & (dow == 6), df["due_date_final"] + pd.Timedelta(days=1))  # Sunday
    df["opt_final_due_date"] = opt_due

    # Ageing_Final always used the same (unbranched) off-day rule as the 'Working' type in the original SQL.
    df["ageing_final"] = (
        df["final_report_sent"].fillna(today_ts) - df["received_date"]
    ).dt.days - _count_offdays_per_row(
        pd.DataFrame({
            "client_tat_type": "Working",
            "_start": df["received_date"],
            "_end": df["final_report_sent"].fillna(today_ts),
        }),
        offsets, "_start", "_end",
    )
    df["ageing_bucket_final"] = _bucket(df["ageing_final"])
    df["tat_status_final"] = _tat_status(
        df["client_category"], df["insuff_status"], df["case_status"], df["ageing_final"], df["process_tat"]
    )

    # --- Insuff_Ageing: business days (Mon-Fri) elapsed, no holiday-table involvement in the original either ---
    has_insuff_date = df["latest_check_insuff_raised_date"].notna()
    insuff_ageing = pd.Series(pd.NA, index=df.index, dtype="Int64")
    if has_insuff_date.any():
        starts = df.loc[has_insuff_date, "latest_check_insuff_raised_date"].dt.date.values.astype("datetime64[D]")
        ends = np.full(starts.shape, today_ts.date(), dtype="datetime64[D]")
        insuff_ageing.loc[has_insuff_date] = np.busday_count(starts, ends + np.timedelta64(1, "D"))
    df["insuff_ageing"] = insuff_ageing
    df["insuff_ageing_bucket"] = _bucket(df["insuff_ageing"].astype("float64"), first_tier_low=-50)

    return df
