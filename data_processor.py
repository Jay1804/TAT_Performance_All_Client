"""Data preparation and TAT/case-status/severity aggregation logic.

Adapted from the TATCalculator / CaseStatusAnalyzer / SeverityAnalyzer classes in
C:\\Gen AI\\PPT_Report_Generator_S3\\generate1.py. The original versions had to
auto-detect column names and date formats because they accepted arbitrary
uploaded CSVs; here the schema is fixed (it comes straight from
sql_queries.QUERY_BASE_DATA), so that detection layer is dropped and the
aggregation logic itself (grouping, TAT % math, pivot construction) is kept
as-is.
"""
import logging
from dataclasses import dataclass
from datetime import date
from typing import List, Optional, Tuple

import pandas as pd

logger = logging.getLogger(__name__)


@dataclass
class TATMetrics:
    year: int
    month: str
    month_num: int
    cases_received: int
    completed_cases: int
    in_tat: int
    out_of_tat: int
    in_tat_pct: float
    out_of_tat_pct: float


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Lower-snake-case all column names, e.g. 'Report Severity' -> 'report_severity'."""
    df = df.copy()
    df.columns = [str(col).strip().lower().replace(" ", "_") for col in df.columns]
    return df


def prepare_dataframe(df: pd.DataFrame, date_column: str = "received_date"):
    """Add received_date_parsed/year/month_num/month_name derived columns.

    Returns (df_with_dates, min_date, max_date). Rows with an unparseable date
    are dropped (mirrors TATCalculator.extract_month_from_date's behaviour).
    """
    df = df.copy()
    df["received_date_parsed"] = pd.to_datetime(df[date_column], errors="coerce")
    df = df[df["received_date_parsed"].notna()].copy()

    if df.empty:
        return df, None, None

    df["year"] = df["received_date_parsed"].dt.year
    df["month_num"] = df["received_date_parsed"].dt.month
    df["month_name"] = df["received_date_parsed"].dt.strftime("%b") + "'" + df["received_date_parsed"].dt.strftime("%y")

    min_date = df["received_date_parsed"].min().date()
    max_date = df["received_date_parsed"].max().date()
    return df, min_date, max_date


class TATCalculator:
    @staticmethod
    def filter_by_date_range(df: pd.DataFrame, start_date: date, end_date: date) -> pd.DataFrame:
        if "received_date_parsed" not in df.columns:
            raise ValueError("received_date_parsed column missing - call prepare_dataframe() first")

        mask = (df["received_date_parsed"].dt.date >= start_date) & (df["received_date_parsed"].dt.date <= end_date)
        return df[mask].copy()

    @staticmethod
    def calculate_tat_metrics(df: pd.DataFrame, tat_column: str) -> List[TATMetrics]:
        results: List[TATMetrics] = []

        if tat_column not in df.columns:
            logger.warning("TAT column '%s' not found", tat_column)
            return results

        grouped = df.groupby(["year", "month_num", "month_name"])

        for (year, month_num, month_name), group in grouped:
            if pd.isna(year):
                continue

            total_cases = len(group)
            completed_cases = total_cases

            tat_values = group[tat_column].astype(str).str.upper()
            in_tat = len(tat_values[tat_values.str.contains("IT|IN|WITHIN", na=False)])
            out_of_tat = len(tat_values[tat_values.str.contains("OT|OUT|BREACH", na=False)])

            total_tat_cases = in_tat + out_of_tat
            if total_tat_cases > 0:
                in_tat_pct = (in_tat / total_tat_cases) * 100
                out_of_tat_pct = (out_of_tat / total_tat_cases) * 100
            else:
                in_tat_pct = out_of_tat_pct = 0

            results.append(TATMetrics(
                year=int(year),
                month=month_name,
                month_num=int(month_num),
                cases_received=total_cases,
                completed_cases=completed_cases,
                in_tat=in_tat,
                out_of_tat=out_of_tat,
                in_tat_pct=in_tat_pct,
                out_of_tat_pct=out_of_tat_pct,
            ))

        results.sort(key=lambda x: (x.year, x.month_num))
        return results

    @staticmethod
    def create_tat_dataframe(metrics_list: List[TATMetrics]) -> pd.DataFrame:
        if not metrics_list:
            return pd.DataFrame()

        data = []
        for m in metrics_list:
            data.append({
                "Year": m.year,
                "Month": m.month,
                "Cases Received": m.cases_received,
                "Completed Cases": m.completed_cases,
                "IN TAT": m.in_tat,
                "Out of TAT": m.out_of_tat,
                "IN TAT %": f"{m.in_tat_pct:.2f}%",
                "Out of TAT %": f"{m.out_of_tat_pct:.2f}%",
            })

        df = pd.DataFrame(data)

        total_received = df["Cases Received"].sum()
        total_in_tat = df["IN TAT"].sum()
        total_out_tat = df["Out of TAT"].sum()

        total_tat_cases = total_in_tat + total_out_tat
        total_in_pct = (total_in_tat / total_tat_cases * 100) if total_tat_cases > 0 else 0

        grand_total = pd.DataFrame([{
            "Year": "Grand Total",
            "Month": "",
            "Cases Received": total_received,
            "Completed Cases": total_received,
            "IN TAT": total_in_tat,
            "Out of TAT": total_out_tat,
            "IN TAT %": f"{total_in_pct:.2f}%",
            "Out of TAT %": f"{100 - total_in_pct:.2f}%",
        }])

        return pd.concat([df, grand_total], ignore_index=True)


def _monthly_pivot(df: pd.DataFrame, value_column: str, case_id_column: str) -> Tuple[pd.DataFrame, List[str]]:
    """Shared pivot builder used by both case-status and severity breakdowns."""
    if "received_date_parsed" not in df.columns:
        raise ValueError("received_date_parsed column missing - call prepare_dataframe() first")

    df_with_year = df.copy()
    df_with_year["month_display"] = df_with_year["received_date_parsed"].dt.strftime("%b'%y")
    df_with_year["year"] = df_with_year["received_date_parsed"].dt.year
    df_with_year["month_num"] = df_with_year["received_date_parsed"].dt.month

    pivot_table = pd.pivot_table(
        df_with_year,
        values=case_id_column,
        index=["year", "month_num", "month_display"],
        columns=value_column,
        aggfunc="count",
        fill_value=0,
    )

    result_df = pivot_table.reset_index()
    result_df = result_df.sort_values(["year", "month_num"]).reset_index(drop=True)

    value_columns = [col for col in result_df.columns if col not in ["year", "month_num", "month_display"]]
    result_df["Grand Total"] = result_df[value_columns].sum(axis=1) if value_columns else 0
    result_df["Month"] = result_df["month_display"]
    result_df["Year"] = result_df["year"]

    final_columns = ["Year", "Month"] + value_columns + ["Grand Total"]
    result_df = result_df[final_columns]
    return result_df, value_columns


class CaseStatusAnalyzer:
    @staticmethod
    def create_case_status_pivot(df: pd.DataFrame, case_status_column: str, case_id_column: str = "case_ars_no") -> pd.DataFrame:
        """Monthly pivot of case counts per status value found in the data.

        The upstream tool hardcoded 4 expected status labels ('Closed by
        Client', 'Completed', 'Insufficient', 'Work in Progress'); our SQL
        query can also surface 'On Hold', so the status columns are derived
        dynamically from whatever values are actually present instead.
        """
        result_df, status_columns = _monthly_pivot(df, case_status_column, case_id_column)

        if result_df.empty:
            return result_df

        column_totals = {"Year": "Grand Total", "Month": ""}
        for col in status_columns:
            column_totals[col] = result_df[col].sum()
        column_totals["Grand Total"] = result_df["Grand Total"].sum()
        grand_total_row = pd.DataFrame([column_totals])
        return pd.concat([result_df, grand_total_row], ignore_index=True)

    @staticmethod
    def combined_position_table(pivot_df: pd.DataFrame, tat_df: pd.DataFrame) -> pd.DataFrame:
        """Merge the monthly case-status pivot with a monthly TAT dataframe
        (df_recal or df_final) into one table - mirrors the Excel Recal/Flat
        Summary tabs' combined 'Over All Cases Position + Sent Cases TAT
        Analysis' block. Both inputs share the same Year/Month keys
        (including a Year='Grand Total' row), so a left-merge lines up
        row-for-row without extra bookkeeping.
        """
        if pivot_df.empty or tat_df.empty:
            return pd.DataFrame()

        left = pivot_df.rename(columns={"Grand Total": "Cases Received"})
        right = tat_df[["Year", "Month", "IN TAT", "Out of TAT", "IN TAT %", "Out of TAT %"]]
        return pd.merge(left, right, on=["Year", "Month"], how="left")


class SeverityAnalyzer:
    @staticmethod
    def create_severity_pivot(df: pd.DataFrame, severity_column: str, case_id_column: str = "case_ars_no") -> pd.DataFrame:
        result_df, severity_columns = _monthly_pivot(df, severity_column, case_id_column)

        if result_df.empty:
            return result_df

        column_totals = {"Year": "Grand Total", "Month": ""}
        for col in severity_columns:
            column_totals[col] = result_df[col].sum()
        column_totals["Grand Total"] = result_df["Grand Total"].sum()
        grand_total_row = pd.DataFrame([column_totals])

        overall_total = column_totals["Grand Total"]
        pct_row = {"Year": "Percentage", "Month": ""}
        for col in severity_columns:
            pct_row[col] = f"{(column_totals[col] / overall_total * 100) if overall_total else 0:.2f}%"
        pct_row["Grand Total"] = "100.00%" if overall_total else "0.00%"
        percentage_row = pd.DataFrame([pct_row])

        return pd.concat([result_df, grand_total_row, percentage_row], ignore_index=True)


AGEING_BUCKET_ORDER = ["0--2", "3--10", "11--14", "15--20", "21--25", "26--30", "30 +"]
PENDING_STATUSES = ["Work in Progress", "Insufficient", "On Hold"]


class AgeingSummary:
    """Aggregate (not monthly) Ageing breakdowns - the PPTX deck is a
    presentation medium, so it gets bucket-distribution charts and compact
    bi-furcation tables rather than Excel's month-by-month grids."""

    @staticmethod
    def bucket_counts(df: pd.DataFrame, bucket_column: str) -> pd.Series:
        if bucket_column not in df.columns:
            return pd.Series(dtype="int64")
        counts = df[bucket_column].value_counts()
        ordered = [b for b in AGEING_BUCKET_ORDER if b in counts.index]
        ordered += [b for b in counts.index if b not in AGEING_BUCKET_ORDER]
        return counts.reindex(ordered).fillna(0).astype(int)

    @staticmethod
    def status_ageing_table(
        df: pd.DataFrame, statuses: List[str], case_status_column: str,
        bucket_column: str, case_id_column: str = "case_ars_no",
        include_contribution_pct: bool = True,
    ) -> pd.DataFrame:
        """Status x Ageing-bucket table (e.g. WIP-only, or Insufficient +
        On Hold combined), with a Grand Total row - the reference
        dashboard's 'Bi-furcation and Ageing' tables. Optionally appends a
        'Contribution %' row (each bucket's Grand Total share of Total
        Cases), mirroring the reference's Ageing Contribution % block."""
        filtered = df[df[case_status_column].isin(statuses)]
        if filtered.empty or bucket_column not in filtered.columns:
            return pd.DataFrame()

        pivot = pd.pivot_table(
            filtered, values=case_id_column, index=case_status_column,
            columns=bucket_column, aggfunc="count", fill_value=0,
        )
        bucket_cols = list(pivot.columns)
        ordered_buckets = [b for b in AGEING_BUCKET_ORDER if b in bucket_cols]
        ordered_buckets += [b for b in bucket_cols if b not in AGEING_BUCKET_ORDER]

        pivot = pivot.reindex(index=statuses, columns=ordered_buckets, fill_value=0).reset_index()
        pivot = pivot.rename(columns={case_status_column: "Status"})
        pivot["Total Cases"] = pivot[ordered_buckets].sum(axis=1)
        pivot = pivot[["Status", "Total Cases"] + ordered_buckets]

        grand_total = {"Status": "Grand Total", "Total Cases": pivot["Total Cases"].sum()}
        for col in ordered_buckets:
            grand_total[col] = pivot[col].sum()
        result = pd.concat([pivot, pd.DataFrame([grand_total])], ignore_index=True)

        if include_contribution_pct:
            total = grand_total["Total Cases"]
            pct_row = {"Status": "Contribution %", "Total Cases": "100.00%" if total else "0.00%"}
            for col in ordered_buckets:
                pct_row[col] = f"{(grand_total[col] / total * 100) if total else 0:.2f}%"
            result = pd.concat([result, pd.DataFrame([pct_row])], ignore_index=True)

        return result

    @staticmethod
    def monthly_bucket_table(df: pd.DataFrame, bucket_column: str, case_id_column: str = "case_ars_no") -> pd.DataFrame:
        """Year/Month x Ageing-bucket monthly counts + Total Cases, with a
        Grand Total row and a Contribution % row - mirrors the reference
        dashboard's 'Complete Cases Ageing' + 'Ageing Bucket Contribution %'
        sections, combined into a single table."""
        if bucket_column not in df.columns:
            return pd.DataFrame()

        result_df, bucket_cols = _monthly_pivot(df, bucket_column, case_id_column)
        if result_df.empty:
            return result_df

        ordered_buckets = [b for b in AGEING_BUCKET_ORDER if b in bucket_cols]
        ordered_buckets += [b for b in bucket_cols if b not in AGEING_BUCKET_ORDER]

        result_df = result_df[["Year", "Month"] + ordered_buckets + ["Grand Total"]]
        result_df = result_df.rename(columns={"Grand Total": "Total Cases"})

        column_totals = {"Year": "Grand Total", "Month": ""}
        for col in ordered_buckets:
            column_totals[col] = result_df[col].sum()
        column_totals["Total Cases"] = result_df["Total Cases"].sum()
        grand_total_row = pd.DataFrame([column_totals])

        total = column_totals["Total Cases"]
        pct_row = {"Year": "Contribution %", "Month": "", "Total Cases": "100.00%" if total else "0.00%"}
        for col in ordered_buckets:
            pct_row[col] = f"{(column_totals[col] / total * 100) if total else 0:.2f}%"
        percentage_row = pd.DataFrame([pct_row])

        return pd.concat([result_df, grand_total_row, percentage_row], ignore_index=True)


class PendingSummary:
    @staticmethod
    def pending_tat_table(
        df: pd.DataFrame, tat_column: str, case_status_column: str = "case_status",
        statuses: Optional[List[str]] = None,
    ) -> pd.DataFrame:
        """Case Status x (Total/IN TAT/Out TAT/percentages) for not-yet-closed
        statuses - the reference dashboard's 'Pending Cases TAT %' table."""
        statuses = statuses or PENDING_STATUSES
        rows = []
        for status in statuses:
            sub = df[df[case_status_column] == status]
            total = len(sub)
            tat_values = sub[tat_column].astype(str).str.upper() if tat_column in sub.columns else pd.Series(dtype=str)
            in_tat = int((tat_values == "IT").sum())
            out_tat = int((tat_values == "OT").sum())
            denom = in_tat + out_tat
            rows.append({
                "Case Status": status,
                "Total Cases": total,
                "IN TAT": in_tat,
                "Out Of TAT": out_tat,
                "IN TAT%": f"{(in_tat / denom * 100) if denom else 0:.2f}%",
                "OUT Of TAT%": f"{(out_tat / denom * 100) if denom else 0:.2f}%",
            })

        result_df = pd.DataFrame(rows)
        total = result_df["Total Cases"].sum()
        in_tat = result_df["IN TAT"].sum()
        out_tat = result_df["Out Of TAT"].sum()
        denom = in_tat + out_tat
        grand_total = {
            "Case Status": "Grand Total", "Total Cases": total, "IN TAT": in_tat, "Out Of TAT": out_tat,
            "IN TAT%": f"{(in_tat / denom * 100) if denom else 0:.2f}%",
            "OUT Of TAT%": f"{(out_tat / denom * 100) if denom else 0:.2f}%",
        }
        return pd.concat([result_df, pd.DataFrame([grand_total])], ignore_index=True)

