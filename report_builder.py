"""Orchestrates one report (charts + PPTX + XLSX) for a filtered DataFrame.

Adapted from IntegratedClientReportGenerator.generate_complete_report in
PPT_Report_Generator_S3/generate1.py: same pipeline (TAT metrics -> case
status/severity pivots -> insights -> charts -> slides -> workbook), but
driven off the fixed column names in report.config instead of runtime
column auto-detection, and with the Outlook-email step removed entirely.
"""
import logging
import os
from datetime import date
from typing import Any, Dict

import pandas as pd

from config import COL_CASE_ID, COL_CASE_STATUS, COL_SEVERITY, COL_TAT_FINAL, COL_TAT_RECAL
from data_processor import (
    AgeingSummary, CaseStatusAnalyzer, PendingSummary, SeverityAnalyzer, TATCalculator,
)
from excel_report import ExcelReportGenerator
from insights import TATInsightsGenerator
from ppt_report import ProfessionalPPTGenerator
from visualizer import TATVisualizer

logger = logging.getLogger(__name__)

WIP_STATUS = "Work in Progress"
NON_WORKABLE_STATUSES = ["Insufficient", "On Hold"]


def generate_client_report(
    df: pd.DataFrame,
    client_name: str,
    temp_dir: str,
    start_date: date,
    end_date: date,
    generate_ppt: bool = True,
    generate_excel: bool = True,
) -> Dict[str, Any]:
    """Build the TAT report for an already-scoped DataFrame (single client or combined)."""
    result: Dict[str, Any] = {
        "success": False,
        "client": client_name,
        "ppt_path": None,
        "excel_path": None,
        "insights": None,
        "df_final": pd.DataFrame(),
        "df_recal": pd.DataFrame(),
        "pivot_df": pd.DataFrame(),
        "severity_pivot_df": pd.DataFrame(),
        "chart_paths": {},
        "error": None,
    }

    try:
        client_temp_dir = os.path.join(temp_dir, "".join(c if c.isalnum() else "_" for c in client_name))
        os.makedirs(client_temp_dir, exist_ok=True)

        df_filtered = TATCalculator.filter_by_date_range(df, start_date, end_date)

        if df_filtered.empty:
            result["error"] = f"No data found for date range {start_date} to {end_date}"
            return result

        raw_filtered_data = df_filtered.copy()

        df_final = pd.DataFrame()
        if COL_TAT_FINAL in df_filtered.columns:
            final_metrics = TATCalculator.calculate_tat_metrics(df_filtered, COL_TAT_FINAL)
            df_final = TATCalculator.create_tat_dataframe(final_metrics)

        df_recal = pd.DataFrame()
        if COL_TAT_RECAL in df_filtered.columns:
            recal_metrics = TATCalculator.calculate_tat_metrics(df_filtered, COL_TAT_RECAL)
            df_recal = TATCalculator.create_tat_dataframe(recal_metrics)

        pivot_df = pd.DataFrame()
        if COL_CASE_STATUS in df_filtered.columns:
            pivot_df = CaseStatusAnalyzer.create_case_status_pivot(df_filtered, COL_CASE_STATUS, COL_CASE_ID)

        severity_pivot_df = pd.DataFrame()
        if COL_SEVERITY in df_filtered.columns:
            severity_pivot_df = SeverityAnalyzer.create_severity_pivot(df_filtered, COL_SEVERITY, COL_CASE_ID)

        insights = TATInsightsGenerator.generate_insights(df_final, df_recal, pivot_df, client_name)
        result["insights"] = insights
        result["df_final"] = df_final
        result["df_recal"] = df_recal
        result["pivot_df"] = pivot_df
        result["severity_pivot_df"] = severity_pivot_df

        # Aggregates mirroring the Excel Recal/Flat Summary tabs' sections,
        # but at a presentation-appropriate grain (aggregate distributions +
        # compact bi-furcation tables, not Excel's month-by-month grids).
        wip_df = df_filtered[df_filtered[COL_CASE_STATUS] == WIP_STATUS] if COL_CASE_STATUS in df_filtered.columns else df_filtered.iloc[0:0]
        non_workable_df = df_filtered[df_filtered[COL_CASE_STATUS].isin(NON_WORKABLE_STATUSES)] if COL_CASE_STATUS in df_filtered.columns else df_filtered.iloc[0:0]

        ageing_recal = AgeingSummary.bucket_counts(df_filtered, "ageing_bucket")
        ageing_final = AgeingSummary.bucket_counts(df_filtered, "ageing_bucket_final")
        wip_ageing_recal = AgeingSummary.bucket_counts(wip_df, "ageing_bucket")
        wip_ageing_final = AgeingSummary.bucket_counts(wip_df, "ageing_bucket_final")
        non_workable_ageing_recal = AgeingSummary.status_ageing_table(df_filtered, NON_WORKABLE_STATUSES, COL_CASE_STATUS, "ageing_bucket", COL_CASE_ID)
        non_workable_ageing_final = AgeingSummary.status_ageing_table(df_filtered, NON_WORKABLE_STATUSES, COL_CASE_STATUS, "ageing_bucket_final", COL_CASE_ID)

        pending_recal = PendingSummary.pending_tat_table(df_filtered, COL_TAT_RECAL, COL_CASE_STATUS)
        pending_final = PendingSummary.pending_tat_table(df_filtered, COL_TAT_FINAL, COL_CASE_STATUS)
        pending_combined = pd.DataFrame()
        if not pending_recal.empty and not pending_final.empty:
            pending_combined = pd.DataFrame({
                "Case Status": pending_recal["Case Status"],
                "Total Cases": pending_recal["Total Cases"],
                "Recal IN TAT%": pending_recal["IN TAT%"],
                "Recal OUT TAT%": pending_recal["OUT Of TAT%"],
                "Final IN TAT%": pending_final["IN TAT%"],
                "Final OUT TAT%": pending_final["OUT Of TAT%"],
            })

        wip_tat_final = TATCalculator.create_tat_dataframe(TATCalculator.calculate_tat_metrics(wip_df, COL_TAT_FINAL)) if not wip_df.empty else pd.DataFrame()
        wip_tat_recal = TATCalculator.create_tat_dataframe(TATCalculator.calculate_tat_metrics(wip_df, COL_TAT_RECAL)) if not wip_df.empty else pd.DataFrame()

        # Charts are only used by the PPTX deck - the Excel workbook is
        # tables/formulas only, no embedded images.
        chart_paths: Dict[str, str] = {}
        if generate_ppt:
            viz = TATVisualizer(client_temp_dir)

            if not df_final.empty and not df_recal.empty:
                chart_paths["trend"] = viz.create_tat_trend_chart(df_final, df_recal, "TAT Performance Trend", client_name)
                chart_paths["monthly"] = viz.create_monthly_comparison(df_final, df_recal, client_name)

            if not pivot_df.empty:
                chart_paths["case_trend"] = viz.create_case_status_trend(pivot_df, client_name)

            if ageing_recal.sum() > 0 or ageing_final.sum() > 0:
                chart_paths["ageing"] = viz.create_bucket_comparison_chart(ageing_recal, ageing_final, "Ageing Bucket Distribution", client_name)

            if wip_ageing_recal.sum() > 0 or wip_ageing_final.sum() > 0:
                chart_paths["wip_ageing"] = viz.create_bucket_comparison_chart(wip_ageing_recal, wip_ageing_final, "Work In Progress Ageing", client_name)

            if not wip_tat_final.empty and not wip_tat_recal.empty:
                chart_paths["wip_monthly"] = viz.create_monthly_comparison(wip_tat_final, wip_tat_recal, f"{client_name} - Work In Progress")

            if not severity_pivot_df.empty:
                chart_paths["severity"] = viz.create_severity_chart(severity_pivot_df, client_name)

            result["chart_paths"] = chart_paths

        date_range_str = f"{start_date.strftime('%b %Y')} - {end_date.strftime('%b %Y')}"
        metrics = {
            "total_cases": len(df_filtered),
            "final_in_tat": f"{insights.get('final_in_tat_pct', 0):.1f}%",
            "recal_in_tat": f"{insights.get('recal_in_tat_pct', 0):.1f}%",
        }
        safe_name = "".join(c if c.isalnum() else "_" for c in client_name)

        if generate_ppt:
            ppt_gen = ProfessionalPPTGenerator(client_temp_dir, client_name)

            ppt_gen.add_professional_title_slide("TAT Performance Dashboard", "Turnaround Time Analysis & Case Status Report", metrics, date_range_str)
            ppt_gen.add_executive_summary_slide(insights["exec_summary_bullets"])

            if chart_paths.get("trend"):
                trend_insights = "\n".join([
                    "📈 Performance Analysis:",
                    f"• TAT performance compared against {95.0}% target",
                    f"• Final Status: {insights['tat_insights'][0] if insights['tat_insights'] else 'N/A'}",
                    "• Stacked bars show proportion of IN vs OUT of TAT",
                    "• Red dashed line indicates organizational target",
                ])
                ppt_gen.add_chart_slide("TAT Performance Comparison", chart_paths["trend"], trend_insights)

            if not df_final.empty:
                ppt_gen.add_table_slide("TAT Status - Final / Flat", df_final, "IN TAT% = IT/(IT+OT) x 100")

            if not df_recal.empty:
                ppt_gen.add_table_slide("TAT Status - Recal", df_recal, "IN TAT% = IT/(IT+OT) x 100")

            if chart_paths.get("monthly"):
                monthly_insights_text = "\n".join([
                    "📊 Trend Analysis:",
                    "• Month-over-month performance tracking",
                    "• Dashed lines show statistical trends",
                    "• Green line indicates target",
                ])
                ppt_gen.add_chart_slide("Monthly TAT Trends - Recal vs Final", chart_paths["monthly"], monthly_insights_text)

            if not pivot_df.empty:
                ppt_gen.add_table_slide(
                    "Case Status Distribution", pivot_df,
                    "Monthly case counts by status", highlight_last_col=True,
                )

            if chart_paths.get("case_trend"):
                case_trend_insights = "\n".join([
                    "📋 Case Volume Analysis:",
                    "• Left chart: Individual status trends over time",
                    "• Right chart: Stacked area shows volume composition",
                ])
                ppt_gen.add_chart_slide("Case Status Trends", chart_paths["case_trend"], case_trend_insights)

            if not pending_combined.empty:
                ppt_gen.add_table_slide(
                    "Pending Cases - TAT % (Recal vs Final)", pending_combined,
                    "Work in Progress, Insufficient and On Hold cases - not yet closed",
                )

            if chart_paths.get("ageing"):
                ppt_gen.add_chart_slide(
                    "Ageing Bucket Distribution", chart_paths["ageing"],
                    "📊 Ageing Analysis:\n• All received cases, bucketed by days elapsed\n"
                    "• Recal vs Final comparison shows how re-calculation shifts ageing\n"
                    "• A '30 +' bucket that dominates signals a backlog risk",
                )

            if chart_paths.get("wip_ageing"):
                ppt_gen.add_chart_slide(
                    "Work In Progress - Ageing", chart_paths["wip_ageing"],
                    "📋 WIP Deep-Dive:\n• Ageing distribution for currently open (Work in Progress) cases only\n"
                    "• Highlights which open cases are approaching or past TAT",
                )

            if chart_paths.get("wip_monthly"):
                ppt_gen.add_chart_slide(
                    "Work In Progress - Monthly TAT Trend", chart_paths["wip_monthly"],
                    "📈 WIP Trend:\n• Monthly IN-TAT % for currently open cases, Recal vs Final\n"
                    "• Tracks whether the open pipeline is improving or slipping over time",
                )

            if not non_workable_ageing_recal.empty:
                ppt_gen.add_table_slide(
                    "Non-Workable Cases - Ageing (Recal)", non_workable_ageing_recal,
                    "Insufficient and On Hold cases, bucketed by Recal ageing",
                )

            if not non_workable_ageing_final.empty:
                ppt_gen.add_table_slide(
                    "Non-Workable Cases - Ageing (Final)", non_workable_ageing_final,
                    "Insufficient and On Hold cases, bucketed by Final ageing",
                )

            if not severity_pivot_df.empty:
                ppt_gen.add_table_slide(
                    "Report Severity Distribution", severity_pivot_df,
                    "Monthly report counts by severity", highlight_last_col=True,
                )

            if chart_paths.get("severity"):
                ppt_gen.add_chart_slide(
                    "Report Severity Distribution", chart_paths["severity"],
                    "🎯 Severity Mix:\n• Share of sent reports by severity category\n"
                    "• A rising Discrepant/Red share warrants closer review",
                )

            ppt_gen.add_thank_you_slide()

            ppt_filename = f"TAT_Report_{safe_name}.pptx"
            result["ppt_path"] = ppt_gen.save(ppt_filename)

        if generate_excel:
            excel_filename = f"TAT_Dashboard_{safe_name}.xlsx"
            excel_path = os.path.join(client_temp_dir, excel_filename)

            ExcelReportGenerator.create_excel_report(
                df_filtered=raw_filtered_data,
                client_name=client_name,
                date_range_str=date_range_str,
                insights=insights,
                output_path=excel_path,
            )
            result["excel_path"] = excel_path

        result["success"] = True
        return result

    except Exception as e:
        result["error"] = str(e)
        logger.error("Error generating report for %s: %s", client_name, e)
        return result
