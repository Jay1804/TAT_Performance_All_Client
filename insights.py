"""Templated (non-AI) performance summary text, ported unchanged from
PPT_Report_Generator_S3/generate1.py TATInsightsGenerator.

Note: the reference project imported the OpenAI SDK but never actually called
it anywhere - all "insights" text there was already generated purely from
computed percentages via f-strings, exactly as below. No AI dependency
existed to remove from this logic.
"""
import logging
from typing import Any, Dict

import pandas as pd

from config import Config

logger = logging.getLogger(__name__)


class TATInsightsGenerator:
    @staticmethod
    def generate_insights(df_final: pd.DataFrame, df_recal: pd.DataFrame, pivot_df: pd.DataFrame, client_name: str = "") -> Dict[str, Any]:
        try:
            final_total = df_final[df_final["Year"] == "Grand Total"] if not df_final.empty else pd.DataFrame()
            recal_total = df_recal[df_recal["Year"] == "Grand Total"] if not df_recal.empty else pd.DataFrame()

            final_in_pct = float(final_total["IN TAT %"].iloc[0].replace("%", "")) if len(final_total) > 0 else 0
            recal_in_pct = float(recal_total["IN TAT %"].iloc[0].replace("%", "")) if len(recal_total) > 0 else 0

            closed_by_client = completed = insufficient = wip = total = 0
            if not pivot_df.empty and len(pivot_df) > 1:
                grand_total_row = pivot_df[pivot_df["Year"] == "Grand Total"]
                if len(grand_total_row) > 0:
                    closed_by_client = int(grand_total_row["Closed by Client"].iloc[0]) if "Closed by Client" in grand_total_row.columns else 0
                    completed = int(grand_total_row["Completed"].iloc[0]) if "Completed" in grand_total_row.columns else 0
                    insufficient = int(grand_total_row["Insufficient"].iloc[0]) if "Insufficient" in grand_total_row.columns else 0
                    wip = int(grand_total_row["Work in Progress"].iloc[0]) if "Work in Progress" in grand_total_row.columns else 0
                    total = int(grand_total_row["Grand Total"].iloc[0]) if "Grand Total" in grand_total_row.columns else 0

            completed_pct = (completed / total * 100) if total > 0 else 0
            wip_pct = (wip / total * 100) if total > 0 else 0

            if final_in_pct >= Config.TAT_TARGET:
                performance_status = "✅ Meeting Target"
            elif final_in_pct >= Config.TAT_TARGET - 10:
                performance_status = "⚠️ Close to Target"
            else:
                performance_status = "❌ Below Target"

            exec_summary_bullets = [
                f"📊 Overall TAT Performance for {client_name}: Final Status achieved {final_in_pct:.1f}% IN TAT",
                f"🔄 Recal Status Performance: {recal_in_pct:.1f}% IN TAT",
                f"✅ Case Completion Rate: {completed_pct:.1f}% of cases completed",
                f"⏳ Work in Progress (WIP): {wip} cases ({wip_pct:.1f}%) require attention",
                f"🎯 Performance Status: {performance_status} (Target: {Config.TAT_TARGET}%)",
            ]

            tat_insights = [
                f"• Final Status IN TAT: {final_in_pct:.1f}% vs Target {Config.TAT_TARGET}%",
                f"• Recal Status IN TAT: {recal_in_pct:.1f}%",
                f"• Gap to Target: {max(0, Config.TAT_TARGET - final_in_pct):.1f}%",
            ]

            return {
                "exec_summary_bullets": exec_summary_bullets,
                "tat_insights": tat_insights,
                "final_in_tat_pct": final_in_pct,
                "recal_in_tat_pct": recal_in_pct,
                "performance_status": performance_status,
            }
        except Exception as e:
            logger.error("Error generating insights: %s", e)
            return {
                "exec_summary_bullets": ["TAT Performance analysis completed"],
                "tat_insights": ["Monitor IN TAT percentages"],
                "final_in_tat_pct": 0,
                "recal_in_tat_pct": 0,
                "performance_status": "Unknown",
            }
