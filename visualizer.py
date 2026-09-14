"""Chart generation, ported unchanged from PPT_Report_Generator_S3/generate1.py TATVisualizer."""
import logging
import os

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns

from config import ColorScheme, Config

logger = logging.getLogger(__name__)


class TATVisualizer:
    def __init__(self, temp_dir: str):
        self.temp_dir = temp_dir
        self.colors = ColorScheme()
        plt.rcParams["font.size"] = 10
        plt.rcParams["axes.titlesize"] = 12
        sns.set_style("whitegrid")

    def create_tat_trend_chart(self, df_final: pd.DataFrame, df_recal: pd.DataFrame, title: str, client_name: str = "") -> str:
        try:
            fig, axes = plt.subplots(1, 2, figsize=(14, 5.5))

            df_final_chart = df_final[df_final["Year"] != "Grand Total"].copy() if not df_final.empty else pd.DataFrame()
            df_recal_chart = df_recal[df_recal["Year"] != "Grand Total"].copy() if not df_recal.empty else pd.DataFrame()

            if not df_final_chart.empty:
                months = df_final_chart["Month"].tolist()
                in_tat = [float(v.replace("%", "")) for v in df_final_chart["IN TAT %"]]
                out_tat = [float(v.replace("%", "")) for v in df_final_chart["Out of TAT %"]]

                x = range(len(months))
                axes[0].bar(x, in_tat, label="IN TAT", color=self.colors.accent_normalized, alpha=0.8)
                axes[0].bar(x, out_tat, bottom=in_tat, label="Out of TAT", color=self.colors.warning_normalized, alpha=0.8)
                axes[0].set_xlabel("Month", fontsize=11)
                axes[0].set_ylabel("Percentage (%)", fontsize=11)
                axes[0].set_title("TAT Status Final", fontweight="bold", fontsize=12)
                axes[0].set_xticks(x)
                axes[0].set_xticklabels(months, rotation=45, ha="right")
                axes[0].legend(loc="upper right")
                axes[0].grid(True, alpha=0.3)
                axes[0].set_ylim(0, 100)
                axes[0].axhline(y=Config.TAT_TARGET, color="red", linestyle="--", alpha=0.7, linewidth=1.5, label=f"Target {Config.TAT_TARGET}%")

            if not df_recal_chart.empty:
                months = df_recal_chart["Month"].tolist()
                in_tat = [float(v.replace("%", "")) for v in df_recal_chart["IN TAT %"]]
                out_tat = [float(v.replace("%", "")) for v in df_recal_chart["Out of TAT %"]]

                x = range(len(months))
                axes[1].bar(x, in_tat, label="IN TAT", color=self.colors.accent_normalized, alpha=0.8)
                axes[1].bar(x, out_tat, bottom=in_tat, label="Out of TAT", color=self.colors.warning_normalized, alpha=0.8)
                axes[1].set_xlabel("Month", fontsize=11)
                axes[1].set_ylabel("Percentage (%)", fontsize=11)
                axes[1].set_title("TAT Status Recal", fontweight="bold", fontsize=12)
                axes[1].set_xticks(x)
                axes[1].set_xticklabels(months, rotation=45, ha="right")
                axes[1].legend(loc="upper right")
                axes[1].grid(True, alpha=0.3)
                axes[1].set_ylim(0, 100)
                axes[1].axhline(y=Config.TAT_TARGET, color="red", linestyle="--", alpha=0.7, linewidth=1.5, label=f"Target {Config.TAT_TARGET}%")

            full_title = f"{title} - {client_name}" if client_name else title
            plt.suptitle(full_title, fontsize=14, fontweight="bold")
            plt.tight_layout()

            path = os.path.join(self.temp_dir, f"tat_trend_comparison_{client_name.replace(' ', '_')}.png")
            plt.savefig(path, dpi=Config.CHART_DPI, bbox_inches="tight", facecolor="white")
            plt.close()
            return path
        except Exception as e:
            logger.error("Error creating trend chart: %s", e)
            return None

    def create_performance_gauge(self, performance_pct: float, title: str, client_name: str = "") -> str:
        try:
            fig, ax = plt.subplots(figsize=(8, 6))

            if performance_pct >= Config.TAT_TARGET:
                color = self.colors.accent_normalized
                status = "Excellent"
            elif performance_pct >= Config.TAT_TARGET - 10:
                color = self.colors.secondary_normalized
                status = "Good"
            else:
                color = self.colors.warning_normalized
                status = "Needs Improvement"

            ax.barh([0], [performance_pct], color=color, height=0.3)
            ax.barh([0], [100], color="#D3D3D3", height=0.3, alpha=0.3)

            ax.set_xlim(0, 100)
            ax.set_ylim(-0.5, 0.5)
            ax.set_xlabel("IN TAT Percentage (%)", fontsize=12)
            full_title = f"{title} - {client_name}" if client_name else title
            ax.set_title(f"{full_title} - {status}", fontsize=14, fontweight="bold")
            ax.text(performance_pct, 0, f"{performance_pct:.1f}%\n{status}",
                    ha="center", va="center", fontsize=16, fontweight="bold", color=color)
            ax.set_yticks([])
            ax.grid(True, axis="x", alpha=0.3)

            plt.tight_layout()
            path = os.path.join(self.temp_dir, f"{title.lower().replace(' ', '_')}_{client_name.replace(' ', '_')}.png")
            plt.savefig(path, dpi=Config.CHART_DPI, bbox_inches="tight", facecolor="white")
            plt.close()
            return path
        except Exception as e:
            logger.error("Error creating gauge chart: %s", e)
            return None

    def create_monthly_comparison(self, df_final: pd.DataFrame, df_recal: pd.DataFrame, client_name: str = "") -> str:
        try:
            fig, ax = plt.subplots(figsize=(12, 5.5))

            df_final_chart = df_final[df_final["Year"] != "Grand Total"].copy() if not df_final.empty else pd.DataFrame()
            df_recal_chart = df_recal[df_recal["Year"] != "Grand Total"].copy() if not df_recal.empty else pd.DataFrame()

            if not df_final_chart.empty and not df_recal_chart.empty:
                months = df_final_chart["Month"].tolist()
                final_in_pct = [float(v.replace("%", "")) for v in df_final_chart["IN TAT %"]]
                recal_in_pct = [float(v.replace("%", "")) for v in df_recal_chart["IN TAT %"]]

                x = range(len(months))
                ax.plot(x, final_in_pct, marker="o", linewidth=2, markersize=8,
                        label="TAT Status Final", color=self.colors.primary_normalized)
                ax.plot(x, recal_in_pct, marker="s", linewidth=2, markersize=8,
                        label="TAT Status Recal", color=self.colors.secondary_normalized)

                if len(x) > 1:
                    z_final = np.polyfit(x, final_in_pct, 1)
                    p_final = np.poly1d(z_final)
                    ax.plot(x, p_final(x), "--", linewidth=1, alpha=0.6, color=self.colors.primary_normalized, label="Final Trend")

                    z_recal = np.polyfit(x, recal_in_pct, 1)
                    p_recal = np.poly1d(z_recal)
                    ax.plot(x, p_recal(x), "--", linewidth=1, alpha=0.6, color=self.colors.secondary_normalized, label="Recal Trend")

                ax.set_xlabel("Month", fontsize=12)
                ax.set_ylabel("IN TAT Percentage (%)", fontsize=12)
                full_title = f"Monthly IN TAT Performance Comparison - {client_name}" if client_name else "Monthly IN TAT Performance Comparison"
                ax.set_title(full_title, fontsize=14, fontweight="bold")
                ax.set_xticks(x)
                ax.set_xticklabels(months, rotation=45, ha="right")
                ax.legend(loc="best")
                ax.grid(True, alpha=0.3)
                ax.axhline(y=Config.TAT_TARGET, color="green", linestyle="--", alpha=0.7, linewidth=2, label=f"Target ({Config.TAT_TARGET}%)")
                ax.set_ylim(0, 100)

                best_final_idx = int(np.argmax(final_in_pct))
                worst_final_idx = int(np.argmin(final_in_pct))
                ax.annotate(f"Best: {final_in_pct[best_final_idx]:.1f}%",
                            (best_final_idx, final_in_pct[best_final_idx]),
                            xytext=(5, 10), textcoords="offset points", fontsize=9, fontweight="bold")
                ax.annotate(f"Worst: {final_in_pct[worst_final_idx]:.1f}%",
                            (worst_final_idx, final_in_pct[worst_final_idx]),
                            xytext=(5, -15), textcoords="offset points", fontsize=9)

            plt.tight_layout()
            path = os.path.join(self.temp_dir, f"monthly_comparison_{client_name.replace(' ', '_')}.png")
            plt.savefig(path, dpi=Config.CHART_DPI, bbox_inches="tight", facecolor="white")
            plt.close()
            return path
        except Exception as e:
            logger.error("Error creating monthly comparison: %s", e)
            return None

    def create_case_status_trend(self, pivot_df: pd.DataFrame, client_name: str = "") -> str:
        try:
            fig, axes = plt.subplots(1, 2, figsize=(14, 5.5))

            df_chart = pivot_df[pivot_df["Year"] != "Grand Total"].copy()

            if not df_chart.empty:
                months = df_chart["Month"].tolist()
                completed = df_chart["Completed"].tolist() if "Completed" in df_chart.columns else [0] * len(months)
                wip = df_chart["Work in Progress"].tolist() if "Work in Progress" in df_chart.columns else [0] * len(months)
                closed = df_chart["Closed by Client"].tolist() if "Closed by Client" in df_chart.columns else [0] * len(months)
                insufficient = df_chart["Insufficient"].tolist() if "Insufficient" in df_chart.columns else [0] * len(months)

                x = range(len(months))

                axes[0].plot(x, completed, marker="o", linewidth=2, markersize=6,
                             label="Completed", color=self.colors.accent_normalized)
                axes[0].plot(x, wip, marker="s", linewidth=2, markersize=6,
                             label="Work in Progress", color=self.colors.warning_normalized)
                axes[0].plot(x, closed, marker="^", linewidth=2, markersize=6,
                             label="Closed by Client", color=self.colors.primary_normalized)
                axes[0].plot(x, insufficient, marker="d", linewidth=2, markersize=6,
                             label="Insufficient", color=self.colors.dark_gray_normalized)

                axes[0].set_xlabel("Month", fontsize=12)
                axes[0].set_ylabel("Number of Cases", fontsize=12)
                axes[0].set_title("Case Status Monthly Trends", fontsize=14, fontweight="bold")
                axes[0].set_xticks(x)
                axes[0].set_xticklabels(months, rotation=45, ha="right")
                axes[0].legend(loc="best")
                axes[0].grid(True, alpha=0.3)

                axes[1].stackplot(x, completed, wip, closed, insufficient,
                                   labels=["Completed", "Work in Progress", "Closed by Client", "Insufficient"],
                                   colors=[self.colors.accent_normalized, self.colors.warning_normalized,
                                           self.colors.primary_normalized, self.colors.dark_gray_normalized],
                                   alpha=0.7)
                axes[1].set_xlabel("Month", fontsize=12)
                axes[1].set_ylabel("Case Volume", fontsize=12)
                axes[1].set_title("Case Volume Composition (Stacked)", fontsize=14, fontweight="bold")
                axes[1].set_xticks(x)
                axes[1].set_xticklabels(months, rotation=45, ha="right")
                axes[1].legend(loc="upper left", fontsize=9)
                axes[1].grid(True, alpha=0.3)

            plt.suptitle(f"Case Status Analysis - {client_name}" if client_name else "Case Status Analysis", fontsize=14, fontweight="bold")
            plt.tight_layout()
            path = os.path.join(self.temp_dir, f"case_status_trend_{client_name.replace(' ', '_')}.png")
            plt.savefig(path, dpi=Config.CHART_DPI, bbox_inches="tight", facecolor="white")
            plt.close()
            return path
        except Exception as e:
            logger.error("Error creating case status trend: %s", e)
            return None

    def create_bucket_comparison_chart(self, recal_counts: pd.Series, final_counts: pd.Series, title: str, client_name: str = "") -> str:
        """Grouped bar chart comparing an Ageing bucket distribution (Recal
        vs Final) - used for both 'Complete Cases Ageing' and the
        WIP-specific breakdown."""
        try:
            buckets = list(dict.fromkeys(list(recal_counts.index) + list(final_counts.index)))
            recal_vals = [int(recal_counts.get(b, 0)) for b in buckets]
            final_vals = [int(final_counts.get(b, 0)) for b in buckets]

            fig, ax = plt.subplots(figsize=(11, 5.5))
            x = np.arange(len(buckets))
            width = 0.35

            ax.bar(x - width / 2, recal_vals, width, label="Recal", color=self.colors.secondary_normalized, alpha=0.9)
            ax.bar(x + width / 2, final_vals, width, label="Final", color=self.colors.primary_normalized, alpha=0.9)

            ax.set_xlabel("Ageing Bucket (days)", fontsize=12)
            ax.set_ylabel("Number of Cases", fontsize=12)
            full_title = f"{title} - {client_name}" if client_name else title
            ax.set_title(full_title, fontsize=14, fontweight="bold")
            ax.set_xticks(x)
            ax.set_xticklabels(buckets)
            ax.legend(loc="best")
            ax.grid(True, axis="y", alpha=0.3)

            for i, v in enumerate(recal_vals):
                if v:
                    ax.text(i - width / 2, v, str(v), ha="center", va="bottom", fontsize=8)
            for i, v in enumerate(final_vals):
                if v:
                    ax.text(i + width / 2, v, str(v), ha="center", va="bottom", fontsize=8)

            plt.tight_layout()
            path = os.path.join(self.temp_dir, f"ageing_comparison_{title.lower().replace(' ', '_')}_{client_name.replace(' ', '_')}.png")
            plt.savefig(path, dpi=Config.CHART_DPI, bbox_inches="tight", facecolor="white")
            plt.close()
            return path
        except Exception as e:
            logger.error("Error creating bucket comparison chart: %s", e)
            return None

    def create_severity_chart(self, severity_pivot_df: pd.DataFrame, client_name: str = "") -> str:
        """Pie chart of the severity distribution, from the pivot's Grand Total row."""
        try:
            grand_total = severity_pivot_df[severity_pivot_df["Year"] == "Grand Total"]
            if grand_total.empty:
                return None

            severity_cols = [c for c in severity_pivot_df.columns if c not in ("Year", "Month", "Grand Total")]
            values = [float(grand_total[c].iloc[0]) for c in severity_cols]
            labels = severity_cols

            nonzero = [(l, v) for l, v in zip(labels, values) if v > 0]
            if not nonzero:
                return None
            labels, values = zip(*nonzero)

            palette = [
                self.colors.accent_normalized, self.colors.warning_normalized, self.colors.primary_normalized,
                self.colors.secondary_normalized, self.colors.dark_gray_normalized,
            ]
            colors = [palette[i % len(palette)] for i in range(len(labels))]

            fig, ax = plt.subplots(figsize=(9, 6.5))
            ax.pie(
                values, labels=labels, autopct="%1.1f%%", startangle=90, colors=colors,
                wedgeprops={"edgecolor": "white", "linewidth": 1.5}, textprops={"fontsize": 10},
            )
            full_title = f"Report Severity Distribution - {client_name}" if client_name else "Report Severity Distribution"
            ax.set_title(full_title, fontsize=14, fontweight="bold")

            plt.tight_layout()
            path = os.path.join(self.temp_dir, f"severity_pie_{client_name.replace(' ', '_')}.png")
            plt.savefig(path, dpi=Config.CHART_DPI, bbox_inches="tight", facecolor="white")
            plt.close()
            return path
        except Exception as e:
            logger.error("Error creating severity chart: %s", e)
            return None
