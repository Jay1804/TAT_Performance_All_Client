"""Shared constants, color scheme and column-name mapping for the Casewise TAT report.

Column names here map to the normalized (lower_snake_case) form produced by
``data_processor.normalize_columns`` from the aliases used in
``sql_queries.QUERY_BASE_DATA``.
"""
from dataclasses import dataclass
from typing import Tuple


@dataclass
class ColorScheme:
    primary: Tuple[int, int, int] = (0, 82, 147)
    secondary: Tuple[int, int, int] = (198, 89, 17)
    accent: Tuple[int, int, int] = (0, 112, 60)
    warning: Tuple[int, int, int] = (191, 13, 62)
    dark_gray: Tuple[int, int, int] = (64, 64, 64)
    light_gray: Tuple[int, int, int] = (242, 242, 242)
    white: Tuple[int, int, int] = (255, 255, 255)

    def rgb_normalized(self, rgb: Tuple[int, int, int]) -> Tuple[float, float, float]:
        return (rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0)

    @property
    def primary_normalized(self) -> Tuple[float, float, float]:
        return self.rgb_normalized(self.primary)

    @property
    def secondary_normalized(self) -> Tuple[float, float, float]:
        return self.rgb_normalized(self.secondary)

    @property
    def accent_normalized(self) -> Tuple[float, float, float]:
        return self.rgb_normalized(self.accent)

    @property
    def warning_normalized(self) -> Tuple[float, float, float]:
        return self.rgb_normalized(self.warning)

    @property
    def dark_gray_normalized(self) -> Tuple[float, float, float]:
        return self.rgb_normalized(self.dark_gray)


class Config:
    CHART_DPI = 150
    TAT_TARGET = 95.0


# Columns as they appear after normalize_columns() lower-snake-cases the SQL aliases.
COL_RECEIVED_DATE = "received_date"
COL_CASE_ID = "case_ars_no"
COL_COMPANY = "company_name"
COL_PROCESS = "process_name"
COL_CASE_STATUS = "case_status"
COL_CLIENT_EXTERNAL_ID = "client_external_id"
COL_SEVERITY = "latest_report_severity"
COL_TAT_FINAL = "tat_status_final"
COL_TAT_RECAL = "tat_status_recal"

RAW_DATA_DROP_COLUMNS = [
    "received_date_parsed",
    "year",
    "month_num",
    "month_name",
    "day",
    "cat",
    "cat_tl",
    "account_manager",
    "rn",
]
