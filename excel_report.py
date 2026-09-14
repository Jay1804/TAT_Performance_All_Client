"""Excel workbook built to use the exact same live-formula architecture as
the client's own 'Larsen & Toubro Limited Dashboard (Final+Flat).xls'
reference file: a single "Raw Data" sheet holds every case-level row in a
fixed column layout, and every number on the two summary tabs is a
COUNTIFS/SUMIFS formula that queries Raw Data directly by column letter -
not a pre-aggregated Python value. Reference formula patterns (verified by
opening the reference file with data_only=False):

  Combined Case Status + TAT block:
    Cases Received = COUNTIFS(month)                       [C8]
    status count    = COUNTIFS(month, case_status=header)  [D8..H8]
    IN/OUT TAT      = COUNTIFS(month, case_type="Sent",
                                tat_status=IT/OT, client)   [I8/J8]
    IN/OUT TAT %    = IN_TAT / SUM(Completed, Closed by Client)  [K8/L8]
    Grand Total row = SUM() over the block above

  Standalone TAT table:
    Cases Received/Completed Cases = COUNTIFS(month, client)
    IN/OUT TAT      = COUNTIFS(client, case_type="Sent", tat_status, month)
    IN/OUT TAT %    = IN_TAT / SUM(IN TAT, Out of TAT)

  Pending Cases TAT %:
    Total = SUM(IN TAT, Out of TAT); IN/OUT = COUNTIFS(client, status, tat_status)

  Ageing tables (monthly + status bi-furcation):
    bucket count = COUNTIFS(client, status/month, ageing_bucket=header)
    Total Cases  = SUM(bucket columns) ; Contribution % = bucket / total

  WIP Case Monthwise:
    IN/OUT TAT = COUNTIFS(client, case_status="Work in Progress", month, tat_status)

  Severity (shared by both tabs, no Recal/Final split):
    count = COUNTIFS(case_type="Sent", severity=header, month, client)
    Percentage row = column Grand Total / overall Grand Total

Recal Summary formulas reference the Recal columns (ageing_bucket,
tat_status_recal); Flat Summary references the Final columns
(ageing_bucket_final, tat_status_final). The client-name criterion is only
included when the report is scoped to a single real client (it's meaningless
against a synthetic "All Clients Combined" label).
"""
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from config import ColorScheme

_colors = ColorScheme()


def _rgb(rgb) -> str:
    return "{:02X}{:02X}{:02X}".format(*rgb)


HEADER_FILL = PatternFill(start_color=_rgb(_colors.primary), end_color=_rgb(_colors.primary), fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=10)
TITLE_FILL = PatternFill(start_color=_rgb(_colors.primary), end_color=_rgb(_colors.primary), fill_type="solid")
TITLE_FONT = Font(color="FFFFFF", bold=True, size=16)
SUBTITLE_FONT = Font(color=_rgb(_colors.dark_gray), size=10, italic=True)
SECTION_FILL = PatternFill(start_color=_rgb(_colors.secondary), end_color=_rgb(_colors.secondary), fill_type="solid")
SECTION_FONT = Font(bold=True, size=11, color="FFFFFF")
GRAND_TOTAL_FILL = PatternFill(start_color="E6E6E6", end_color="E6E6E6", fill_type="solid")
GRAND_TOTAL_FONT = Font(bold=True, size=10)
DATA_FONT = Font(size=10)
KPI_LABEL_FONT = Font(size=9, color=_rgb(_colors.dark_gray))
KPI_VALUE_FONT = Font(size=18, bold=True, color=_rgb(_colors.primary))
THIN_BORDER = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
CENTER = Alignment(horizontal="center", vertical="center")
PCT_FORMAT = "0.00%"

SHEET_WIDTH_COLS = 12

# Fixed Raw Data column layout - column letters below (RAW["case_status"] etc.)
# are hardcoded into every COUNTIFS formula, so this order must not change
# without updating every formula builder that references it.
RAW_DATA_COLUMNS = [
    "company_name", "candidate_name", "case_ars_no", "location", "process_name",
    "received_date", "month_name", "latest_check_insuff_raised_date",
    "latest_check_insuff_fulfilled_date", "go_ahead_date", "reopen_dates",
    "final_report_sent", "case_status", "case_type", "latest_report_severity",
    "process_tat", "start_date", "ageing", "ageing_bucket", "tat_status_recal",
    "opt_final_due_date", "ageing_final", "ageing_bucket_final", "tat_status_final",
]
RAW = {name: get_column_letter(i + 1) for i, name in enumerate(RAW_DATA_COLUMNS)}

STATUS_COLUMNS = ["Completed", "Closed by Client", "Insufficient", "Work in Progress", "On Hold"]
AGEING_BUCKETS = ["0--2", "3--10", "11--14", "15--20", "21--25", "26--30", "30 +"]
SEVERITY_ORDER = [
    "Clear/Green", "No Response Received/Amber", "Discrepant/Red", "Attention Required/Amber",
    "Insufficient/Amber", "Closed by client", "Minor Discrepant/Red",
]
RAW_SHEET_NAME = "Raw Data"


def _month_list(df: pd.DataFrame) -> List[Tuple[int, str]]:
    m = df[["year", "month_num", "month_name"]].drop_duplicates().sort_values(["year", "month_num"])
    return list(zip(m["year"].tolist(), m["month_name"].tolist()))


def _severity_list(df: pd.DataFrame) -> List[str]:
    present = set(df["latest_report_severity"].dropna().unique()) if "latest_report_severity" in df.columns else set()
    ordered = [s for s in SEVERITY_ORDER if s in present]
    ordered += sorted(s for s in present if s not in SEVERITY_ORDER)
    return ordered


def _countifs(raw_sheet: str, *criteria: Tuple[str, str]) -> str:
    """criteria: (raw_column_name, criteria_expr) pairs. criteria_expr is a
    literal Excel expression (already quoted if it's a string literal, or a
    cell reference like $B8)."""
    parts = []
    for col_name, expr in criteria:
        parts.append(f"'{raw_sheet}'!${RAW[col_name]}:${RAW[col_name]}")
        parts.append(expr)
    return f"=COUNTIFS({','.join(parts)})"


def _write_header_block(ws, client_name: str, date_range_str: str, variant_label: str, primary_sheet_ref: Optional[str]) -> int:
    """Row 1: 'Client Name:-' / B1 (matches the reference exactly - the title
    bar's CONCATENATE(B1,...) formula depends on the client name being in
    row 1, not wherever the title happens to sit). Row 2 blank spacer.
    Row 3: title bar. Row 4: subtitle. Returns the next free row."""
    ws.cell(row=1, column=1, value="Client Name:-")
    b1 = ws.cell(row=1, column=2)
    if primary_sheet_ref:
        b1.value = f"='{primary_sheet_ref}'!B1"
    else:
        b1.value = client_name

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=SHEET_WIDTH_COLS)
    title_cell = ws.cell(row=3, column=1, value=f'=CONCATENATE(B1," - Dashboard ({variant_label})")')
    title_cell.fill = TITLE_FILL
    title_cell.font = TITLE_FONT
    title_cell.alignment = CENTER
    ws.row_dimensions[3].height = 30

    ws.merge_cells(start_row=4, start_column=1, end_row=4, end_column=SHEET_WIDTH_COLS)
    subtitle = ws.cell(
        row=4, column=1,
        value=f"Reporting Period: {date_range_str}    |    Generated: {datetime.now().strftime('%B %d, %Y')}",
    )
    subtitle.font = SUBTITLE_FONT
    subtitle.alignment = CENTER
    return 6


def _write_kpi_strip(ws, row: int, total_cases_ref: str, in_tat_pct_ref: str, performance_status: str) -> int:
    kpis = [("Total Cases", total_cases_ref), ("IN TAT %", in_tat_pct_ref), ("Performance Status", performance_status)]
    cols_per_kpi = SHEET_WIDTH_COLS // len(kpis)

    for i, (label, value) in enumerate(kpis):
        start_col = i * cols_per_kpi + 1
        end_col = start_col + cols_per_kpi - 1
        ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
        label_cell = ws.cell(row=row, column=start_col, value=label)
        label_cell.font = KPI_LABEL_FONT
        label_cell.alignment = CENTER

        ws.merge_cells(start_row=row + 1, start_column=start_col, end_row=row + 1, end_column=end_col)
        value_cell = ws.cell(row=row + 1, column=start_col, value=value)
        value_cell.font = KPI_VALUE_FONT
        value_cell.alignment = CENTER
        if i == 1:
            value_cell.number_format = PCT_FORMAT

    ws.row_dimensions[row + 1].height = 26
    return row + 3


def _style_row(ws, row: int, n_cols: int, is_header: bool, is_total: bool, highlight_last_col: bool = False):
    for col_idx in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col_idx)
        cell.border = THIN_BORDER
        cell.alignment = CENTER
        if is_header:
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
        elif is_total or (highlight_last_col and col_idx == n_cols):
            cell.fill = GRAND_TOTAL_FILL
            cell.font = GRAND_TOTAL_FONT
        else:
            cell.font = DATA_FONT


def _section_title(ws, row: int, title: str, n_cols: int):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=n_cols)
    cell = ws.cell(row=row, column=1, value=title)
    cell.fill = SECTION_FILL
    cell.font = SECTION_FONT


def _dual_section_title(ws, row: int, left_title: str, left_cols: int, right_title: str, right_cols: int):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=left_cols)
    left_cell = ws.cell(row=row, column=1, value=left_title)
    left_cell.fill = SECTION_FILL
    left_cell.font = SECTION_FONT

    ws.merge_cells(start_row=row, start_column=left_cols + 1, end_row=row, end_column=left_cols + right_cols)
    right_cell = ws.cell(row=row, column=left_cols + 1, value=right_title)
    right_cell.fill = SECTION_FILL
    right_cell.font = SECTION_FONT


class SummarySheetBuilder:
    """Writes one formula-driven summary tab (Recal or Flat).

    No formula here filters by client name against Raw Data - Raw Data is
    already scoped to exactly the client(s)/date range the user selected
    upstream (in the SQL pull), so an additional client-name COUNTIFS
    criterion would be redundant at best, and wrong whenever the report
    scope is a synthetic multi-client label that matches no real row.
    """

    def __init__(self, ws, raw_sheet: str, ageing_col: str, tat_col: str):
        self.ws = ws
        self.raw = raw_sheet
        self.ageing_col = ageing_col  # "ageing_bucket" or "ageing_bucket_final"
        self.tat_col = tat_col  # "tat_status_recal" or "tat_status_final"

    # ---- Section 1: combined Case Status + TAT --------------------------
    def write_combined_block(self, row: int, months: List[Tuple[int, str]]) -> int:
        ws = self.ws
        n_cols = 3 + len(STATUS_COLUMNS) + 4  # Year,Month,Cases Received + statuses + IN/OUT TAT + 2 pct
        _dual_section_title(
            ws, row, "Over All Cases Position (WIP, Completed, Insufficient and CCC)", 3 + len(STATUS_COLUMNS),
            "Sent Cases TAT Analysis", 4,
        )
        row += 1

        header = ["Year", "Month", "Cases Received"] + STATUS_COLUMNS + ["IN TAT", "Out of TAT", "IN TAT %", "Out of TAT %"]
        for col_idx, name in enumerate(header, start=1):
            ws.cell(row=row, column=col_idx, value=name)
        _style_row(ws, row, n_cols, is_header=True, is_total=False)
        header_row = row
        row += 1
        data_start = row

        for year, month in months:
            month_cell = f"$B{row}"
            ws.cell(row=row, column=1, value=year)
            ws.cell(row=row, column=2, value=month)
            ws.cell(row=row, column=3, value=_countifs(self.raw, ("month_name", month_cell)))
            for col_offset, status in enumerate(STATUS_COLUMNS):
                col = 4 + col_offset
                ws.cell(row=row, column=col, value=_countifs(
                    self.raw, ("month_name", month_cell), ("case_status", f'"{status}"'),
                ))
            in_col = 4 + len(STATUS_COLUMNS)
            out_col = in_col + 1
            ws.cell(row=row, column=in_col, value=_countifs(
                self.raw, ("month_name", month_cell), ("case_type", '"Sent"'), (self.tat_col, '"IT"'),
            ))
            ws.cell(row=row, column=out_col, value=_countifs(
                self.raw, ("month_name", month_cell), ("case_type", '"Sent"'), (self.tat_col, '"OT"'),
            ))
            completed_letter = get_column_letter(4)
            closed_letter = get_column_letter(5)
            denom = f"SUM(${completed_letter}{row}:${closed_letter}{row})"
            in_letter = get_column_letter(in_col)
            out_letter = get_column_letter(out_col)
            ws.cell(row=row, column=in_col + 2, value=f"=IFERROR({in_letter}{row}/{denom},0)").number_format = PCT_FORMAT
            ws.cell(row=row, column=in_col + 3, value=f"=IFERROR({out_letter}{row}/{denom},0)").number_format = PCT_FORMAT
            _style_row(ws, row, n_cols, is_header=False, is_total=False)
            row += 1

        data_end = row - 1
        # Grand total: SUM the raw count columns (Cases Received..Out of TAT), then percentages
        in_col = 4 + len(STATUS_COLUMNS)
        out_col = in_col + 1
        for col_idx in range(3, out_col + 1):
            letter = get_column_letter(col_idx)
            ws.cell(row=row, column=col_idx, value=f"=SUM({letter}{data_start}:{letter}{data_end})")
        completed_letter, closed_letter = get_column_letter(4), get_column_letter(5)
        in_letter, out_letter = get_column_letter(in_col), get_column_letter(out_col)
        denom = f"SUM(${completed_letter}{row},${closed_letter}{row})"
        ws.cell(row=row, column=in_col + 2, value=f"={in_letter}{row}/{denom}").number_format = PCT_FORMAT
        ws.cell(row=row, column=in_col + 3, value=f"={out_letter}{row}/{denom}").number_format = PCT_FORMAT
        ws.cell(row=row, column=1, value="Grand Total")
        _style_row(ws, row, n_cols, is_header=False, is_total=True)
        self.combined_grand_total_row = row
        self.combined_in_tat_col = in_col
        self.combined_pct_col = in_col + 2
        return row + 2

    # ---- Section 2: standalone TAT table ---------------------------------
    def write_tat_table(self, row: int, title: str, months: List[Tuple[int, str]]) -> int:
        ws = self.ws
        n_cols = 8
        _section_title(ws, row, title, n_cols)
        row += 1

        header = ["Year", "Month", "Cases Received", "Completed Cases", "IN TAT", "Out of TAT", "IN TAT %", "Out of TAT %"]
        for col_idx, name in enumerate(header, start=1):
            ws.cell(row=row, column=col_idx, value=name)
        _style_row(ws, row, n_cols, is_header=True, is_total=False)
        row += 1
        data_start = row

        for year, month in months:
            month_cell = f"$B{row}"
            ws.cell(row=row, column=1, value=year)
            ws.cell(row=row, column=2, value=month)
            ws.cell(row=row, column=3, value=_countifs(self.raw, ("month_name", month_cell)))
            ws.cell(row=row, column=4, value=_countifs(
                self.raw, ("month_name", month_cell), ("case_status", '"Completed"'),
            ))
            ws.cell(row=row, column=5, value=_countifs(
                self.raw, ("case_type", '"Sent"'), (self.tat_col, '"IT"'), ("month_name", month_cell),
            ))
            ws.cell(row=row, column=6, value=_countifs(
                self.raw, ("case_type", '"Sent"'), (self.tat_col, '"OT"'), ("month_name", month_cell),
            ))
            ws.cell(row=row, column=7, value=f"=IFERROR(E{row}/SUM(E{row}:F{row}),0)").number_format = PCT_FORMAT
            ws.cell(row=row, column=8, value=f"=IFERROR(F{row}/SUM(E{row}:F{row}),0)").number_format = PCT_FORMAT
            _style_row(ws, row, n_cols, is_header=False, is_total=False)
            row += 1

        data_end = row - 1
        for col_idx in range(3, 7):
            letter = get_column_letter(col_idx)
            ws.cell(row=row, column=col_idx, value=f"=SUM({letter}{data_start}:{letter}{data_end})")
        ws.cell(row=row, column=7, value=f"=IFERROR(E{row}/SUM(E{row}:F{row}),0)").number_format = PCT_FORMAT
        ws.cell(row=row, column=8, value=f"=IFERROR(F{row}/SUM(E{row}:F{row}),0)").number_format = PCT_FORMAT
        ws.cell(row=row, column=1, value="Grand Total")
        _style_row(ws, row, n_cols, is_header=False, is_total=True)
        return row + 2

    # ---- Section 3: Pending Cases TAT % ------------------------------------
    def write_pending_tat(self, row: int) -> int:
        ws = self.ws
        n_cols = 6
        statuses = ["Work in Progress", "Insufficient", "On Hold"]
        _section_title(ws, row, "Pending Cases TAT %", n_cols)
        row += 1
        for col_idx, name in enumerate(["Case Status", "Total Cases", "IN TAT", "Out Of TAT", "IN TAT%", "OUT Of TAT%"], start=1):
            ws.cell(row=row, column=col_idx, value=name)
        _style_row(ws, row, n_cols, is_header=True, is_total=False)
        row += 1
        data_start = row

        for status in statuses:
            ws.cell(row=row, column=1, value=status)
            ws.cell(row=row, column=3, value=_countifs(self.raw, ("case_status", f'"{status}"'), (self.tat_col, '"IT"')))
            ws.cell(row=row, column=4, value=_countifs(self.raw, ("case_status", f'"{status}"'), (self.tat_col, '"OT"')))
            ws.cell(row=row, column=2, value=f"=SUM(C{row}:D{row})")
            ws.cell(row=row, column=5, value=f"=IFERROR(C{row}/B{row},0)").number_format = PCT_FORMAT
            ws.cell(row=row, column=6, value=f"=IFERROR(D{row}/B{row},0)").number_format = PCT_FORMAT
            _style_row(ws, row, n_cols, is_header=False, is_total=False)
            row += 1

        data_end = row - 1
        for col_idx in (2, 3, 4):
            letter = get_column_letter(col_idx)
            ws.cell(row=row, column=col_idx, value=f"=SUM({letter}{data_start}:{letter}{data_end})")
        ws.cell(row=row, column=5, value=f"=IFERROR(C{row}/B{row},0)").number_format = PCT_FORMAT
        ws.cell(row=row, column=6, value=f"=IFERROR(D{row}/B{row},0)").number_format = PCT_FORMAT
        ws.cell(row=row, column=1, value="Grand Total")
        _style_row(ws, row, n_cols, is_header=False, is_total=True)
        return row + 2

    # ---- Ageing (monthly) + contribution -----------------------------------
    def write_ageing_monthly(self, row: int, title: str, months: List[Tuple[int, str]]) -> Tuple[int, int]:
        """Returns (next_row, grand_total_row) - grand_total_row is needed by the contribution table."""
        ws = self.ws
        n_cols = 2 + len(AGEING_BUCKETS) + 1
        _section_title(ws, row, title, n_cols)
        row += 1
        for col_idx, name in enumerate(["Year", "Month"] + AGEING_BUCKETS + ["Total Cases"], start=1):
            ws.cell(row=row, column=col_idx, value=name)
        _style_row(ws, row, n_cols, is_header=True, is_total=False)
        row += 1
        data_start = row

        bucket_start_col = 3
        total_col = 2 + len(AGEING_BUCKETS) + 1
        for year, month in months:
            month_cell = f"$B{row}"
            ws.cell(row=row, column=1, value=year)
            ws.cell(row=row, column=2, value=month)
            for i, bucket in enumerate(AGEING_BUCKETS):
                col = bucket_start_col + i
                ws.cell(row=row, column=col, value=_countifs(
                    self.raw, (self.ageing_col, f'"{bucket}"'), ("month_name", month_cell), ("case_type", '"Sent"'),
                ))
            bucket_letters = [get_column_letter(bucket_start_col + i) for i in range(len(AGEING_BUCKETS))]
            ws.cell(row=row, column=total_col, value=f"=SUM({','.join(f'{l}{row}' for l in bucket_letters)})")
            _style_row(ws, row, n_cols, is_header=False, is_total=False)
            row += 1

        data_end = row - 1
        for col_idx in range(3, total_col + 1):
            letter = get_column_letter(col_idx)
            ws.cell(row=row, column=col_idx, value=f"=SUM({letter}{data_start}:{letter}{data_end})")
        ws.cell(row=row, column=1, value="Grand Total")
        _style_row(ws, row, n_cols, is_header=False, is_total=True)
        grand_total_row = row
        return row + 2, grand_total_row

    def write_ageing_contribution(
        self, row: int, title: str, source_grand_total_row: int,
        source_bucket_start_col: int, source_total_col: int,
    ) -> int:
        """source_bucket_start_col/source_total_col pinpoint where the
        buckets and the Total Cases column actually sit in the table this
        contribution table is summarizing - the monthly Ageing table has
        Total Cases AFTER the buckets (Year,Month,buckets...,Total), while
        the status bi-furcation table has it BEFORE (Status,Total,buckets...)."""
        ws = self.ws
        n_cols = 1 + len(AGEING_BUCKETS)
        _section_title(ws, row, title, n_cols)
        row += 1
        ws.cell(row=row, column=1, value="Ageing Bucket")
        for i, bucket in enumerate(AGEING_BUCKETS):
            ws.cell(row=row, column=2 + i, value=bucket)
        _style_row(ws, row, n_cols, is_header=True, is_total=False)
        row += 1

        total_col_letter = get_column_letter(source_total_col)
        ws.cell(row=row, column=1, value="Ageing Contribution")
        for i, bucket in enumerate(AGEING_BUCKETS):
            bucket_letter = get_column_letter(source_bucket_start_col + i)
            cell = ws.cell(row=row, column=2 + i, value=f"=IFERROR({bucket_letter}{source_grand_total_row}/{total_col_letter}{source_grand_total_row},0)")
            cell.number_format = PCT_FORMAT
        _style_row(ws, row, n_cols, is_header=False, is_total=False)
        return row + 2

    # ---- Status bi-furcation + ageing (WIP / Non-workable) -----------------
    def write_status_ageing(self, row: int, title: str, statuses: List[str]) -> Tuple[int, int]:
        ws = self.ws
        n_cols = 2 + len(AGEING_BUCKETS)
        _section_title(ws, row, title, n_cols)
        row += 1
        for col_idx, name in enumerate(["Status", "Total Cases"] + AGEING_BUCKETS, start=1):
            ws.cell(row=row, column=col_idx, value=name)
        _style_row(ws, row, n_cols, is_header=True, is_total=False)
        row += 1
        data_start = row

        bucket_start_col = 3
        for status in statuses:
            ws.cell(row=row, column=1, value=status)
            for i, bucket in enumerate(AGEING_BUCKETS):
                col = bucket_start_col + i
                ws.cell(row=row, column=col, value=_countifs(
                    self.raw, ("case_status", f'"{status}"'), (self.ageing_col, f'"{bucket}"'),
                ))
            bucket_letters = [get_column_letter(bucket_start_col + i) for i in range(len(AGEING_BUCKETS))]
            ws.cell(row=row, column=2, value=f"=SUM({','.join(f'{l}{row}' for l in bucket_letters)})")
            _style_row(ws, row, n_cols, is_header=False, is_total=False)
            row += 1

        data_end = row - 1
        for col_idx in range(2, n_cols + 1):
            letter = get_column_letter(col_idx)
            ws.cell(row=row, column=col_idx, value=f"=SUM({letter}{data_start}:{letter}{data_end})")
        ws.cell(row=row, column=1, value="Grand Total")
        _style_row(ws, row, n_cols, is_header=False, is_total=True)
        grand_total_row = row
        return row + 2, grand_total_row

    # ---- WIP Case Monthwise -------------------------------------------------
    def write_wip_monthly(self, row: int, months: List[Tuple[int, str]]) -> int:
        ws = self.ws
        n_cols = 7
        _section_title(ws, row, "Work In Progress - Case Monthwise", n_cols)
        row += 1
        for col_idx, name in enumerate(["Year", "Month", "Total Cases", "IN TAT", "Out of TAT", "IN TAT %", "Out of TAT %"], start=1):
            ws.cell(row=row, column=col_idx, value=name)
        _style_row(ws, row, n_cols, is_header=True, is_total=False)
        row += 1
        data_start = row

        for year, month in months:
            month_cell = f"$B{row}"
            ws.cell(row=row, column=1, value=year)
            ws.cell(row=row, column=2, value=month)
            ws.cell(row=row, column=4, value=_countifs(
                self.raw, ("case_status", '"Work in Progress"'), ("month_name", month_cell), (self.tat_col, '"IT"'),
            ))
            ws.cell(row=row, column=5, value=_countifs(
                self.raw, ("case_status", '"Work in Progress"'), ("month_name", month_cell), (self.tat_col, '"OT"'),
            ))
            ws.cell(row=row, column=3, value=f"=SUM(D{row}:E{row})")
            ws.cell(row=row, column=6, value=f"=IFERROR(D{row}/C{row},0)").number_format = PCT_FORMAT
            ws.cell(row=row, column=7, value=f"=IFERROR(E{row}/C{row},0)").number_format = PCT_FORMAT
            _style_row(ws, row, n_cols, is_header=False, is_total=False)
            row += 1

        data_end = row - 1
        for col_idx in (3, 4, 5):
            letter = get_column_letter(col_idx)
            ws.cell(row=row, column=col_idx, value=f"=SUM({letter}{data_start}:{letter}{data_end})")
        ws.cell(row=row, column=6, value=f"=IFERROR(D{row}/C{row},0)").number_format = PCT_FORMAT
        ws.cell(row=row, column=7, value=f"=IFERROR(E{row}/C{row},0)").number_format = PCT_FORMAT
        ws.cell(row=row, column=1, value="Grand Total")
        _style_row(ws, row, n_cols, is_header=False, is_total=True)
        return row + 2

    # ---- Severity (shared, no Recal/Final split) ---------------------------
    def write_severity(self, row: int, months: List[Tuple[int, str]], severities: List[str]) -> int:
        ws = self.ws
        n_cols = 3 + len(severities)
        _section_title(ws, row, "Sent Case Report - Severity Wise", n_cols)
        row += 1
        for col_idx, name in enumerate(["Year", "Month", "Completed Cases"] + severities, start=1):
            ws.cell(row=row, column=col_idx, value=name)
        _style_row(ws, row, n_cols, is_header=True, is_total=False)
        header_row = row
        row += 1
        data_start = row

        sev_start_col = 4
        for year, month in months:
            month_cell = f"$B{row}"
            ws.cell(row=row, column=1, value=year)
            ws.cell(row=row, column=2, value=month)
            for i in range(len(severities)):
                col = sev_start_col + i
                sev_header_cell = f"{get_column_letter(col)}${header_row}"
                ws.cell(row=row, column=col, value=_countifs(
                    self.raw, ("case_type", '"Sent"'), ("latest_report_severity", sev_header_cell),
                    ("month_name", month_cell),
                ))
            sev_letters = [get_column_letter(sev_start_col + i) for i in range(len(severities))]
            ws.cell(row=row, column=3, value=f"=SUM({','.join(f'{l}{row}' for l in sev_letters)})")
            _style_row(ws, row, n_cols, is_header=False, is_total=False)
            row += 1

        data_end = row - 1
        for col_idx in range(3, n_cols + 1):
            letter = get_column_letter(col_idx)
            ws.cell(row=row, column=col_idx, value=f"=SUM({letter}{data_start}:{letter}{data_end})")
        ws.cell(row=row, column=1, value="Grand Total")
        _style_row(ws, row, n_cols, is_header=False, is_total=True)
        grand_total_row = row
        row += 1

        ws.cell(row=row, column=1, value="Percentage")
        for i in range(len(severities)):
            col = sev_start_col + i
            letter = get_column_letter(col)
            cell = ws.cell(row=row, column=col, value=f"=IFERROR({letter}{grand_total_row}/$C${grand_total_row},0)")
            cell.number_format = PCT_FORMAT
        _style_row(ws, row, n_cols, is_header=False, is_total=False)
        for col_idx in range(1, n_cols + 1):
            ws.cell(row=row, column=col_idx).fill = GRAND_TOTAL_FILL
            ws.cell(row=row, column=col_idx).font = GRAND_TOTAL_FONT

        self.severity_grand_total_row = grand_total_row
        self.severity_total_col = 3
        return row + 2


def _autosize_columns(ws, max_width: int = 40):
    for column_cells in ws.columns:
        length = max((len(str(c.value)) for c in column_cells if c.value is not None), default=0)
        col_letter = get_column_letter(column_cells[0].column)
        ws.column_dimensions[col_letter].width = min(max(length + 2, 10), max_width)


def _write_raw_data_sheet(wb, df_raw_data: pd.DataFrame) -> int:
    ws = wb.create_sheet(RAW_SHEET_NAME)

    for col_idx, col_name in enumerate(RAW_DATA_COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER

    df = df_raw_data.copy()
    for col_name in RAW_DATA_COLUMNS:
        if col_name not in df.columns:
            df[col_name] = None
    df = df[RAW_DATA_COLUMNS]
    df = df.astype(object).where(df.notna(), None)

    for r_idx, row_values in enumerate(dataframe_to_rows(df, index=False, header=False), start=2):
        for col_idx, value in enumerate(row_values, start=1):
            if hasattr(value, "to_pydatetime"):
                value = value.to_pydatetime()
            cell = ws.cell(row=r_idx, column=col_idx, value=value)
            cell.border = THIN_BORDER

    _autosize_columns(ws, max_width=30)
    if len(df) > 0:
        ws.auto_filter.ref = ws.dimensions
    return len(df)


def _build_summary_sheet(
    wb, sheet_name: str, variant_label: str, client_name: str, date_range_str: str,
    df_filtered: pd.DataFrame, ageing_col: str, tat_col: str,
    performance_status: str, primary_sheet_ref: Optional[str],
) -> None:
    ws = wb.create_sheet(sheet_name)
    months = _month_list(df_filtered)
    severities = _severity_list(df_filtered)

    row = _write_header_block(ws, client_name, date_range_str, variant_label, primary_sheet_ref)

    builder = SummarySheetBuilder(ws, RAW_SHEET_NAME, ageing_col, tat_col)

    kpi_row = row
    row += 3  # KPI strip placeholder - filled in after we know the combined table's cell coordinates

    combined_start_row = row
    row = builder.write_combined_block(row, months)

    _write_kpi_strip(
        ws, kpi_row,
        total_cases_ref=f"=C{builder.combined_grand_total_row}",
        in_tat_pct_ref=f"={get_column_letter(builder.combined_pct_col)}{builder.combined_grand_total_row}",
        performance_status=performance_status,
    )

    row = builder.write_tat_table(row, f"Cases Received / Sent Cases TAT Analysis ({variant_label})", months)
    row = builder.write_pending_tat(row)

    row, ageing_gt_row = builder.write_ageing_monthly(row, "Complete Cases Ageing", months)
    # "Complete Cases Ageing" header is Year,Month,<7 buckets>,Total Cases -> buckets start at col 3, Total Cases at col 10.
    row = builder.write_ageing_contribution(row, "Ageing Bucket Contribution %", ageing_gt_row, source_bucket_start_col=3, source_total_col=2 + len(AGEING_BUCKETS) + 1)

    row, wip_gt_row = builder.write_status_ageing(row, "Work In Progress - Bi-furcation and Ageing", ["Work in Progress"])
    # "WIP Bi-furcation" header is Status,Total Cases,<7 buckets> -> Total Cases at col 2, buckets start at col 3.
    row = builder.write_ageing_contribution(row, "Work In Progress - Ageing Contribution %", wip_gt_row, source_bucket_start_col=3, source_total_col=2)

    row = builder.write_wip_monthly(row, months)

    row, _ = builder.write_status_ageing(row, "Non-Workable Cases - Bi-furcation and Ageing", ["Insufficient", "On Hold"])

    row = builder.write_severity(row, months, severities)

    _autosize_columns(ws)
    ws.freeze_panes = "A2"


class ExcelReportGenerator:
    @staticmethod
    def create_excel_report(
        df_filtered: pd.DataFrame,
        client_name: str,
        date_range_str: str,
        insights: Dict[str, Any],
        output_path: str,
    ) -> str:
        wb = Workbook()
        wb.remove(wb.active)

        performance_status = insights.get("performance_status", "N/A")

        # "Flat Summary" holds the literal client name (B1); "Recal Summary"
        # references it by formula - matching the reference file's own
        # B1 = ='Final Summary "Flat"'!B1 pattern.
        _build_summary_sheet(
            wb, "Flat Summary", "Final / Flat", client_name, date_range_str,
            df_filtered, "ageing_bucket_final", "tat_status_final",
            performance_status, primary_sheet_ref=None,
        )
        _build_summary_sheet(
            wb, "Recal Summary", "Recal", client_name, date_range_str,
            df_filtered, "ageing_bucket", "tat_status_recal",
            performance_status, primary_sheet_ref="Flat Summary",
        )

        _write_raw_data_sheet(wb, df_filtered)

        # Put Recal Summary first (matches the reference's tab order), Flat Summary second.
        wb.move_sheet("Recal Summary", offset=-1)

        wb.save(output_path)
        return output_path
