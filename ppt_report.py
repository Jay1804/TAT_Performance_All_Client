"""PowerPoint deck generation.

Every slide is built on the blank layout so the whole deck shares one
consistent, branded look (header bar, accent rule, footer) instead of mixing
in PowerPoint's default themed layouts. Slides map to a small set of
reusable primitives: a title slide, a section-divider slide (with optional
KPI cards), a table slide, a chart slide, an executive-summary slide, and a
thank-you slide.
"""
import os
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

from config import ColorScheme

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
HEADER_H = Inches(1.0)
FOOTER_Y = Inches(7.08)
MARGIN = Inches(0.5)
TOTAL_ROW_LABELS = {"grand total", "total"}
PCT_ROW_LABELS = {"contribution %", "percentage"}


class ProfessionalPPTGenerator:
    def __init__(self, temp_dir: str, client_name: str = ""):
        self.temp_dir = temp_dir
        self.client_name = client_name
        self.prs = Presentation()
        self.prs.slide_width = SLIDE_W
        self.prs.slide_height = SLIDE_H
        self.colors = ColorScheme()
        self._page = 0

    # ------------------------------------------------------------------ #
    # Shared primitives
    # ------------------------------------------------------------------ #
    def _blank_slide(self):
        return self.prs.slides.add_slide(self.prs.slide_layouts[6])

    def _rect(self, slide, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE, line=False):
        shp = slide.shapes.add_shape(shape, x, y, w, h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = RGBColor(*color)
        if line:
            shp.line.color.rgb = RGBColor(*color)
        else:
            shp.line.fill.background()
        shp.shadow.inherit = False
        return shp

    def _text(self, slide, x, y, w, h, text, size=14, bold=False, color=None,
              align=PP_ALIGN.LEFT, italic=False, anchor=None, font_name=None):
        box = slide.shapes.add_textbox(x, y, w, h)
        tf = box.text_frame
        tf.word_wrap = True
        if anchor is not None:
            tf.vertical_anchor = anchor
        tf.text = text
        p = tf.paragraphs[0]
        p.font.size = Pt(size)
        p.font.bold = bold
        p.font.italic = italic
        p.font.color.rgb = RGBColor(*(color or self.colors.dark_gray))
        if font_name:
            p.font.name = font_name
        p.alignment = align
        return box

    def _add_header(self, slide, title: str, kicker: Optional[str] = None):
        self._rect(slide, Emu(0), Emu(0), SLIDE_W, HEADER_H, self.colors.primary)
        self._rect(slide, Emu(0), HEADER_H, SLIDE_W, Inches(0.05), self.colors.secondary)

        title_y = Inches(0.32) if kicker else Inches(0.22)
        title_h = Inches(0.55) if kicker else Inches(0.6)
        if kicker:
            self._text(slide, MARGIN, Inches(0.1), Inches(11), Inches(0.3), kicker.upper(),
                       size=11, bold=True, color=(255, 255, 255) if False else self.colors.secondary)
        self._text(slide, MARGIN, title_y, Inches(12.3), title_h, title,
                   size=24, bold=True, color=(255, 255, 255))

    def _add_footer(self, slide):
        self._rect(slide, MARGIN, FOOTER_Y, Inches(12.333), Emu(9525), self.colors.light_gray)
        label = f"{self.client_name} | TAT Performance Dashboard" if self.client_name else "TAT Performance Dashboard"
        self._text(slide, MARGIN, Inches(7.13), Inches(8), Inches(0.3), label,
                   size=9, color=self.colors.dark_gray)
        self._page += 1
        self._text(slide, Inches(11.833), Inches(7.13), Inches(1), Inches(0.3), str(self._page),
                   size=9, color=self.colors.dark_gray, align=PP_ALIGN.RIGHT)

    def _decorated_slide(self, title: str, kicker: Optional[str] = None):
        slide = self._blank_slide()
        self._add_header(slide, title, kicker)
        self._add_footer(slide)
        return slide

    def _kpi_card(self, slide, x, y, w, h, label: str, value: str):
        card = self._rect(slide, x, y, w, h, (255, 255, 255), shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        card.line.color.rgb = RGBColor(*self.colors.light_gray)
        card.line.width = Pt(1)
        self._text(slide, x, y + Inches(0.12), w, Inches(0.5), value, size=26, bold=True,
                   color=self.colors.primary, align=PP_ALIGN.CENTER)
        self._text(slide, x, y + h - Inches(0.42), w, Inches(0.3), label, size=11,
                   color=self.colors.dark_gray, align=PP_ALIGN.CENTER)

    # ------------------------------------------------------------------ #
    # Title / divider / thank-you slides
    # ------------------------------------------------------------------ #
    def add_professional_title_slide(self, title: str, subtitle: str, metrics: Dict, date_range: str = ""):
        slide = self._blank_slide()
        self._rect(slide, Emu(0), Emu(0), SLIDE_W, SLIDE_H, self.colors.primary)
        self._rect(slide, Emu(0), Inches(5.5), SLIDE_W, Inches(0.06), self.colors.secondary)

        display_title = f"{title}\n{self.client_name}" if self.client_name else title
        self._text(slide, Inches(0.5), Inches(1.7), Inches(12.333), Inches(1.6), display_title,
                   size=42, bold=True, color=(255, 255, 255), align=PP_ALIGN.CENTER)
        self._text(slide, Inches(0.5), Inches(3.35), Inches(12.333), Inches(0.7), subtitle,
                   size=20, color=(255, 255, 255), align=PP_ALIGN.CENTER)
        self._text(slide, Inches(0.5), Inches(4.15), Inches(12.333), Inches(0.45), date_range,
                   size=13, color=self.colors.light_gray, align=PP_ALIGN.CENTER)

        if metrics:
            kpis = [
                ("Total Cases", str(metrics.get("total_cases", "N/A"))),
                ("Final IN TAT", str(metrics.get("final_in_tat", "N/A"))),
                ("Recal IN TAT", str(metrics.get("recal_in_tat", "N/A"))),
            ]
            card_w, gap = Inches(3.4), Inches(0.4)
            total_w = card_w * 3 + gap * 2
            start_x = int((SLIDE_W - total_w) / 2)
            for i, (label, value) in enumerate(kpis):
                x = start_x + i * (card_w + gap)
                card = self._rect(slide, x, Inches(5.85), card_w, Inches(0.95), self.colors.primary,
                                  shape=MSO_SHAPE.ROUNDED_RECTANGLE)
                card.line.color.rgb = RGBColor(255, 255, 255)
                card.line.width = Pt(1.25)
                self._text(slide, x, Inches(5.93), card_w, Inches(0.45), value, size=22, bold=True,
                           color=(255, 255, 255), align=PP_ALIGN.CENTER)
                self._text(slide, x, Inches(6.42), card_w, Inches(0.3), label, size=10,
                           color=self.colors.light_gray, align=PP_ALIGN.CENTER)

        self._text(slide, Inches(0.5), Inches(7.05), Inches(12.333), Inches(0.35),
                   datetime.now().strftime("%B %d, %Y"), size=12, color=self.colors.light_gray, align=PP_ALIGN.CENTER)

    def add_section_divider_slide(self, section_label: str, subtitle: str = "", kpis: Optional[List[Tuple[str, str]]] = None):
        slide = self._blank_slide()
        self._rect(slide, Emu(0), Emu(0), SLIDE_W, SLIDE_H, self.colors.secondary)
        self._rect(slide, Emu(0), Inches(3.55), SLIDE_W, Inches(0.06), (255, 255, 255))

        self._text(slide, Inches(0.5), Inches(2.35), Inches(12.333), Inches(1.1), section_label,
                   size=40, bold=True, color=(255, 255, 255), align=PP_ALIGN.CENTER)
        if subtitle:
            self._text(slide, Inches(0.5), Inches(3.75), Inches(12.333), Inches(0.55), subtitle,
                       size=16, color=(255, 255, 255), align=PP_ALIGN.CENTER)

        if kpis:
            card_w, gap = Inches(3.4), Inches(0.4)
            total_w = card_w * len(kpis) + gap * (len(kpis) - 1)
            start_x = int((SLIDE_W - total_w) / 2)
            for i, (label, value) in enumerate(kpis):
                x = start_x + i * (card_w + gap)
                self._kpi_card(slide, x, Inches(4.7), card_w, Inches(1.1), label, value)

    def add_thank_you_slide(self):
        slide = self._blank_slide()
        self._rect(slide, Emu(0), Emu(0), SLIDE_W, SLIDE_H, self.colors.primary)
        self._rect(slide, Emu(0), Inches(3.5), SLIDE_W, Inches(0.06), self.colors.secondary)

        self._text(slide, Inches(0.5), Inches(2.5), Inches(12.333), Inches(1.2), "Thank You",
                   size=52, bold=True, color=(255, 255, 255), align=PP_ALIGN.CENTER)
        self._text(slide, Inches(0.5), Inches(3.9), Inches(12.333), Inches(0.6),
                   "For your continued support and cooperation", size=20,
                   color=self.colors.light_gray, align=PP_ALIGN.CENTER)
        self._text(slide, Inches(0.5), Inches(6.9), Inches(12.333), Inches(0.4),
                   "TAT Performance Dashboard  |  Generated on " + datetime.now().strftime("%B %d, %Y"),
                   size=11, color=self.colors.light_gray, align=PP_ALIGN.CENTER)

    # ------------------------------------------------------------------ #
    # Executive summary
    # ------------------------------------------------------------------ #
    def add_executive_summary_slide(self, insights_list: List[str]):
        slide = self._decorated_slide("Executive Summary", kicker="Key Performance Insights")

        panel_x, panel_y = MARGIN, Inches(1.25)
        panel_w, panel_h = Inches(12.333), Inches(5.65)
        panel = self._rect(slide, panel_x, panel_y, panel_w, panel_h, self.colors.light_gray,
                           shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        panel.line.fill.background()

        box = slide.shapes.add_textbox(panel_x + Inches(0.35), panel_y + Inches(0.3),
                                        panel_w - Inches(0.7), panel_h - Inches(0.6))
        tf = box.text_frame
        tf.word_wrap = True
        first = True
        for insight in insights_list:
            clean = insight.replace("•", "").strip()
            if not clean:
                continue
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.text = f"●  {clean}"
            p.font.size = Pt(16)
            p.font.color.rgb = RGBColor(*self.colors.dark_gray)
            p.space_after = Pt(16)
            run = p.runs[0]
            run.font.size = Pt(16)

    # ------------------------------------------------------------------ #
    # Table slides
    # ------------------------------------------------------------------ #
    def add_table_slide(self, title: str, df: pd.DataFrame, description: str = "",
                        kicker: Optional[str] = None, highlight_last_col: bool = False, max_rows: int = 30):
        slide = self._decorated_slide(title, kicker)

        top = Inches(1.2)
        if description:
            self._text(slide, MARGIN, top, Inches(12.333), Inches(0.35), description,
                       size=11, italic=True, color=self.colors.dark_gray)
            top = top + Inches(0.4)

        rows, cols = len(df), len(df.columns)
        if rows == 0 or cols == 0:
            self._text(slide, MARGIN, top + Inches(0.3), Inches(12), Inches(0.5),
                       "No data available for this section.", size=13, italic=True)
            return

        truncated = rows > max_rows
        shown_rows = min(rows, max_rows)
        left, width = MARGIN, Inches(12.333)
        avail_h = Inches(7.0) - top
        row_h = min(Inches(0.4), Emu(int(avail_h / (shown_rows + 1))))

        data_font_size = 10 if cols <= 8 else (9 if cols <= 12 else 8)
        header_font_size = data_font_size + 1

        table_shape = slide.shapes.add_table(shown_rows + 1, cols, left, top, width, row_h * (shown_rows + 1))
        table = table_shape.table

        first_col_letters = int(width / Emu(1)) // cols
        for col_idx in range(cols):
            if col_idx < 2:
                table.columns[col_idx].width = Emu(int(width * 0.7 / cols))

        for col_idx, col_name in enumerate(df.columns):
            cell = table.cell(0, col_idx)
            cell.text = str(col_name)
            para = cell.text_frame.paragraphs[0]
            para.font.bold = True
            para.font.size = Pt(header_font_size)
            para.font.color.rgb = RGBColor(255, 255, 255)
            para.alignment = PP_ALIGN.CENTER
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(*self.colors.primary)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        for row_idx in range(shown_rows):
            row_label = str(df.iloc[row_idx, 0]).strip().lower()
            is_total_row = row_label in TOTAL_ROW_LABELS
            is_pct_row = row_label in PCT_ROW_LABELS

            for col_idx in range(cols):
                cell = table.cell(row_idx + 1, col_idx)
                value = df.iloc[row_idx, col_idx]
                cell.text = "" if pd.isna(value) else str(value)
                para = cell.text_frame.paragraphs[0]
                para.font.size = Pt(data_font_size)
                para.alignment = PP_ALIGN.CENTER
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                if col_idx == 0:
                    para.font.bold = True

                is_total_col = highlight_last_col and col_idx == cols - 1
                if is_total_row:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(*self.colors.secondary)
                    para.font.color.rgb = RGBColor(255, 255, 255)
                    para.font.bold = True
                elif is_pct_row:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(255, 235, 205)
                    para.font.italic = True
                elif is_total_col:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(*self.colors.primary)
                    para.font.color.rgb = RGBColor(255, 255, 255)
                    para.font.bold = True
                elif row_idx % 2 == 0:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(*self.colors.light_gray)
                else:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor(255, 255, 255)

        if truncated:
            note_y = top + row_h * (shown_rows + 1) + Inches(0.05)
            self._text(slide, left, note_y, width, Inches(0.3),
                       f"Showing first {shown_rows} of {rows} rows.", size=9, italic=True,
                       color=self.colors.dark_gray)

    # ------------------------------------------------------------------ #
    # Chart slides
    # ------------------------------------------------------------------ #
    def add_chart_slide(self, title: str, chart_path: str, insights: str, kicker: Optional[str] = None):
        slide = self._decorated_slide(title, kicker)

        chart_left, chart_top = MARGIN, Inches(1.3)
        chart_width = Inches(7.6)
        if chart_path and os.path.exists(chart_path):
            slide.shapes.add_picture(chart_path, chart_left, chart_top, width=chart_width)

        panel_x = Inches(8.35)
        panel_w = Inches(4.5)
        panel_h = Inches(5.55)
        panel = self._rect(slide, panel_x, chart_top, panel_w, panel_h, self.colors.light_gray,
                           shape=MSO_SHAPE.ROUNDED_RECTANGLE)
        panel.line.fill.background()

        self._text(slide, panel_x + Inches(0.3), chart_top + Inches(0.2), panel_w - Inches(0.6), Inches(0.4),
                   "Key Insights", size=14, bold=True, color=self.colors.primary)

        box = slide.shapes.add_textbox(panel_x + Inches(0.3), chart_top + Inches(0.7),
                                        panel_w - Inches(0.6), panel_h - Inches(0.9))
        tf = box.text_frame
        tf.word_wrap = True
        first = True
        for line in insights.split("\n"):
            clean = line.strip().lstrip("•").strip()
            if not clean:
                continue
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            is_heading = clean.endswith(":") or any(ch in clean for ch in ("📈", "📊", "📋", "🎯"))
            p.text = clean if is_heading else f"●  {clean}"
            p.font.size = Pt(13 if is_heading else 11.5)
            p.font.bold = is_heading
            p.font.color.rgb = RGBColor(*(self.colors.primary if is_heading else self.colors.dark_gray))
            p.space_after = Pt(10)

    def save(self, filename: str) -> str:
        path = os.path.join(self.temp_dir, filename)
        self.prs.save(path)
        return path
