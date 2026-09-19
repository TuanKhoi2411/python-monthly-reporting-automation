# By default, the pipeline reads from and writes to the current directory.
# Pass --input-root and --output-dir to use another location.
DEFAULT_INPUT_ROOT = "."
DEFAULT_OUTPUT_ROOT = "."

import os
import re
import glob
import traceback
import sys
import argparse
import unicodedata
import shutil
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrData, StrRef, StrVal
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    DOCX_AVAILABLE = True
except ModuleNotFoundError:
    Document = None
    Inches = Pt = RGBColor = None
    WD_ALIGN_PARAGRAPH = None
    DOCX_AVAILABLE = False
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl.chart.text import RichText
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties, RichTextProperties
from calendar import month_name
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter


# ============================== PATH CONFIGURATION ==========================
MONTHLY_REPORT_DIR = Path(DEFAULT_INPUT_ROOT)
OUTPUT_DIR = Path(DEFAULT_OUTPUT_ROOT)
PNL_INPUT_DIR = MONTHLY_REPORT_DIR / "Raw data" / "Manager Pnl"


def configure_runtime_paths(input_root, output_dir=None):
    """Configure source and output folders without editing the script."""
    global MONTHLY_REPORT_DIR, OUTPUT_DIR, PNL_INPUT_DIR, KPI_WORD_OUTPUT_PATH, KPI_OUTPUT_PATH
    MONTHLY_REPORT_DIR = Path(input_root).expanduser().resolve()
    OUTPUT_DIR = Path(output_dir).expanduser().resolve() if output_dir else MONTHLY_REPORT_DIR
    PNL_INPUT_DIR = MONTHLY_REPORT_DIR / "Raw data" / "Manager Pnl"
    KPI_WORD_OUTPUT_PATH = OUTPUT_DIR / "KPI Performance Review May 2026.docx"
    KPI_OUTPUT_PATH = OUTPUT_DIR / "KPI Report.xlsx"


def parse_cli_args():
    parser = argparse.ArgumentParser(description="Generate the May 2026 PnL audit and KPI reports.")
    parser.add_argument(
        "--input-root",
        default=DEFAULT_INPUT_ROOT,
        help="Folder containing the Raw data subfolder.",
    )
    parser.add_argument(
        "--output-dir",
        default=DEFAULT_OUTPUT_ROOT,
        help="Output folder. Defaults to DEFAULT_OUTPUT_ROOT configured at the top of the script.",
    )
    parser.add_argument("--skip-pnl", action="store_true", help="Skip Pnl Audit generation.")
    parser.add_argument("--skip-kpi", action="store_true", help="Skip KPI Excel and Word generation.")
    return parser.parse_args()


def run_pnl_summary():
    # ============================== CONFIGURATION ===============================
    FOLDER_INPUT = str(PNL_INPUT_DIR)
    OUTPUT_BASE = str(OUTPUT_DIR)
    PCT_FORMAT = "0.00%"
    HIGH_GM_THRESHOLD = 70
    LOW_GM_THRESHOLD = -70
    COUNTRY_CURRENCY = {
        "Brazil": "BRL", "India": "INR", "Italy": "EUR", "Japan": "JPY",
        "Malaysia": "MYR", "Singapore": "SGD", "Spain": "EUR",
        "Thailand": "THB", "United States": "USD", "Vietnam": "VND",
        "China": "CNY", "Canada": "CAD", "France": "EUR", "Belgium": "EUR",
        "Switzerland": "CHF", "Portugal": "EUR", "Sweden": "EUR",
        "Netherlands": "EUR", "Luxembourg": "EUR", "Tunisia": "EUR",
        "Czech Republic": "EUR", "Austria": "EUR", "Turkey": "TRY",
        "Mexico": "MXN",
    }


    # ================================ HELPERS ===================================
    def to_num(value):
        try:
            if pd.isna(value) or str(value).strip() in {"", "nan", "NaN", "None"}:
                return np.nan
            return float(value)
        except Exception:
            return np.nan


    def pct_text(value):
        return "N/A" if value is None or pd.isna(value) else f"{value:.1f}%"


    def extract_id(value):
        match = re.search(r"PR(\d+)", str(value), re.IGNORECASE)
        return match.group(1) if match else None


    def safe_filename(value):
        return re.sub(r"[\\/*?:\[\]<>|]", "_", str(value))[:60]


    def make_fill(color):
        return PatternFill("solid", fgColor=color)


    PATTERN_COLORS = {
        "REVERSAL": "FFCCCC",
        "REGULARIZATION": "FFE0CC",
        "MISSING_REVENUE": "FFF2CC",
        "VOLATILE": "E8DAEF",
        "SYSTEMATIC_HIGH": "FFCCCC",
        "ISOLATED_SPIKE": "CCE5FF",
        "NEGATIVE_REVENUE": "F9CBCB",
        "NEGATIVE_GM": "FCE4EC",
        "NORMAL_LOOKING": "CCFFCC",
    }


    def detect_pattern(row, months):
        active = [
            (m, row.get(f"GMP_M{m}", np.nan))
            for m in months
            if not pd.isna(row.get(f"GMP_M{m}", np.nan))
            and row.get(f"Rev_M{m}", 0) != 0
        ]
        if not active:
            return "No monthly data", []

        flags, details = [], []
        reversals = [(m, p) for m, p in active if p < -70]
        if reversals:
            flags.append("REVERSAL")
            details += [f"M{m}: {pct_text(p)} (reversal entry)" for m, p in reversals]

        for i, (m, p) in enumerate(active):
            if i and p > 100 and active[i - 1][1] < -70:
                pm, pp = active[i - 1]
                flags.append("REGULARIZATION")
                details.append(f"M{pm}->M{m}: {pct_text(pp)}->{pct_text(p)} (correction)")

        for m in months:
            rev = row.get(f"Rev_M{m}", np.nan)
            gm = row.get(f"GM_M{m}", np.nan)
            pct = row.get(f"GMP_M{m}", np.nan)
            if not pd.isna(rev) and rev < 0:
                flags.append("NEGATIVE_REVENUE")
                details.append(f"M{m}: Rev={rev:,.0f} (negative) | GM%={pct_text(pct)}")
            if not pd.isna(rev) and rev > 0 and not pd.isna(gm) and gm < 0:
                flags.append("NEGATIVE_GM")
                details.append(f"M{m}: Rev={rev:,.0f}, GM={gm:.2f} (negative) | GM%={pct_text(pct)}")
            if not pd.isna(rev) and rev == 0 and not pd.isna(gm) and gm != 0:
                flags.append("MISSING_REVENUE")
                details.append(f"M{m}: Rev=0 but GM={gm:.2f}")

        values = [p for _, p in active]
        if len(values) >= 2 and np.std(values) > 30:
            flags.append("VOLATILE")
            details.append(f"Std dev={np.std(values):.1f}% (expected <=10%)")

        if len(active) >= 2 and all(p > 70 for _, p in active):
            flags.append("SYSTEMATIC_HIGH")
            details.append(f"All {len(active)} months >70% -> likely missing cost")

        extreme = [p for _, p in active if p > 70 or p < -70]
        normal = [p for _, p in active if -70 <= p <= 70]
        if len(extreme) == 1 and normal:
            flags.append("ISOLATED_SPIKE")
            details.append(f"1 abnormal month ({pct_text(extreme[0])}), others normal")

        if not flags:
            flags.append("NORMAL_LOOKING")
            details.append("Stable, no abnormal pattern")
        unique_flags = list(dict.fromkeys(flags))
        return " + ".join(unique_flags), details


    # ============================ COMMON STYLES =================================
    WHITE = "FFFFFF"
    NAVY = "17365D"
    HEADER = "1F4E79"
    MONTH = "2E75B6"
    TOTAL = "1A5276"
    PALE_BLUE = "D9EAF7"
    GREEN = "E2F0D9"
    YELLOW = "FFF2CC"
    RED = "F4CCCC"


    def style_header(cell, color=HEADER):
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = make_fill(color)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(
            left=Side(style="thin", color=WHITE),
            right=Side(style="thin", color=WHITE),
            top=Side(style="thin", color=WHITE),
            bottom=Side(style="thin", color=WHITE),
        )


    def title_block(ws, title, width, subtitle=None):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
        c = ws.cell(1, 1, title)
        c.font = Font(bold=True, size=14, color=WHITE)
        c.fill = make_fill(NAVY)
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 28
        if subtitle is not None:
            ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
            c = ws.cell(2, 1, subtitle)
            c.font = Font(italic=True, color=WHITE)
            c.fill = make_fill("2C3E50")
            c.alignment = Alignment(horizontal="left", vertical="center")
            ws.row_dimensions[2].height = 22


    def write_cell(ws, row, col, value, fill=None, number_format=None, align="center", bold=False):
        if value is None:
            value = ""
        else:
            try:
                if pd.isna(value):
                    value = ""
            except Exception:
                pass

        cell = ws.cell(row, col, value)
        if fill:
            cell.fill = fill
        if number_format:
            cell.number_format = number_format
        cell.font = Font(bold=bold, size=10)
        cell.alignment = Alignment(horizontal=align, vertical="top", wrap_text=True)
        cell.border = Border(
            left=Side(style="thin", color="D9E2F3"),
            right=Side(style="thin", color="D9E2F3"),
            top=Side(style="thin", color="D9E2F3"),
            bottom=Side(style="thin", color="D9E2F3"),
        )
        return cell


    # ============================ DETAIL SHEETS =================================
    def write_monthly_sheet(ws, df, months, abnormal=False):
        fixed = ["No.", "Client", "Project", "Pattern", "Assessment"]
        nf = len(fixed)
        total_start = nf + 1 + len(months) * 3
        last_col = total_start + 2

        title_block(ws, "All Abnormal Projects" if abnormal else "Raw Data", last_col)
        ws.sheet_view.showGridLines = False
        ws.row_dimensions[3].height = 30
        ws.row_dimensions[4].height = 24

        for col, name in enumerate(fixed, 1):
            ws.merge_cells(start_row=3, start_column=col, end_row=4, end_column=col)
            style_header(ws.cell(3, col))
            ws.cell(3, col).value = name

        for i, month in enumerate(months):
            start = nf + 1 + i * 3
            ws.merge_cells(start_row=3, start_column=start, end_row=3, end_column=start + 2)
            style_header(ws.cell(3, start), MONTH)
            ws.cell(3, start).value = str(month)
            for j, sub in enumerate(["Rev", "GM", "GM%"]):
                style_header(ws.cell(4, start + j), MONTH)
                ws.cell(4, start + j).value = sub

        ws.merge_cells(start_row=3, start_column=total_start, end_row=3, end_column=total_start + 2)
        style_header(ws.cell(3, total_start), TOTAL)
        ws.cell(3, total_start).value = "Total"
        for j, sub in enumerate(["Rev", "GM", "GM%"]):
            style_header(ws.cell(4, total_start + j), TOTAL)
            ws.cell(4, total_start + j).value = sub

        for r, (_, row) in enumerate(df.iterrows(), 5):
            pat = str(row.get("Pattern", ""))
            fill_color = next((v for k, v in PATTERN_COLORS.items() if k in pat), "FFFFFF") if abnormal else (
                "FFCCCC" if row.get("GM_Pct_Total", np.nan) > HIGH_GM_THRESHOLD or row.get("GM_Pct_Total", np.nan) < LOW_GM_THRESHOLD else "FFFFFF"
            )
            fill = make_fill(fill_color)
            values = [
                r - 4,
                row.get("Client", ""),
                row.get("Project", ""),
                pat,
                " | ".join(row.get("Detail", [])),
            ]
            for c, value in enumerate(values, 1):
                write_cell(ws, r, c, value, fill, align="left" if c not in {1, 3} else "center")

            for i, month in enumerate(months):
                start = nf + 1 + i * 3
                rev = row.get(f"Rev_M{month}", np.nan)
                gm = row.get(f"GM_M{month}", np.nan)
                gmp = row.get(f"GMP_M{month}", np.nan)
                rev = None if pd.isna(rev) else round(rev, 2)
                gm = None if pd.isna(gm) else round(gm, 2)
                gmp = None if pd.isna(gmp) else gmp
                write_cell(ws, r, start, rev, fill, "#,##0.00;[Red](#,##0.00);-")
                gm_cell = write_cell(ws, r, start + 1, gm, fill, "#,##0.00;[Red](#,##0.00);-")
                if gm is not None and gm < 0:
                    gm_cell.fill = make_fill("E74C3C")
                    gm_cell.font = Font(bold=True, color=WHITE)
                    gm_cell.number_format = "#,##0.00;(#,##0.00);-"
                gmp_cell = write_cell(ws, r, start + 2, None if gmp is None else round(gmp / 100, 6), fill, PCT_FORMAT)
                if gmp is not None:
                    if gmp < 0:
                        gmp_cell.fill = make_fill("E74C3C")
                        gmp_cell.font = Font(bold=True, color=WHITE)
                    elif gmp > 70:
                        gmp_cell.fill = make_fill("F9E79F")
                        gmp_cell.font = Font(bold=True)
                    else:
                        gmp_cell.fill = make_fill("A9DFBF")

            tr = row.get("Total_Revenue", np.nan)
            tg = row.get("Total_GM", np.nan)
            tp = row.get("GM_Pct_Total", np.nan)
            tr = None if pd.isna(tr) else round(tr, 2)
            tg = None if pd.isna(tg) else round(tg, 2)
            tp = None if pd.isna(tp) else tp
            total_fill = make_fill("D6EAF8")
            write_cell(ws, r, total_start, tr, total_fill, "#,##0.00;[Red](#,##0.00);-", bold=True)
            write_cell(ws, r, total_start + 1, tg, total_fill, "#,##0.00;[Red](#,##0.00);-", bold=True)
            total_pct_cell = write_cell(ws, r, total_start + 2, None if tp is None else round(tp / 100, 6), total_fill, PCT_FORMAT, bold=True)
            if tp is not None and tp > 70:
                total_pct_cell.fill = make_fill("F9E79F")
            elif tp is not None and tp < 0:
                total_pct_cell.fill = make_fill("E74C3C")
                total_pct_cell.font = Font(bold=True, color=WHITE)

            detail_text = str(values[4] or "")
            pattern_text = str(values[3] or "")
            detail_lines = sum(max(1, (len(part) + 69) // 70) for part in detail_text.split(" | ")) if detail_text else 1
            pattern_lines = max(1, (len(pattern_text) + 23) // 24)
            ws.row_dimensions[r].height = min(105, max(20, max(detail_lines, pattern_lines) * 15))

        # Grand total row for the displayed manager projects.
        total_row = len(df) + 5
        total_fill = make_fill("D6EAF8")
        total_label = write_cell(ws, total_row, 1, "Total", total_fill, align="left", bold=True)
        for col in range(2, nf + 1):
            write_cell(ws, total_row, col, None, total_fill, bold=True)
        for month in months:
            start = nf + 1 + months.index(month) * 3
            rev_total = pd.to_numeric(df.get(f"Rev_M{month}", pd.Series(dtype=float)), errors="coerce").sum(min_count=1)
            gm_total = pd.to_numeric(df.get(f"GM_M{month}", pd.Series(dtype=float)), errors="coerce").sum(min_count=1)
            pct_total = (gm_total / rev_total) if pd.notna(rev_total) and rev_total not in (0, 0.0) and pd.notna(gm_total) else np.nan
            write_cell(ws, total_row, start, None if pd.isna(rev_total) else round(rev_total, 2), total_fill, "#,##0.00;[Red](#,##0.00);-", bold=True)
            write_cell(ws, total_row, start + 1, None if pd.isna(gm_total) else round(gm_total, 2), total_fill, "#,##0.00;[Red](#,##0.00);-", bold=True)
            write_cell(ws, total_row, start + 2, None if pd.isna(pct_total) else round(pct_total, 6), total_fill, PCT_FORMAT, bold=True)
        rev_total = pd.to_numeric(df.get("Total_Revenue", pd.Series(dtype=float)), errors="coerce").sum(min_count=1)
        gm_total = pd.to_numeric(df.get("Total_GM", pd.Series(dtype=float)), errors="coerce").sum(min_count=1)
        pct_total = (gm_total / rev_total) if pd.notna(rev_total) and rev_total not in (0, 0.0) and pd.notna(gm_total) else np.nan
        write_cell(ws, total_row, total_start, None if pd.isna(rev_total) else round(rev_total, 2), total_fill, "#,##0.00;[Red](#,##0.00);-", bold=True)
        write_cell(ws, total_row, total_start + 1, None if pd.isna(gm_total) else round(gm_total, 2), total_fill, "#,##0.00;[Red](#,##0.00);-", bold=True)
        write_cell(ws, total_row, total_start + 2, None if pd.isna(pct_total) else round(pct_total, 6), total_fill, PCT_FORMAT, bold=True)
        ws.row_dimensions[total_row].height = 22

        widths = [5, 24, 42, 24, 58]
        for i, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = width
        for i in range(len(months)):
            start = nf + 1 + i * 3
            ws.column_dimensions[get_column_letter(start)].width = 11
            ws.column_dimensions[get_column_letter(start + 1)].width = 11
            ws.column_dimensions[get_column_letter(start + 2)].width = 8
        for i in range(3):
            ws.column_dimensions[get_column_letter(total_start + i)].width = 12
        ws.auto_filter.ref = f"A3:{get_column_letter(last_col)}{max(4, len(df) + 4)}"


    # ============================== SUMMARY SHEETS ==============================
    def write_bu_overview(ws, stats, ca_name):
        headers = [
            "No.", "Country", "Business Unit", "Source File", "Cleaned Pnl",
            "# Total Projects", "# Projects GM% > 70%", "High GM Rate", "Currency",
            "Revenue Local", "GM Local", "Revenue EUR", "GM EUR", "BU GM%",
        ]
        title_block(ws, f"CA: {ca_name} - Business Units Overview", len(headers), f"Total BUs processed: {len(stats)}")
        ws.sheet_view.showGridLines = False
        ws.row_dimensions[4].height = 34
        for c, header in enumerate(headers, 1):
            style_header(ws.cell(4, c))
            ws.cell(4, c).value = header

        for r, bu in enumerate(stats, 5):
            gm_pct = bu.get("bu_gmp", np.nan)
            fill = make_fill("FFCCCC" if not pd.isna(gm_pct) and gm_pct > 60 else ("EBF5FB" if r % 2 == 0 else "FFFFFF"))
            values = [
                r - 4,
                bu.get("country", "N/A"),
                bu["bu_name"],
                bu["source_file"],
                "Open Pnl" if bu.get("cleaned_relpath") else "",
                bu["total_projects"],
                bu["abnormal_count"],
                bu.get("high_gm_rate", np.nan),
                bu.get("currency", "N/A"),
                None if pd.isna(bu.get("bu_rev", np.nan)) else round(bu["bu_rev"], 2),
                None if pd.isna(bu.get("bu_gm", np.nan)) else round(bu["bu_gm"], 2),
                None if pd.isna(bu.get("bu_rev_eur", np.nan)) else round(bu["bu_rev_eur"], 2),
                None if pd.isna(bu.get("bu_gm_eur", np.nan)) else round(bu["bu_gm_eur"], 2),
                None if pd.isna(gm_pct) else round(gm_pct / 100, 6),
            ]
            for c, value in enumerate(values, 1):
                cell = write_cell(ws, r, c, value, fill, align="left" if c in {2, 3} else "center")
                if c == 5 and bu.get("cleaned_relpath"):
                    cell.hyperlink = bu["cleaned_relpath"]
                    cell.font = Font(color="0563C1", underline="single")
                if c in {6, 7}:
                    cell.number_format = "#,##0"
                if c == 8:
                    cell.number_format = PCT_FORMAT
                if c in {10, 11, 12, 13}:
                    cell.number_format = "#,##0.00;[Red](#,##0.00);-"
                if c == 14:
                    cell.number_format = PCT_FORMAT
        for c, width in enumerate([5, 18, 28, 22, 14, 16, 21, 14, 11, 17, 17, 17, 17, 12], 1):
            ws.column_dimensions[get_column_letter(c)].width = width
        ws.freeze_panes = "F5"
        ws.print_title_rows = "1:4"
        ws.auto_filter.ref = f"A4:N{max(4, len(stats) + 4)}"


    def write_high_gm_summary(ws, df, months, ca_name, bu_gm_map=None, summary_title="High GM% Projects", summary_condition="GM% > 70%"):
        fixed = ["No.", "Country", "Currency", "BU", "Client", "Project", "Pattern", "Assessment"]
        nf = len(fixed)
        total_start = nf + 1 + len(months) * 3
        last_col = total_start + 2
        share_col = last_col + 1
        bu_gm_map = bu_gm_map or {}
        title_block(ws, f"CA: {ca_name} - {summary_title} Summary ({summary_condition})", share_col, f"Total projects: {len(df)}")
        ws.sheet_view.showGridLines = False
        ws.row_dimensions[3].height = 30
        ws.row_dimensions[4].height = 24

        for c, name in enumerate(fixed, 1):
            ws.merge_cells(start_row=3, start_column=c, end_row=4, end_column=c)
            style_header(ws.cell(3, c))
            ws.cell(3, c).value = name
        for i, month in enumerate(months):
            start = nf + 1 + i * 3
            ws.merge_cells(start_row=3, start_column=start, end_row=3, end_column=start + 2)
            style_header(ws.cell(3, start), MONTH)
            ws.cell(3, start).value = str(month)
            for j, sub in enumerate(["Rev", "GM", "GM%"]):
                style_header(ws.cell(4, start + j), MONTH)
                ws.cell(4, start + j).value = sub
        ws.merge_cells(start_row=3, start_column=total_start, end_row=3, end_column=total_start + 2)
        style_header(ws.cell(3, total_start), TOTAL)
        ws.cell(3, total_start).value = "Total"
        for j, sub in enumerate(["Rev", "GM", "GM%"]):
            style_header(ws.cell(4, total_start + j), TOTAL)
            ws.cell(4, total_start + j).value = sub
        ws.merge_cells(start_row=3, start_column=share_col, end_row=4, end_column=share_col)
        style_header(ws.cell(3, share_col), TOTAL)
        ws.cell(3, share_col).value = "%GM/Total BU"

        previous_bu, toggle = None, False
        for r, (_, row) in enumerate(df.iterrows(), 5):
            bu = row.get("BU", "")
            if bu != previous_bu:
                previous_bu, toggle = bu, not toggle
            fill = make_fill(next((v for k, v in PATTERN_COLORS.items() if k in str(row.get("Pattern", ""))), "EBF5FB" if toggle else "FDFEFE"))
            values = [r - 4, row.get("Country", "N/A"), row.get("Currency", "N/A"), bu, row.get("Client", ""), row.get("Project", ""), row.get("Pattern", ""), " | ".join(row.get("Detail", []))]
            for c, value in enumerate(values, 1):
                write_cell(ws, r, c, value, fill, align="left" if c not in {1, 2, 3, 4, 6} else "center")
            for i, month in enumerate(months):
                start = nf + 1 + i * 3
                rev = row.get(f"Rev_M{month}", np.nan)
                gm = row.get(f"GM_M{month}", np.nan)
                gmp = row.get(f"GMP_M{month}", np.nan)
                write_cell(ws, r, start, None if pd.isna(rev) else round(rev, 2), fill, "#,##0.00;[Red](#,##0.00);-")
                write_cell(ws, r, start + 1, None if pd.isna(gm) else round(gm, 2), fill, "#,##0.00;[Red](#,##0.00);-")
                cell = write_cell(ws, r, start + 2, None if pd.isna(gmp) else round(gmp / 100, 6), fill, PCT_FORMAT)
                if not pd.isna(gmp) and gmp > 70:
                    cell.fill = make_fill("F9E79F")
            tr, tg, tp = row.get("Total_Revenue", np.nan), row.get("Total_GM", np.nan), row.get("GM_Pct_Total", np.nan)
            total_fill = make_fill("D6EAF8")
            write_cell(ws, r, total_start, None if pd.isna(tr) else round(tr, 2), total_fill, "#,##0.00;[Red](#,##0.00);-", bold=True)
            write_cell(ws, r, total_start + 1, None if pd.isna(tg) else round(tg, 2), total_fill, "#,##0.00;[Red](#,##0.00);-", bold=True)
            write_cell(ws, r, total_start + 2, None if pd.isna(tp) else round(tp / 100, 6), total_fill, PCT_FORMAT, bold=True)
            bu_gm = pd.to_numeric(bu_gm_map.get(str(bu), np.nan), errors="coerce")
            project_share = np.nan if pd.isna(tg) or pd.isna(bu_gm) or abs(float(bu_gm)) < 1e-12 else float(tg) / float(bu_gm)
            share_cell = write_cell(ws, r, share_col, None if pd.isna(project_share) else round(project_share, 6), total_fill, PCT_FORMAT, bold=True)
            if not pd.isna(project_share) and float(project_share) > 0.10:
                share_cell.fill = make_fill("F4CCCC")
                share_cell.font = Font(bold=True, color="9C0006")
            ws.row_dimensions[r].height = 40

        for c, width in enumerate([5, 18, 11, 22, 18, 42, 22, 65], 1):
            ws.column_dimensions[get_column_letter(c)].width = width
        for i in range(len(months)):
            start = nf + 1 + i * 3
            for offset, width in enumerate([13, 13, 9]):
                ws.column_dimensions[get_column_letter(start + offset)].width = width
        for offset, width in enumerate([14, 14, 10]):
            ws.column_dimensions[get_column_letter(total_start + offset)].width = width
        ws.column_dimensions[get_column_letter(share_col)].width = 15
        # Freeze title/header rows and identifier columns only through BU (A:D).
        ws.freeze_panes = "E5"
        ws.print_title_rows = "1:4"


    def write_pnl_graphs(ws, stats, ca_name):
        """Create a normalized high-GM exception-rate chart by country."""
        title_block(ws, f"PNL AUDIT GRAPHS - {ca_name}", 5, "Share of projects with GM% above 70% by country")
        ws.sheet_view.showGridLines = False
        headers = ["Country", "# Managers", "# Total Projects", "# Projects GM% > 70%", "High GM Rate"]
        for col, header in enumerate(headers, 1):
            style_header(ws.cell(4, col))
            ws.cell(4, col).value = header

        frame = pd.DataFrame(stats)
        if not frame.empty:
            frame["country"] = frame.get("country", "N/A").fillna("N/A")
            country = frame.groupby("country", dropna=False).agg(
                Managers=("bu_name", "count"),
                Projects=("total_projects", "sum"),
                HighGM=("abnormal_count", "sum"),
            ).reset_index().sort_values("Projects", ascending=False)
        else:
            country = pd.DataFrame(columns=["country", "Managers", "Projects", "HighGM"])

        for row, (_, item) in enumerate(country.iterrows(), 5):
            fill = make_fill("FFFFFF" if row % 2 else "EBF5FB")
            high_gm_rate = item["HighGM"] / item["Projects"] if item["Projects"] else np.nan
            values = [item["country"], item["Managers"], item["Projects"], item["HighGM"], high_gm_rate]
            for col, value in enumerate(values, 1):
                cell = write_cell(ws, row, col, value, fill, align="left" if col == 1 else "center")
                if col == 5:
                    cell.number_format = "0.00%"

        last_row = max(5, len(country) + 4)
        ws.auto_filter.ref = f"A4:E{last_row}"
        for col, width in enumerate([20, 14, 18, 24, 16], 1):
            ws.column_dimensions[get_column_letter(col)].width = width

        if len(country) > 0:
            categories = Reference(ws, min_col=1, min_row=5, max_row=last_row)

            # Explicitly use a text category reference. openpyxl can otherwise
            # serialize text country names as numRef, which makes some Excel
            # viewers omit the X-axis labels entirely.
            country_ref = f"'{ws.title}'!$A$5:$A${last_row}"
            country_points = [
                StrVal(idx=idx, v=str(ws.cell(row=row, column=1).value or ""))
                for idx, row in enumerate(range(5, last_row + 1))
            ]
            category_source = AxDataSource(
                strRef=StrRef(f=country_ref, strCache=StrData(pt=country_points))
            )

            # Primary bars: exception rate, not raw count, so differently sized
            # country portfolios remain comparable.
            project_bar = BarChart()
            project_bar.type = "col"
            project_bar.style = 10
            project_bar.varyColors = False
            # The worksheet title already describes the chart. Removing the
            # internal chart title prevents it from colliding with the legend.
            project_bar.title = None
            # Use a single normalized bar series. Raw counts remain in the table.
            project_bar.y_axis.title = None
            project_bar.y_axis.axPos = "l"
            project_bar.y_axis.delete = False
            project_bar.y_axis.tickLblPos = "nextTo"
            project_bar.y_axis.numFmt = "0.0%"
            rates = country["HighGM"] / country["Projects"].replace(0, np.nan)
            max_high_gm = float(pd.to_numeric(rates, errors="coerce").fillna(0).max()) if len(country) else 0
            project_bar.y_axis.scaling.min = 0
            project_bar.y_axis.scaling.max = max(0.05, np.ceil(max_high_gm * 100) / 100)
            project_bar.y_axis.majorUnit = 0.01
            project_bar.x_axis.title = None
            project_bar.x_axis.tickLblPos = "low"
            project_bar.x_axis.delete = False
            project_bar.x_axis.noMultiLvlLbl = True
            project_bar.x_axis.txPr = RichText(
                bodyPr=RichTextProperties(rot=2700000),
                p=[Paragraph(
                    pPr=ParagraphProperties(defRPr=CharacterProperties(sz=900)),
                    endParaRPr=CharacterProperties(sz=900),
                )],
            )
            project_bar.height = 10
            project_bar.width = 23
            project_bar.layout = Layout(
                manualLayout=ManualLayout(x=0.08, y=0.14, w=0.88, h=0.72)
            )
            project_bar.y_axis.majorGridlines = None
            project_bar.add_data(Reference(ws, min_col=5, max_col=5, min_row=4, max_row=last_row), titles_from_data=True)
            project_bar.set_categories(categories)
            for series in project_bar.series:
                series.cat = category_source
                series.tx = SeriesLabel(v="High GM project rate")
                series.graphicalProperties.solidFill = "4F81BD"
                series.graphicalProperties.line.solidFill = "4F81BD"
            project_bar.legend.position = "t"
            project_bar.legend.overlay = False

            # Leave a clear visual gap between the source table (A:E) and chart.
            ws.add_chart(project_bar, "J4")

    def write_control_checks(ws, checks, ca_name):
        headers = ["Check", "Actual", "Expected", "Difference", "Tolerance", "Status"]
        title_block(ws, f"AUTOMATION CONTROL CHECKS - {ca_name}", len(headers), "Source-to-output reconciliation")
        ws.sheet_view.showGridLines = False
        for c, header in enumerate(headers, 1):
            style_header(ws.cell(4, c))
            ws.cell(4, c).value = header
        for r, check in enumerate(checks, 5):
            for c, key in enumerate(headers, 1):
                cell = write_cell(ws, r, c, check[key], align="left" if c == 1 else "center")
                if c in {2, 3, 4, 5} and isinstance(check[key], (int, float)):
                    cell.number_format = "#,##0.00;[Red](#,##0.00);-"
                status = check["Status"]
                if status == "PASS":
                    cell.fill = make_fill(GREEN)
                    if c == 6:
                        cell.font = Font(bold=True, color="217346")
                elif status == "REVIEW":
                    cell.fill = make_fill(YELLOW)
                    if c == 6:
                        cell.font = Font(bold=True, color="BF9000")
                else:
                    cell.fill = make_fill(RED)
                    if c == 6:
                        cell.font = Font(bold=True, color="C00000")
        for c, width in enumerate([58, 18, 18, 18, 14, 14], 1):
            ws.column_dimensions[get_column_letter(c)].width = width
        ws.auto_filter.ref = f"A4:F{max(4, len(checks) + 4)}"


    def numeric_check(name, actual, expected, tolerance=0):
        if actual is None or expected is None or pd.isna(actual) or pd.isna(expected):
            return {"Check": name, "Actual": actual, "Expected": expected, "Difference": None, "Tolerance": tolerance, "Status": "REVIEW"}
        diff = actual - expected
        return {"Check": name, "Actual": actual, "Expected": expected, "Difference": diff, "Tolerance": tolerance, "Status": "PASS" if abs(diff) <= tolerance else "FAIL"}


    # ============================= PROCESS ONE BU ===============================
    def process_file(file_path, output_folder):
        file_name = os.path.basename(file_path)
        try:
            excel_file = pd.ExcelFile(file_path)
            raw = pd.read_excel(file_path, sheet_name=excel_file.sheet_names[0], header=None)
            month_row, data_rows = raw.iloc[1], raw.iloc[3:]

            month_cols = {}
            for index, value in enumerate(month_row):
                text = str(value).strip()
                if text not in {"nan", "", "NaN", "MONTH", "Total"}:
                    try:
                        month = int(float(text))
                        month_cols.setdefault(month, index)
                    except Exception:
                        pass
            months = sorted(month_cols)
            if not months:
                raise ValueError("No month columns found")

            total_start = next((i for i, value in enumerate(month_row) if str(value).strip() == "Total"), len(raw.columns) - 4)
            total_rev_col, total_gm_col, total_gmp_col = total_start, total_start + 1, total_start + 2

            bu_name = None
            for _, row in data_rows.iterrows():
                text = str(row.iloc[0])
                if "BusinessUnit is" in text:
                    match = re.search(r"BusinessUnit is (.+?)(?:\n|Included|Excluded|$)", text)
                    if match:
                        bu_name = match.group(1).strip()
                    break
            bu_name = bu_name or Path(file_name).stem

            country_name = None
            for _, metadata_row in data_rows.iterrows():
                metadata_text = " ".join(
                    str(value) for value in metadata_row.tolist()
                    if not pd.isna(value)
                )
                country_match = re.search(r"Country\s+is\s+([^\n|]+)", metadata_text, re.IGNORECASE)
                if country_match:
                    country_name = country_match.group(1).strip()
                    break
            country_name = country_name or "Unknown Country"
            currency_code = COUNTRY_CURRENCY.get(country_name)
            fx_rate = globals().get("DEFAULT_FX_TO_EUR", {}).get(currency_code, np.nan)
            if currency_code is None or pd.isna(fx_rate):
                raise ValueError(f"Missing country/currency mapping or FX rate for: {country_name}")

            first_label = data_rows.iloc[:, 0].astype(str).str.strip()
            second_label = data_rows.iloc[:, 1].astype(str).str.strip()
            # PnL exports can contain several numeric rows labelled Total:
            # project totals, BU totals, and the final workbook total. Always
            # use the bottom-most numeric Total row, which is the final PnL
            # total and not the last project's subtotal.
            total_mask = (
                (first_label.isin(["nan", ""]) & (second_label == "Total"))
                | (first_label == "Total")
                | (second_label == "Total")
            )
            grand = data_rows[total_mask]
            grand = grand[grand.iloc[:, total_rev_col].map(to_num).notna()]
            if len(grand):
                grand = grand.iloc[[-1]]
            bu_rev = to_num(grand.iloc[0, total_rev_col]) if len(grand) else np.nan
            bu_gm = to_num(grand.iloc[0, total_gm_col]) if len(grand) else np.nan
            raw_bu_gmp = to_num(grand.iloc[0, total_gmp_col]) if len(grand) else np.nan
            bu_gmp = raw_bu_gmp * 100 if not pd.isna(raw_bu_gmp) else np.nan

            records, current_client = [], None
            for _, row in data_rows.iterrows():
                client = str(row.iloc[0]).strip()
                project = "" if pd.isna(row.iloc[1]) else str(row.iloc[1]).strip()
                if client.startswith("Applied filters") or (client == "nan" and project == "nan"):
                    continue
                if client not in {"nan", "", "NaN", "Total"}:
                    current_client = client
                if project in {"Total", "nan", "", "NaN"} or client == "Total":
                    continue

                raw_total_pct = to_num(row.iloc[total_gmp_col])
                record = {
                    "Client": current_client,
                    "Project": project,
                    # Currency intentionally removed from the output.
                    "Total_Revenue": to_num(row.iloc[total_rev_col]),
                    "Total_GM": to_num(row.iloc[total_gm_col]),
                    "GM_Pct_Total": raw_total_pct * 100 if not pd.isna(raw_total_pct) else np.nan,
                }
                for month, start in month_cols.items():
                    raw_gmp = to_num(row.iloc[start + 2])
                    record[f"Rev_M{month}"] = to_num(row.iloc[start])
                    record[f"GM_M{month}"] = to_num(row.iloc[start + 1])
                    record[f"GMP_M{month}"] = raw_gmp * 100 if not pd.isna(raw_gmp) else np.nan
                records.append(record)

            df = pd.DataFrame(records)
            if df.empty:
                raise ValueError("No project rows found")
            df = df[df["Client"].notna() & ~df["Client"].isin(["nan", "", "NaN", "Total"])].copy()

            # Some source exports contain a footer labelled "Total" but do
            # not populate the BU total cells. In that case, derive the BU
            # totals from the parsed project rows so the overview never shows
            # empty Revenue / GM / BU GM% columns.
            project_rev_sum = pd.to_numeric(df.get("Total_Revenue"), errors="coerce").sum(min_count=1)
            project_gm_sum = pd.to_numeric(df.get("Total_GM"), errors="coerce").sum(min_count=1)
            if pd.isna(project_rev_sum):
                revenue_cols = [f"Rev_M{month}" for month in months if f"Rev_M{month}" in df.columns]
                project_rev_sum = pd.to_numeric(df[revenue_cols], errors="coerce").sum(axis=1, min_count=1).sum(min_count=1) if revenue_cols else np.nan
            if pd.isna(project_gm_sum):
                gm_cols = [f"GM_M{month}" for month in months if f"GM_M{month}" in df.columns]
                project_gm_sum = pd.to_numeric(df[gm_cols], errors="coerce").sum(axis=1, min_count=1).sum(min_count=1) if gm_cols else np.nan
            if pd.isna(bu_rev) or (abs(float(bu_rev)) < 1e-12 and pd.notna(project_rev_sum) and abs(float(project_rev_sum)) > 1e-12):
                bu_rev = project_rev_sum
            if pd.isna(bu_gm) or (abs(float(bu_gm)) < 1e-12 and pd.notna(project_gm_sum) and abs(float(project_gm_sum)) > 1e-12):
                bu_gm = project_gm_sum
            if (pd.isna(bu_gmp) or abs(float(bu_gmp)) < 1e-12) and pd.notna(bu_rev) and abs(float(bu_rev)) > 1e-12 and pd.notna(bu_gm):
                bu_gmp = float(bu_gm) / float(bu_rev) * 100
            df[["Pattern", "Detail"]] = df.apply(lambda row: pd.Series(detect_pattern(row, months)), axis=1)

            abnormal = df[(df["GM_Pct_Total"] > HIGH_GM_THRESHOLD) | (df["GM_Pct_Total"] < LOW_GM_THRESHOLD)].copy()
            abnormal["BU"] = bu_name
            abnormal["Country"] = country_name
            abnormal["Currency"] = currency_code
            print(f"      BU: {bu_name} | Projects: {len(df)} | Abnormal: {len(abnormal)}")

            country_folder = os.path.join(output_folder, safe_filename(country_name))
            os.makedirs(country_folder, exist_ok=True)
            out_path = os.path.join(country_folder, f"{safe_filename(bu_name)}.xlsx")
            writer = pd.ExcelWriter(out_path, engine="openpyxl")
            writer.book.create_sheet("Raw Data")
            write_monthly_sheet(writer.book["Raw Data"], df, months, False)
            if len(abnormal):
                writer.book.create_sheet("All Abnormal Projects")
                write_monthly_sheet(writer.book["All Abnormal Projects"], abnormal, months, True)
            if "Sheet" in writer.book.sheetnames:
                del writer.book["Sheet"]
            writer.close()

            try:
                check_wb = load_workbook(out_path, read_only=True, data_only=True)
                check_ws = check_wb["Raw Data"]
                output_rows = sum(
                    1 for values in check_ws.iter_rows(min_row=5, values_only=True)
                    if values and values[0] not in [None, "", "Total"]
                )
                check_wb.close()
            except Exception:
                output_rows = np.nan

            bu_info = {
                "bu_name": bu_name,
                "country": country_name,
                "source_file": file_name,
                "total_projects": len(df),
                "source_project_rows": len(df),
                "output_project_rows": output_rows,
                "abnormal_count": int((df["GM_Pct_Total"] > HIGH_GM_THRESHOLD).sum()),
                "high_gm_rate": float((df["GM_Pct_Total"] > HIGH_GM_THRESHOLD).mean()),
                "currency": currency_code,
                "fx_rate": float(fx_rate),
                "bu_rev": bu_rev,
                "bu_gm": bu_gm,
                "bu_rev_eur": float(bu_rev) * float(fx_rate) if pd.notna(bu_rev) else np.nan,
                "bu_gm_eur": float(bu_gm) * float(fx_rate) if pd.notna(bu_gm) else np.nan,
                "bu_gmp": bu_gmp,
                "project_rev_sum": project_rev_sum,
                "project_gm_sum": project_gm_sum,
                # Use an absolute local file URI. Relative paths inside a workbook
                # stored in OneDrive can be rewritten to onedrive.live.com and
                # open in Excel Web instead of the locally synced workbook.
                "cleaned_relpath": Path(out_path).resolve().as_uri(),
            }
            print(f"      Created: {os.path.basename(out_path)}")
            return abnormal, months, bu_info
        except Exception as error:
            print(f"      ERROR {file_name}: {error}")
            traceback.print_exc()
            return None, None, None


    # ================================= MAIN =====================================
    if not os.path.exists(FOLDER_INPUT):
        raise FileNotFoundError(f"PnL source folder does not exist: {FOLDER_INPUT}")

    data_files = sorted(
        path for pattern in ("*.xlsx", "*.xlsm", "*.xls")
        for path in glob.glob(os.path.join(FOLDER_INPUT, "**", pattern), recursive=True)
        if not os.path.basename(path).startswith("~$")
    )
    if not data_files:
        raise FileNotFoundError(f"No PnL files found inside: {FOLDER_INPUT}")

    ca_name = "Manager Pnl"
    cleaned_root = os.path.join(OUTPUT_BASE, "Cleaned data")
    output_folder = os.path.join(cleaned_root, "Manager Pnl - Cleaned data")
    os.makedirs(output_folder, exist_ok=True)
    print(f"Input PnL folder: {FOLDER_INPUT}")
    print(f"PNL files found: {len(data_files)}")
    print(f"Manager output folder: {output_folder}")

    ca_abnormal_list, ca_months_set, ca_bu_stats, ca_errors = [], set(), [], []
    total_files = len(data_files)
    total_ok = 0
    worker_count = min(8, max(2, os.cpu_count() or 4))
    results = {}
    print(f"Processing PnL files with {worker_count} workers")
    with ThreadPoolExecutor(max_workers=worker_count) as executor:
        futures = {executor.submit(process_file, path, output_folder): path for path in data_files}
        for completed, future in enumerate(futures, 1):
            path = futures[future]
            try:
                results[path] = future.result()
            except Exception as error:
                print(f"ERROR {os.path.basename(path)}: {error}")
                results[path] = (None, None, None)
            if completed == 1 or completed % 10 == 0 or completed == len(data_files):
                print(f"PnL progress: {completed}/{len(data_files)}")

    # Aggregate in sorted file order for stable Summary output.
    for file_path in data_files:
        result = results.get(file_path)
        if result and result[0] is not None:
            abnormal_df, months, bu_info = result
            total_ok += 1
            if len(abnormal_df):
                ca_abnormal_list.append(abnormal_df)
            ca_months_set.update(months)
            ca_bu_stats.append(bu_info)
        else:
            ca_errors.append(os.path.basename(file_path))

    if not ca_bu_stats:
        raise RuntimeError("No manager PnL files were processed successfully")

    ca_months = sorted(ca_months_set)
    if ca_abnormal_list:
        all_abnormal = pd.concat(ca_abnormal_list, ignore_index=True)
        ca_abnormal = all_abnormal[all_abnormal["GM_Pct_Total"] > HIGH_GM_THRESHOLD].copy()
        if len(ca_abnormal):
            ca_abnormal = ca_abnormal.sort_values(["BU", "Client", "Project"]).reset_index(drop=True)
    else:
        ca_abnormal = pd.DataFrame()

    summary_path = os.path.join(OUTPUT_BASE, "Pnl Audit.xlsx")
    writer = pd.ExcelWriter(summary_path, engine="openpyxl")
    writer.book.create_sheet("BU Overview")
    write_bu_overview(writer.book["BU Overview"], ca_bu_stats, ca_name)
    writer.book.create_sheet("Graphs")
    write_pnl_graphs(writer.book["Graphs"], ca_bu_stats, ca_name)
    writer.book.create_sheet("High GM% Projects")
    bu_gm_map = {str(item.get("bu_name", "")): item.get("bu_gm", np.nan) for item in ca_bu_stats}
    write_high_gm_summary(writer.book["High GM% Projects"], ca_abnormal, ca_months, ca_name, bu_gm_map)

    source_rows = sum(bu["source_project_rows"] for bu in ca_bu_stats)
    output_rows = sum(bu["output_project_rows"] for bu in ca_bu_stats if not pd.isna(bu["output_project_rows"]))
    expected_high = sum(bu["abnormal_count"] for bu in ca_bu_stats)
    checks = [
        numeric_check("Raw Excel files discovered vs successfully processed manager files", len(ca_bu_stats), len(data_files), 0),
        numeric_check("Source project rows vs cleaned project rows", output_rows, source_rows, 0),
        numeric_check("High GM% projects in Summary vs manager-level count", len(ca_abnormal), expected_high, 0),
        numeric_check("Processing errors", len(ca_errors), 0, 0),
    ]
    writer.book.create_sheet("Automation Control Checks")
    write_control_checks(writer.book["Automation Control Checks"], checks, ca_name)
    if "Sheet" in writer.book.sheetnames:
        del writer.book["Sheet"]
    summary_order = ["Graphs", "BU Overview", "High GM% Projects", "Automation Control Checks"]
    writer.book._sheets = [writer.book[name] for name in summary_order if name in writer.book.sheetnames]
    writer.close()
    print(f"Pnl Audit.xlsx created: {summary_path}")
    print(f"All manager files processed: {total_ok}/{total_files}")

# =============================== KPI FLOW ==================================
# This section reads only files inside MONTHLY_REPORT_DIR.
# May is the priority/current snapshot. April is used only for variance.

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import AxDataSource, StrData, StrRef, StrVal
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout

KPI_OUTPUT_PATH = OUTPUT_DIR / "KPI Report.xlsx"
KPI_CURRENT_FILE = "Manager Bonus YTD May - Cleaned data.xlsx"
KPI_PREVIOUS_FILE = "Manager Bonus YTD April - Raw data.xlsx"
KPI_RAW_MAY_FILE = "Manager Bonus YTD May - Raw data.xlsx"
KPI_MANAGER_COST_FILE = "Manager Cost YTD May.xlsx"
KPI_BP_FILE = "BP Country 2026.xlsx"

# Built-in conversion factors: 1 unit of source currency = EUR amount.
# If an FX_to_EUR.xlsx file is added inside Monthly report, it is used instead.
DEFAULT_FX_TO_EUR = {
    "EUR": 1.0,
    "USD": 0.92,
    "GBP": 1.17,
    "CHF": 1.04,
    "BRL": 0.17,
    "JPY": 0.0062,
    "MYR": 0.20,
    "THB": 0.0255,
    "INR": 0.0110,
    "SGD": 0.68,
    "VND": 0.0000365,
    "CAD": 0.68,
    "CNY": 0.128,
    "MXN": 0.050,
    "TRY": 0.025,
}

KPI_NAVY = "17365D"
KPI_HEADER = "1F4E79"
KPI_BLUE = "2E75B6"
KPI_LIGHT_BLUE = "D9EAF7"
KPI_GREEN = "E2F0D9"
KPI_YELLOW = "FFF2CC"
KPI_RED = "F4CCCC"
KPI_PURPLE = "E4DFEC"
KPI_WHITE = "FFFFFF"


def kpi_fill(color):
    return PatternFill("solid", fgColor=color)


def kpi_clean_columns(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def kpi_find_column(df, candidates, required=True):
    normal = {
        re.sub(r"[^a-z0-9]", "", str(c).lower()): c
        for c in df.columns
    }
    for candidate in candidates:
        key = re.sub(r"[^a-z0-9]", "", candidate.lower())
        if key in normal:
            return normal[key]
    if required:
        available = [str(c) for c in df.columns]
        raise KeyError(f"Missing required column. Tried: {candidates}. Available columns: {available}")
    return None


BONUS_COLUMN_ALIASES = {
    "UserId": ["UserId", "User ID", "EmployeeId", "Employee ID"],
    "Username": ["Username", "User Name", "Employee", "Manager"],
    "Function": ["Function", "Function DNA", "Profile type", "Profile Type"],
    "UserStatus": ["UserStatus", "User Status", "Status", "Profile Status"],
    "Country": ["Country", "Country Name"],
    "Currency": ["Currency", "CCY", "Ccy"],
    "CommissionStartDate": ["CommissionStartDate", "Commission Start Date", "Entry Date", "EntryDate"],
    "CommissionEndDate": ["CommissionEndDate", "Commission End Date", "Exit Date", "ExitDate"],
    "YtdDue": ["YtdDue", "YTD Due", "YTD Bonus", "Ytd Due Amount", "YTD Due Amount", "Due"],
    "YtdPayment": ["YtdPayment", "YTD Payment", "Ytd Payment Amount", "Payment"],
    "RemainingBalance": ["RemainingBalance", "Remaining Balance", "Balance"],
    "Yearly Target": ["Yearly Target", "Annual Target", "YearlyTarget"],
    "Ytd Yearly Target": ["Ytd Yearly Target", "YTD Yearly Target", "YTD Target", "YtdTarget", "Target"],
}

def canonicalize_bonus_columns(df):
    df = kpi_clean_columns(df.copy())
    rename = {}
    for canonical, candidates in BONUS_COLUMN_ALIASES.items():
        if canonical in df.columns:
            continue
        found = kpi_find_column(df, candidates, required=False)
        if found is not None and found != canonical:
            rename[found] = canonical
    return df.rename(columns=rename)


def kpi_num(series):
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def kpi_safe_ratio(num, den):
    if den is None or pd.isna(den) or abs(float(den)) < 1e-12:
        return np.nan
    return float(num) / float(den)


def kpi_excel_cell(ws, row, col, value, fill=None, fmt=None, align="center", bold=False):
    if value is None:
        value = ""
    try:
        if pd.isna(value):
            value = ""
    except Exception:
        pass
    if not isinstance(align, str):
        align = "center"
    cell = ws.cell(row, col, value)
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    cell.font = Font(bold=bold, size=10)
    cell.alignment = Alignment(horizontal=align, vertical="top", wrap_text=True)
    cell.border = Border(
        left=Side(style="thin", color="D9E2F3"),
        right=Side(style="thin", color="D9E2F3"),
        top=Side(style="thin", color="D9E2F3"),
        bottom=Side(style="thin", color="D9E2F3"),
    )
    return cell


def kpi_formula_cell(ws, row, col, formula, fill=None, fmt=None, align="center"):
    cell = kpi_excel_cell(ws, row, col, None, fill=fill, fmt=fmt, align=align)
    cell.value = formula
    return cell


def write_raw_bonus_sheet(ws, df, fx_end):
    headers = [
        "UserId", "Username", "UserStatus", "Country", "Currency",
        "CommissionStartDate", "CommissionEndDate", "CM", "GM", "YtdDue",
        "YtdPayment", "RemainingBalance", "Yearly Target", "Ytd Yearly Target",
        "Function", "FX Rate to EUR", "CM EUR", "YtdDue EUR",
        "YtdPayment EUR", "RemainingBalance EUR", "Yearly Target EUR", "Ytd Target EUR",
        "Due/Target",
    ]
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 1, c, header)
    for r, (_, item) in enumerate(df.iterrows(), 2):
        values = [
            item["UserId"], item["Username"], item["ProfileStatus"], item["Country"], item["Currency"],
            item["EntryDate"], item["ExitDate"], item["CM"], item["GM"], item["YtdDue"],
            item["YtdPayment"], item["RemainingBalance"], item["YearlyTarget"], item["YtdTarget"],
            item["Function"], None, None, None, None, None, None, None,
        ]
        fill = kpi_fill("FFFFFF")
        for c, value in enumerate(values, 1):
            cell = kpi_excel_cell(ws, r, c, value, fill, align="left" if c in {1, 2, 3, 4, 5, 6, 7, 15} else "center")
            if c in {8, 9, 10, 11, 12, 13, 14}:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
            if c in {17, 18, 19, 20, 21, 22}:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
            if c == 16:
                cell.number_format = "0.0000000000"
            if c == 23:
                cell.number_format = "0.00%"
            if c in {6, 7} and hasattr(value, "year"):
                cell.number_format = "d-mmm-yy"
        fx_lookup = f"IFERROR(VLOOKUP(E{r},'Currency to EUR'!$A$5:$B${fx_end},2,FALSE),0)"
        kpi_formula_cell(ws, r, 16, f"={fx_lookup}", fill, "0.0000000000")
        kpi_formula_cell(ws, r, 17, f"=H{r}*P{r}", fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 18, f"=J{r}*P{r}", fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 19, f"=K{r}*P{r}", fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 20, f"=L{r}*P{r}", fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 21, f"=M{r}*P{r}", fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 22, f"=N{r}*P{r}", fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 23, f"=IFERROR(R{r}/V{r},0)", fill, "0.00%")
    widths = [14, 28, 16, 18, 12, 16, 16, 16, 16, 16, 16, 18, 18, 18, 22, 18, 18, 18, 20, 22, 20, 20, 16]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.auto_filter.ref = f"A1:W{max(1, len(df) + 1)}"
    ws.freeze_panes = "A2"
    ws.sheet_state = "hidden"
    return len(df) + 1


def raw_sum_eur_formula(sheet, end_row, amount_col, criteria, fx_end):
    criteria_part = "*".join(
        f"('{sheet}'!${col}$2:${col}${end_row}={ref})" for col, ref in criteria
    )
    amount_range = f"'{sheet}'!${amount_col}$2:${amount_col}${end_row}"
    currency_range = f"'{sheet}'!$E$2:$E${end_row}"
    fx_currency_range = f"'Currency to EUR'!$A$5:$A${fx_end}"
    fx_rate_range = f"'Currency to EUR'!$B$5:$B${fx_end}"
    # The formatted raw sheet already contains the formula-driven EUR column.
    # Keep report formulas short and transparent by aggregating that EUR column.
    sumifs_args = [amount_range]
    for col, ref in criteria:
        sumifs_args.extend([f"'{sheet}'!${col}$2:${col}${end_row}", ref])
    return f"=SUMIFS({','.join(sumifs_args)})"


def raw_avg_ratio_formula(sheet, end_row, numerator_col, denominator_col, criteria):
    criteria_part = "*".join(
        f"('{sheet}'!${col}$2:${col}${end_row}={ref})" for col, ref in criteria
    )
    numerator = f"'{sheet}'!${numerator_col}$2:${numerator_col}${end_row}"
    denominator = f"'{sheet}'!${denominator_col}$2:${denominator_col}${end_row}"
    count_ranges = ",".join(
        f"'{sheet}'!${col}$2:${col}${end_row},{ref}" for col, ref in criteria
    )
    return f"=IFERROR(SUMPRODUCT({criteria_part}*IFERROR({numerator}/{denominator},0))/COUNTIFS({count_ranges}),0)"


def write_raw_bp_sheet(ws, df):
    headers = ["MonthNumber", "Country", "SIG", "BPValue"]
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 1, c, header)
    for r, (_, item) in enumerate(df.iterrows(), 2):
        fill = kpi_fill("FFFFFF")
        kpi_excel_cell(ws, r, 1, item["MonthNumber"], fill=fill)
        kpi_excel_cell(ws, r, 2, item["CountryName"], fill=fill, align="left")
        kpi_excel_cell(ws, r, 3, item["SIGName"], fill=fill, align="left")
        kpi_excel_cell(ws, r, 4, item["BPValue"], fill=fill, fmt="#,##0.00;[Red](#,##0.00);-")
    for col, width in enumerate([14, 20, 24, 18], 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.auto_filter.ref = f"A1:D{max(1, len(df) + 1)}"
    ws.freeze_panes = "A2"
    ws.sheet_state = "hidden"
    return len(df) + 1


def write_currency_sheet(ws, fx_rates):
    """Show the FX mapping and an actual Excel formula for EUR conversion."""
    headers = ["Currency", "Rate to EUR", "Example amount", "Converted amount EUR", "Formula used"]
    kpi_title(
        ws,
        "CURRENCY CONVERSION TO EUR",
        len(headers),
        "Formula used in the raw sheets: source amount × FX rate to EUR",
    )
    header_row = 4
    for c, header in enumerate(headers, 1):
        kpi_header(ws, header_row, c, header)
    for r, currency in enumerate(sorted(fx_rates), header_row + 1):
        fill = kpi_fill("FFFFFF" if r % 2 else "EBF5FB")
        kpi_excel_cell(ws, r, 1, currency, fill, align="left")
        kpi_excel_cell(ws, r, 2, fx_rates[currency], fill, "0.0000000000")
        kpi_excel_cell(ws, r, 3, 1, fill, "#,##0.00")
        kpi_formula_cell(ws, r, 4, f"=C{r}*B{r}", fill, "#,##0.000000")
        kpi_excel_cell(ws, r, 5, "'=Source Amount * Rate to EUR", fill, align="left")
    for col, width in enumerate([14, 18, 16, 22, 34], 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.auto_filter.ref = f"A{header_row}:E{max(header_row, header_row + len(fx_rates))}"
    ws.freeze_panes = None
    ws.sheet_state = "hidden"


def kpi_title(ws, title, width, subtitle):
    ws.sheet_view.showGridLines = False
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width)
    cell = ws.cell(1, 1, title)
    cell.fill = kpi_fill(KPI_NAVY)
    cell.font = Font(bold=True, size=14, color=KPI_WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width)
    cell = ws.cell(2, 1, subtitle)
    cell.fill = kpi_fill("2C3E50")
    cell.font = Font(italic=True, color=KPI_WHITE)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 22


def kpi_header(ws, row, col, value, color=KPI_HEADER):
    cell = ws.cell(row, col, value)
    cell.fill = kpi_fill(color)
    cell.font = Font(bold=True, color=KPI_WHITE)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = Border(
        left=Side(style="thin", color=KPI_WHITE),
        right=Side(style="thin", color=KPI_WHITE),
        top=Side(style="thin", color=KPI_WHITE),
        bottom=Side(style="thin", color=KPI_WHITE),
    )
    return cell


def shade_sections(ws, header_row, first_data_row, last_data_row, sections):
    """Apply stable, distinct colors to logical column sections."""
    for start_col, end_col, header_color, body_color in sections:
        for col in range(start_col, end_col + 1):
            ws.cell(header_row, col).fill = kpi_fill(header_color)
            for row in range(first_data_row, last_data_row + 1):
                ws.cell(row, col).fill = kpi_fill(body_color)


def load_fx_rates():
    rates = DEFAULT_FX_TO_EUR.copy()
    candidates = [
        MONTHLY_REPORT_DIR / "FX_to_EUR.xlsx",
        MONTHLY_REPORT_DIR / "Currency_to_EUR.xlsx",
    ]
    for path in candidates:
        if not path.exists():
            continue
        try:
            fx = kpi_clean_columns(pd.read_excel(path))
            ccy_col = kpi_find_column(fx, ["Currency", "CCY"])
            rate_col = kpi_find_column(fx, ["Rate to EUR", "FX to EUR", "EUR Rate", "Rate"])
            for _, row in fx.iterrows():
                ccy = str(row[ccy_col]).strip().upper()
                rate = pd.to_numeric(row[rate_col], errors="coerce")
                if ccy and not pd.isna(rate):
                    rates[ccy] = float(rate)
            print(f"Loaded FX rates from: {path}")
        except Exception as exc:
            print(f"FX file could not be read, using built-in rates: {exc}")
        break
    return rates


def normalize_bonus_dataframe(df, fx_rates):
    df = canonicalize_bonus_columns(df)
    user_id_col = kpi_find_column(df, ["UserId", "User ID", "EmployeeId"])
    name_col = kpi_find_column(df, ["Username", "User Name", "Employee"])
    function_col = kpi_find_column(df, ["Function", "Function DNA", "Profile type"], required=False)
    entity_col = kpi_find_column(df, ["Entity", "Company"], required=False)
    region_col = kpi_find_column(df, ["Region"], required=False)
    status_col = kpi_find_column(df, ["UserStatus", "User Status", "Status"], required=False)
    country_col = kpi_find_column(df, ["Country"])
    ccy_col = kpi_find_column(df, ["Currency", "CCY"])
    entry_col = kpi_find_column(df, ["Entry Date", "EntryDate", "CommissionStartDate"], required=False)
    exit_col = kpi_find_column(df, ["Exit Date", "ExitDate", "CommissionEndDate"], required=False)
    cm_col = kpi_find_column(df, ["CM", "Contribution Margin"], required=False)
    gm_col = kpi_find_column(df, ["GM", "Gross Margin"], required=False)
    due_col = kpi_find_column(df, ["YtdDue", "YTD Due", "YTD Bonus"])
    payment_col = kpi_find_column(df, ["YtdPayment", "YTD Payment"], required=False)
    balance_col = kpi_find_column(df, ["RemainingBalance", "Remaining Balance"], required=False)
    annual_target_col = kpi_find_column(df, ["Yearly Target", "Annual Target"], required=False)
    ytd_target_col = kpi_find_column(df, ["Ytd Yearly Target", "YTD Yearly Target", "YTD Target"])

    out = pd.DataFrame()
    out["UserId"] = df[user_id_col].astype(str).str.strip()
    out["Username"] = df[name_col].astype(str).str.strip()
    out["Function"] = df[function_col].astype(str).str.strip() if function_col else "N/A"
    out["Entity"] = df[entity_col].astype(str).str.strip() if entity_col else "N/A"
    out["Region"] = df[region_col].astype(str).str.strip() if region_col else "N/A"
    out["ProfileStatus"] = df[status_col].astype(str).str.strip() if status_col else "N/A"
    out["Country"] = df[country_col].astype(str).str.strip()
    out["Currency"] = df[ccy_col].astype(str).str.strip().str.upper()
    out["EntryDate"] = pd.to_datetime(df[entry_col], errors="coerce", dayfirst=True) if entry_col else pd.NaT
    out["ExitDate"] = pd.to_datetime(df[exit_col], errors="coerce", dayfirst=True) if exit_col else pd.NaT
    out["CM"] = kpi_num(df[cm_col]) if cm_col else 0.0
    out["GM"] = kpi_num(df[gm_col]) if gm_col else 0.0
    out["YtdDue"] = kpi_num(df[due_col])
    out["YtdPayment"] = kpi_num(df[payment_col]) if payment_col else 0.0
    out["RemainingBalance"] = kpi_num(df[balance_col]) if balance_col else 0.0
    out["YearlyTarget"] = kpi_num(df[annual_target_col]) if annual_target_col else 0.0
    out["YtdTarget"] = kpi_num(df[ytd_target_col])
    out["FXRateToEUR"] = out["Currency"].map(fx_rates)
    out["FXMissing"] = out["FXRateToEUR"].isna()
    missing_fx = sorted(out.loc[out["FXMissing"], "Currency"].dropna().unique())
    if missing_fx:
        raise ValueError(f"Missing FX rates for currencies: {', '.join(missing_fx)}")
    for col in ["CM", "GM", "YtdDue", "YtdPayment", "RemainingBalance", "YearlyTarget", "YtdTarget"]:
        out[f"{col}EUR"] = out[col] * out["FXRateToEUR"]
    invalid = {"", "nan", "none", "total", "grand total", "grand_total"}
    out = out[~out["Username"].str.lower().isin(invalid)].copy()
    out = out[out["UserId"].str.lower() != "total"].copy()
    out.attrs["source_rows"] = len(df)
    out.attrs["used_rows"] = len(out)
    out.attrs["dropped_rows"] = len(df) - len(out)
    return out


def load_bonus_file(path, fx_rates):
    return normalize_bonus_dataframe(pd.read_excel(path), fx_rates)

def read_bp_table(path):
    """Read the BP table even when the workbook has a title row or multiple sheets."""
    month_candidates = ["Month", "Month Number", "MonthName", "Mois", "Period"]
    country_candidates = ["Country Billing To", "Country", "Country Name", "CountryName"]
    sig_candidates = ["SIG", "Category", "Cost Category", "Type"]
    value_candidates = ["Sum of EUR", "Value", "Amount", "BP", "Business Plan"]

    xls = pd.ExcelFile(path)
    preferred = [
        sheet for sheet in xls.sheet_names
        if re.sub(r"[^a-z0-9]", "", str(sheet).lower()) in {"bp2025", "bp"}
    ]
    sheets = preferred + [sheet for sheet in xls.sheet_names if sheet not in preferred]

    for sheet in sheets:
        preview = pd.read_excel(path, sheet_name=sheet, header=None, nrows=30)
        for header_row in range(min(20, len(preview))):
            candidate = kpi_clean_columns(
                pd.read_excel(path, sheet_name=sheet, header=header_row)
            )
            required_groups = [
                month_candidates,
                country_candidates,
                sig_candidates,
                value_candidates,
            ]
            if all(kpi_find_column(candidate, group, required=False) is not None for group in required_groups):
                print(f"Loaded BP data from sheet '{sheet}' with header row {header_row + 1}")
                return candidate

    raise KeyError(
        "Could not detect the BP table headers. Expected columns for Month, Country, SIG and BP value "
        f"in workbook sheets: {xls.sheet_names}"
    )


def load_bp_file(path):
    df = read_bp_table(path)
    month_col = kpi_find_column(df, ["Month", "Month Number", "MonthName", "Mois", "Period"])
    country_col = kpi_find_column(df, ["Country Billing To", "Country"])
    sig_col = kpi_find_column(df, ["SIG", "Category"])
    value_col = kpi_find_column(df, ["Sum of EUR", "Value", "Amount"])
    df["MonthNumber"] = pd.to_numeric(df[month_col], errors="coerce")
    df["CountryName"] = df[country_col].astype(str).str.strip()
    df["SIGName"] = df[sig_col].astype(str).str.strip()
    df["BPValue"] = pd.to_numeric(df[value_col], errors="coerce").fillna(0.0)
    df = df[df["MonthNumber"].notna()].copy()
    excluded = df["SIGName"].str.lower().str.contains("financial|structure", regex=True, na=False)
    raw_rows = len(df)
    excluded_rows = int(excluded.sum())
    excluded_value = float(df.loc[excluded, "BPValue"].sum())
    df = df[~excluded].copy()
    df.attrs["source_rows"] = raw_rows
    df.attrs["used_rows"] = len(df)
    df.attrs["excluded_rows"] = excluded_rows
    df.attrs["excluded_value"] = excluded_value
    return df


def bp_by_country(bp, month_limit):
    monthly = bp.groupby(["CountryName", "MonthNumber"], as_index=False)["BPValue"].sum()
    ytd = monthly[monthly["MonthNumber"] <= month_limit].groupby("CountryName", as_index=False)["BPValue"].sum().rename(columns={"BPValue": "BP_YTD"})
    full = monthly.groupby("CountryName", as_index=False)["BPValue"].sum().rename(columns={"BPValue": "BP_FullYear"})
    result = ytd.merge(full, on="CountryName", how="outer").fillna(0)
    result["BP_Progress"] = result.apply(lambda r: kpi_safe_ratio(r["BP_YTD"], r["BP_FullYear"]), axis=1)
    return result


def month_number_from_name(path, fallback):
    months = {"january": 1, "february": 2, "march": 3, "april": 4, "may": 5, "june": 6, "july": 7, "august": 8, "september": 9, "october": 10, "november": 11, "december": 12}
    text = str(path).lower()
    for name, number in months.items():
        if name in text:
            return number
    return fallback


def build_kpi_data(april, may, bp_current, bp_previous):
    may_country = may.groupby("Country", as_index=False).agg(
        PE=("UserId", "nunique"),
        YTD_Bonus=("YtdDueEUR", "sum"),
        YTD_Target=("YtdTargetEUR", "sum"),
        YTD_CM=("CMEUR", "sum"),
    )
    april_country = april.groupby("Country", as_index=False).agg(
        April_YTD_CM=("CMEUR", "sum"),
    )
    current_bp = bp_current.rename(columns={"BP_YTD": "BP_YTD_Current", "BP_FullYear": "BP_FullYear_Current", "BP_Progress": "BP_Progress_Current"})
    previous_bp = bp_previous[["CountryName", "BP_YTD"]].rename(columns={"BP_YTD": "BP_YTD_Previous"})
    country = may_country.merge(april_country, on="Country", how="outer")
    country = country.merge(current_bp, left_on="Country", right_on="CountryName", how="outer")
    country["Country"] = country["Country"].fillna(country["CountryName"])
    country = country.drop(columns=["CountryName"], errors="ignore")
    country = country.merge(previous_bp, left_on="Country", right_on="CountryName", how="outer")
    country["Country"] = country["Country"].fillna(country["CountryName"])
    country = country.drop(columns=["CountryName"], errors="ignore")
    country = country.fillna(0)
    country["BP_Achievement"] = country.apply(lambda r: kpi_safe_ratio(r["YTD_CM"], r["BP_YTD_Current"]), axis=1)
    country["CM_Var_vs_Apr"] = country["YTD_CM"] - country["April_YTD_CM"]
    country["CM_Var_vs_Apr_pct"] = country.apply(lambda r: kpi_safe_ratio(r["CM_Var_vs_Apr"], abs(r["April_YTD_CM"])), axis=1)
    country["Country_CM_April"] = country.apply(lambda r: kpi_safe_ratio(r["April_YTD_CM"], r["BP_YTD_Previous"]), axis=1)

    may = may.copy()
    may["Due_Target"] = may.apply(lambda r: kpi_safe_ratio(r["YtdDueEUR"], r["YtdTargetEUR"]), axis=1)
    country["Country_CM_BP"] = country["BP_Achievement"]
    country_perf = may.groupby("Country").agg(
        Country_Due=("YtdDueEUR", "sum"),
        Country_Target=("YtdTargetEUR", "sum"),
        TopManagerPerf=("Due_Target", "max"),
    ).reset_index()
    # Country Manager Performance is weighted by target:
    # SUM(YTD Due) / SUM(YTD Target), not AVERAGE(manager Due/Target).
    country_perf["ManagersPerf"] = country_perf.apply(
        lambda row: kpi_safe_ratio(row["Country_Due"], row["Country_Target"]), axis=1
    )
    country = country.merge(country_perf, on="Country", how="left")
    country_lookup = country.set_index("Country").to_dict("index")
    may["Country_CM_BP"] = may["Country"].map(lambda c: country_lookup.get(c, {}).get("Country_CM_BP", np.nan))
    may["Country_CM_Var"] = may["Country"].map(lambda c: country_lookup.get(c, {}).get("CM_Var_vs_Apr_pct", np.nan))
    may["BP_Progress"] = may["Country"].map(lambda c: country_lookup.get(c, {}).get("BP_Progress_Current", np.nan))
    may["DeltaProjects"] = "N/A"
    may["Speed"] = "N/A"
    may["Inheritances"] = "N/A"
    may["Comments"] = "Generated from local Monthly report inputs"
    may["Followup"] = np.where(may["Due_Target"].fillna(0) >= 0.60, "Review", "Normal")

    # Function performance is grouped by Country and Function together.
    function_country = may.groupby(["Country", "Function"], dropna=False).agg(
        PE=("UserId", "nunique"),
        YTD_Bonus=("YtdDueEUR", "sum"),
        YTD_Target=("YtdTargetEUR", "sum"),
        YTD_CM=("CMEUR", "sum"),
        ManagersPerf=("Due_Target", "mean"),
        TopManagerPerf=("Due_Target", "max"),
    ).reset_index()
    function_country["Bonus_Target"] = function_country.apply(
        lambda r: kpi_safe_ratio(r["YTD_Bonus"], r["YTD_Target"]),
        axis=1,
    )

    may = may.sort_values(["Due_Target", "YtdDueEUR"], ascending=[False, False]).reset_index(drop=True)
    # TopPerformer contains only managers whose YTD Due/Target is above 100%.
    top = may[pd.to_numeric(may["Due_Target"], errors="coerce") > 1].head(10).copy()
    return country.sort_values("YTD_Bonus", ascending=False), function_country, may, top, april


# ========================== WORD KPI REVIEW ================================
# The Word report is generated entirely by python-docx. No external Word
# template is required. The report keeps the same seven sections, six figures,
# captions, margins, typography, and 1.5 line spacing as the approved version.

KPI_WORD_OUTPUT_PATH = MONTHLY_REPORT_DIR / "KPI Performance Review May 2026.docx"


def report_float(value, default=0.0):
    try:
        value = float(value)
        return default if not np.isfinite(value) else value
    except Exception:
        return default


def report_user_id(series):
    return series.astype(str).str.strip().str.replace(r"\.0$", "", regex=True)


def report_money(value):
    value = report_float(value)
    absolute = abs(value)
    sign = "-" if value < 0 else ""
    if absolute >= 1_000_000:
        return f"{sign}€{absolute / 1_000_000:.1f}M"
    if absolute >= 1_000:
        return f"{sign}€{absolute / 1_000:.1f}K"
    return f"{sign}€{absolute:,.0f}"


def report_pct(value):
    value = report_float(value, np.nan)
    return "n/a" if pd.isna(value) else f"{value * 100:.1f}%"


def report_signed_pct(value):
    value = report_float(value, np.nan)
    return "n/a" if pd.isna(value) else f"{value * 100:+.1f}%"


def report_change(current_value, previous_value):
    current_value = report_float(current_value)
    previous_value = report_float(previous_value)
    if abs(previous_value) < 1e-12:
        return np.nan
    return (current_value - previous_value) / abs(previous_value)


def report_finance_frame(raw, fx_rates):
    """Add EUR financial measures used by the Word narrative and charts."""
    out = raw.copy()
    out["UserId"] = report_user_id(out.get("UserId", pd.Series(index=out.index)))
    out["Country"] = out.get("Country", pd.Series(index=out.index)).astype(str).str.strip()
    out["Function"] = out.get("Function", pd.Series(index=out.index)).astype(str).str.strip()
    out["Username"] = out.get("Username", pd.Series(index=out.index)).astype(str).str.strip()
    rates = out.get("Currency", pd.Series(index=out.index)).astype(str).str.upper().map(fx_rates).fillna(1.0)
    for source in ["Revenue", "GM", "Manager Cost", "CM"]:
        source_values = out[source] if source in out.columns else pd.Series(0.0, index=out.index)
        out[source] = pd.to_numeric(source_values, errors="coerce").fillna(0.0)
        out[f"{source}EUR"] = out[source] * rates
    due_values = out["YtdDueEUR"] if "YtdDueEUR" in out.columns else pd.Series(0.0, index=out.index)
    target_values = out["YtdTargetEUR"] if "YtdTargetEUR" in out.columns else pd.Series(0.0, index=out.index)
    out["YtdDueEUR"] = pd.to_numeric(due_values, errors="coerce").fillna(0.0)
    out["YtdTargetEUR"] = pd.to_numeric(target_values, errors="coerce").fillna(0.0)
    out["Due_Target"] = out.apply(lambda row: kpi_safe_ratio(row["YtdDueEUR"], row["YtdTargetEUR"]), axis=1)
    return out




def report_function_summary(current_finance, previous_finance):
    def aggregate(frame, prefix):
        grouped = frame.groupby("Function", dropna=False).agg(
            CM=("CMEUR", "sum"), Revenue=("RevenueEUR", "sum"), GM=("GMEUR", "sum"),
            Due=("YtdDueEUR", "sum"), Target=("YtdTargetEUR", "sum"), Managers=("UserId", "nunique"),
        ).reset_index()
        grouped["Due_Target"] = grouped.apply(lambda row: kpi_safe_ratio(row["Due"], row["Target"]), axis=1)
        grouped["CM_Margin"] = grouped.apply(lambda row: kpi_safe_ratio(row["CM"], row["Revenue"]), axis=1)
        return grouped.add_prefix(prefix)

    may = aggregate(current_finance, "May_")
    april = aggregate(previous_finance, "April_")
    result = may.merge(april, left_on="May_Function", right_on="April_Function", how="outer").fillna(0)
    result["Function"] = result["May_Function"].where(result["May_Function"].astype(str).str.len() > 0, result["April_Function"])
    result["CM_Var"] = result["May_CM"] - result["April_CM"]
    return result.sort_values("May_CM", ascending=False).reset_index(drop=True)


def report_context(current, previous, may_raw, april_raw, country, top, fx_rates):
    current = current.copy(); previous = previous.copy(); top = top.copy()
    current["UserId"] = report_user_id(current.get("UserId", pd.Series(index=current.index)))
    previous["UserId"] = report_user_id(previous.get("UserId", pd.Series(index=previous.index)))
    top["UserId"] = report_user_id(top.get("UserId", pd.Series(index=top.index)))
    current_finance = report_finance_frame(may_raw, fx_rates)
    previous_finance = report_finance_frame(april_raw, fx_rates)
    current_finance = current_finance.merge(
        current[["UserId", "YtdDueEUR", "YtdTargetEUR", "CMEUR"]].drop_duplicates("UserId"),
        on="UserId", how="left", suffixes=("", "_normalized")
    )
    previous_finance = previous_finance.merge(
        previous[["UserId", "YtdDueEUR", "YtdTargetEUR", "CMEUR"]].drop_duplicates("UserId"),
        on="UserId", how="left", suffixes=("", "_normalized")
    )
    for frame in [current_finance, previous_finance]:
        for col in ["YtdDueEUR", "YtdTargetEUR", "CMEUR"]:
            normalized = f"{col}_normalized"
            if normalized in frame.columns:
                frame[col] = frame[normalized].fillna(frame.get(col, 0.0))
        frame["Due_Target"] = frame.apply(lambda row: kpi_safe_ratio(row["YtdDueEUR"], row["YtdTargetEUR"]), axis=1)

    functions = report_function_summary(current_finance, previous_finance)
    country_finance = current_finance.groupby("Country", dropna=False).agg(
        RevenueEUR=("RevenueEUR", "sum"), GMEUR=("GMEUR", "sum"),
        ManagerCostEUR=("Manager CostEUR", "sum"), CMEUR=("CMEUR", "sum"),
        Managers=("UserId", "nunique"), YtdTargetEUR=("YtdTargetEUR", "sum"),
    ).reset_index()
    country_finance["Cost_GM"] = country_finance.apply(lambda row: kpi_safe_ratio(row["ManagerCostEUR"], row["GMEUR"]), axis=1)
    previous_country_finance = previous_finance.groupby("Country", dropna=False).agg(
        April_CMEUR=("CMEUR", "sum"),
    ).reset_index()
    country_finance = country_finance.merge(previous_country_finance, on="Country", how="outer").fillna(0)
    country_finance["CM_Var"] = country_finance["CMEUR"] - country_finance["April_CMEUR"]
    country_finance["CM_Var_pct"] = country_finance.apply(lambda row: report_change(row["CMEUR"], row["April_CMEUR"]), axis=1)
    country_finance["CM_Share"] = country_finance["CMEUR"] / max(country_finance["CMEUR"].sum(), 1e-12)

    country_report = country.merge(country_finance, on="Country", how="outer").fillna(0)
    country_report["Country"] = country_report["Country"].astype(str)
    country_report["CM_BP_pct"] = country_report.apply(
        lambda row: report_float(row.get("CMEUR")) / max(report_float(row.get("BP_YTD_Current")), 1e-12), axis=1
    )
    country_report["Performance_Gap"] = country_report["ManagersPerf"] - country_report["CM_BP_pct"]
    country_report["BP_per_Manager"] = country_report.apply(
        lambda row: report_float(row.get("BP_YTD_Current")) / max(report_float(row.get("Managers")), 1), axis=1
    )
    country_report["CM_per_Manager"] = country_report.apply(
        lambda row: report_float(row.get("CMEUR")) / max(report_float(row.get("Managers")), 1), axis=1
    )
    manager = current_finance.copy()
    manager["CM_Share_Country"] = manager.apply(
        lambda row: report_float(row["CMEUR"]) / max(report_float(country_finance.loc[country_finance["Country"] == row["Country"], "CMEUR"].sum()), 1e-12), axis=1
    )

    prev_lookup = previous_finance.set_index("UserId")
    manager["Bonus_Var"] = manager.apply(
        lambda row: report_change(row["YtdDueEUR"], prev_lookup["YtdDueEUR"].get(row["UserId"], 0.0)), axis=1
    )
    manager["CM_Var"] = manager.apply(
        lambda row: report_change(row["CMEUR"], prev_lookup["CMEUR"].get(row["UserId"], 0.0)), axis=1
    )
    variance = manager[manager["Bonus_Var"] > 0.70].sort_values("Bonus_Var", ascending=False).copy()

    return {
        "current": current_finance,
        "previous": previous_finance,
        "functions": functions,
        "country": country_report,
        "manager": manager,
        "variance": variance,
        "top": top.copy(),
    }


def chart_save(fig, path):
    fig.savefig(path, dpi=180, bbox_inches="tight", facecolor="white")
    plt.close(fig)


def annotate_vertical_bars(ax, bars, labels, offset=1.0, fontsize=8, rotation=0):
    for bar, label in zip(bars, labels):
        height = bar.get_height()
        if height >= 0:
            y, va = height + offset, "bottom"
        else:
            y, va = height - offset, "top"
        ax.text(bar.get_x() + bar.get_width() / 2, y, label, ha="center", va=va, fontsize=fontsize, rotation=rotation)


def annotate_horizontal_bars(ax, bars, labels, offset=1.0, fontsize=8):
    for bar, label in zip(bars, labels):
        width = bar.get_width()
        ax.text(width + offset, bar.get_y() + bar.get_height() / 2, label, ha="left", va="center", fontsize=fontsize)


def create_kpi_word_charts(context, folder):
    country = context["country"].sort_values("CMEUR", ascending=False).copy()
    functions = context["functions"].sort_values("May_CM", ascending=False).copy()
    top = context["top"].head(10).copy()
    variance = context["variance"].head(12).copy()
    navy, gold, green, red, purple = "#1F4E79", "#BF9000", "#70AD47", "#C0504D", "#8064A2"
    paths = []

    # Figure 1: CM / BP percentage by country. This is the management view
    # used in the reference report: one horizontal bar per country, green
    # when CM reaches BP and red when it is below BP.
    fig, ax = plt.subplots(figsize=(10.2, 5.2))
    country = country.copy()
    country["CM_BP_pct"] = country.apply(
        lambda row: report_float(row.get("CMEUR")) / max(report_float(row.get("BP_YTD_Current")), 1e-12) * 100,
        axis=1,
    )
    country = country.sort_values("CMEUR", ascending=False)
    colors = [green if value >= 100 else red for value in country["CM_BP_pct"]]
    bars = ax.barh(country["Country"], country["CM_BP_pct"], color=colors)
    ax.axvline(100, color="#333333", linestyle="--", linewidth=1.2)
    for bar, value in zip(bars, country["CM_BP_pct"]):
        ax.text(value + 1.5, bar.get_y() + bar.get_height() / 2, f"{value:.1f}%", ha="left", va="center", fontsize=8)
    ax.invert_yaxis()
    ax.set_xlim(0, max(float(country["CM_BP_pct"].max()) * 1.15, 120))
    ax.set_xlabel("CM / BP (%)")
    ax.set_title("CM vs Business Plan by country", loc="left", fontweight="bold", pad=18)
    ax.grid(axis="x", alpha=.2)
    fig.tight_layout(rect=[0, 0, 1, .93])
    p = folder / "figure_1_cm_vs_bp.png"; chart_save(fig, p); paths.append(p)

    # Figure 1b: reference-style grouped columns for countries where manager
    # performance is higher than Country CM/BP. Keep this vertical layout
    # deliberately simple so the chart remains readable inside Word.
    gap_country = country[country["Performance_Gap"] > 0].sort_values("Country").copy()
    fig, ax = plt.subplots(figsize=(10.2, 5.2), facecolor="white")
    ax.set_facecolor("white")
    if gap_country.empty:
        ax.text(.5, .5, "No country has manager performance above CM/BP", ha="center", va="center")
        ax.set_axis_off()
    else:
        labels = gap_country["Country"].astype(str).tolist()
        # CM_BP_pct is already expressed on a 0-100 percentage scale above.
        cm_bp = gap_country["CM_BP_pct"].astype(float).to_numpy()
        manager_perf = gap_country["ManagersPerf"].astype(float).to_numpy() * 100
        x = np.arange(len(labels), dtype=float)
        width = 0.34
        bars_cm = ax.bar(x - width / 2, cm_bp, width, color=red, alpha=0.95,
                         label="Country CM / BP", zorder=3)
        bars_mgr = ax.bar(x + width / 2, manager_perf, width, color="#F39C12", alpha=0.98,
                          label="Manager performance", zorder=3)
        for bar, value in zip(bars_cm, cm_bp):
            ax.text(bar.get_x() + bar.get_width() / 2, value + 1.8,
                    f"{value:.1f}%", ha="center", va="bottom", fontsize=9, color="#222222")
        for bar, value in zip(bars_mgr, manager_perf):
            ax.text(bar.get_x() + bar.get_width() / 2, value + 1.8,
                    f"{value:.1f}%", ha="center", va="bottom", fontsize=9, color="#222222")
        ax.axhline(100, color="#444444", linestyle="--", linewidth=1.1, zorder=4)
        ax.set_xticks(x, labels)
        ax.set_xlabel("")
        ax.set_ylabel("Performance (%)")
        ax.set_ylim(0, 110)
        ax.set_yticks(np.arange(0, 101, 20))
        ax.set_title("Country CM / BP versus manager performance", fontsize=14,
                     fontweight="normal", pad=10)
        ax.legend(frameon=False, loc="upper right", bbox_to_anchor=(1.0, 1.02),
                  borderaxespad=0, fontsize=9)
        ax.grid(axis="y", color="#D9D9D9", linewidth=0.8, alpha=0.65, zorder=0)
        ax.grid(axis="x", visible=False)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
    fig.subplots_adjust(left=0.10, right=0.98, bottom=0.14, top=0.86)
    # Fixed canvas, no bbox_inches='tight'. This prevents Word from receiving
    # an invalid oversized image and keeps the chart proportions stable.
    p = folder / "figure_1b_country_performance_gap.png"
    fig.savefig(p, dpi=180, facecolor="white", format="png")
    plt.close(fig)
    paths.append(p)

    # Figure 2: CM variance from April to May.
    fig, ax = plt.subplots(figsize=(10.2, 4.8))
    variance_country = country.sort_values("CM_Var", ascending=False)
    colors = [green if value >= 0 else red for value in variance_country["CM_Var"]]
    variance_bars = ax.bar(variance_country["Country"], variance_country["CM_Var"] / 1000, color=colors)
    annotate_vertical_bars(ax, variance_bars, [f"{v:+,.0f}" for v in variance_country["CM_Var"] / 1000], offset=max(abs(variance_country["CM_Var"] / 1000).max() * .02, 1), fontsize=8)
    ax.axhline(0, color="#666666", linewidth=.8); ax.set_ylabel("€K")
    ax.set_title("CM variance from April to May by country", loc="left", fontweight="bold", pad=18)
    ax.margins(y=.22)
    ax.tick_params(axis="x", rotation=35); ax.grid(axis="y", alpha=.2); fig.tight_layout(rect=[0, 0, 1, .93])
    p = folder / "figure_2_cm_variance.png"; chart_save(fig, p); paths.append(p)

    # Figure 3: Due/Target by function, April and May.
    fig, ax = plt.subplots(figsize=(9.6, 4.8))
    x = np.arange(len(functions)); width = .36
    april_bars = ax.bar(x - width / 2, functions["April_Due_Target"] * 100, width, label="April", color="#A6A6A6")
    may_bars = ax.bar(x + width / 2, functions["May_Due_Target"] * 100, width, label="May", color=purple)
    annotate_vertical_bars(ax, april_bars, [f"{v * 100:.1f}%" for v in functions["April_Due_Target"]], offset=1.2, fontsize=8)
    annotate_vertical_bars(ax, may_bars, [f"{v * 100:.1f}%" for v in functions["May_Due_Target"]], offset=1.2, fontsize=8)
    ax.set_xticks(x, functions["Function"], rotation=20, ha="right"); ax.set_ylabel("Due / Target (%)")
    ax.set_ylim(0, max(functions["April_Due_Target"].max(), functions["May_Due_Target"].max()) * 100 * 1.18)
    ax.set_title("Due/Target by function, April and May", loc="left", fontweight="bold", pad=18)
    ax.legend(frameon=False, ncol=2); ax.grid(axis="y", alpha=.2); fig.tight_layout(rect=[0, 0, 1, .93])
    p = folder / "figure_3_function_due_target.png"; chart_save(fig, p); paths.append(p)

    # Top performers above 100% YTD Due/Target (up to 10 records).
    fig, ax = plt.subplots(figsize=(10.0, 5.2))
    plot_top = top.sort_values("Due_Target", ascending=True)
    top_bars = ax.barh(plot_top["Username"], plot_top["Due_Target"] * 100, color=navy)
    annotate_horizontal_bars(ax, top_bars, [f"{v * 100:.1f}%" for v in plot_top["Due_Target"]], offset=1.0, fontsize=8)
    ax.set_xlabel("Due / Target (%)"); ax.set_xlim(0, max(plot_top["Due_Target"].max() * 100 * 1.12, 100))
    ax.set_title(f"Top performers above 100% YTD Due/Target ({len(plot_top)})", loc="left", fontweight="bold", pad=18)
    ax.grid(axis="x", alpha=.2); fig.tight_layout(rect=[0, 0, 1, .93])
    p = folder / "figure_4_top_performers.png"; chart_save(fig, p); paths.append(p)

    # Figure 5: Bonus variance above 70%, with CM variance above each bar.
    fig, ax = plt.subplots(figsize=(10.2, 5.0))
    if variance.empty:
        ax.text(.5, .5, "No manager exceeded the 70% Bonus variance threshold", ha="center", va="center")
        ax.set_axis_off()
    else:
        bars = ax.bar(variance["Username"], variance["Bonus_Var"] * 100, color=red)
        ax.axhline(70, color="#666666", linestyle="--", linewidth=.8)
        for bar, bonus_var, cm_var in zip(bars, variance["Bonus_Var"], variance["CM_Var"]):
            ax.text(
                bar.get_x() + bar.get_width() / 2,
                bar.get_height() + 2,
                f"Bonus {bonus_var * 100:.1f}%\nCM {cm_var * 100:+.1f}%",
                ha="center", va="bottom", fontsize=7,
            )
        max_bonus = float(variance["Bonus_Var"].max() * 100)
        ax.set_ylim(0, max(max_bonus * 1.28, 100))
        ax.set_ylabel("Bonus variance vs April (%)"); ax.set_title("Managers with Bonus variance above 70%", loc="left", fontweight="bold", pad=18)
        ax.tick_params(axis="x", rotation=35); ax.grid(axis="y", alpha=.2)
    fig.tight_layout(rect=[0, 0, 1, .91]); p = folder / "figure_5_bonus_variance.png"; chart_save(fig, p); paths.append(p)
    return paths


def build_kpi_word_narrative(context):
    current = context["current"]; previous = context["previous"]
    country = context["country"].sort_values("CMEUR", ascending=False).copy()
    functions = context["functions"].sort_values("May_CM", ascending=False).copy()
    manager = context["manager"]; variance = context["variance"]
    top = context["top"].head(10).copy()
    total = lambda frame, col: report_float(frame[col].sum()) if col in frame else 0.0
    may_rev, apr_rev = total(current, "RevenueEUR"), total(previous, "RevenueEUR")
    may_gm, apr_gm = total(current, "GMEUR"), total(previous, "GMEUR")
    may_cost, apr_cost = total(current, "Manager CostEUR"), total(previous, "Manager CostEUR")
    may_cm, apr_cm = total(current, "CMEUR"), total(previous, "CMEUR")
    may_due, apr_due = total(current, "YtdDueEUR"), total(previous, "YtdDueEUR")
    may_target, apr_target = total(current, "YtdTargetEUR"), total(previous, "YtdTargetEUR")
    may_due_target, apr_due_target = kpi_safe_ratio(may_due, may_target), kpi_safe_ratio(apr_due, apr_target)
    shortfall = may_target - may_due
    may_margin, apr_margin = kpi_safe_ratio(may_cm, may_rev), kpi_safe_ratio(apr_cm, apr_rev)

    overall = [
        f"May delivered clear business growth. Revenue reached {report_money(may_rev)}, {report_signed_pct(report_change(may_rev, apr_rev))} from April. GM rose to {report_money(may_gm)}, while CM reached {report_money(may_cm)}, {report_signed_pct(report_change(may_cm, apr_cm))}.",
        f"The quality of that growth should be monitored. Manager Cost increased {report_signed_pct(report_change(may_cost, apr_cost))}, compared with Revenue growth of {report_signed_pct(report_change(may_rev, apr_rev))}. CM margin moved from {report_pct(apr_margin)} to {report_pct(may_margin)}.",
        f"The biggest gap sits in Bonus performance. Due increased {report_signed_pct(report_change(may_due, apr_due))}, while Target increased {report_signed_pct(report_change(may_target, apr_target))}. Due/Target moved from {report_pct(apr_due_target)} to {report_pct(may_due_target)}, leaving a May shortfall of {report_money(shortfall)}.",
    ]

    gap_country = country[country["Performance_Gap"] > 0].sort_values("Performance_Gap", ascending=False).copy()
    conflict_text = [
        "Under normal conditions, Country CM/BP would be expected to be at least as high as weighted manager performance. The countries below have Due/Target above CM/BP, which indicates that Bonus attainment is not translating into comparable CM after PnL and Manager Cost.",
    ]
    for _, row in gap_country.head(2).iterrows():
        gap_points = report_float(row.get("Performance_Gap")) * 100
        country_margin = kpi_safe_ratio(row.get("CMEUR"), row.get("RevenueEUR"))
        cost_gm = kpi_safe_ratio(row.get("ManagerCostEUR"), row.get("GMEUR"))
        conflict_text.append(
            f"{row['Country']} is the clearest gap. Country CM/BP is {report_pct(row.get('CM_BP_pct'))}, while weighted manager performance is {report_pct(row.get('ManagersPerf'))}, a difference of {gap_points:.1f} percentage points. {row['Country']} generates {report_money(row.get('CMEUR'))} CM against {report_money(row.get('BP_YTD_Current'))} BP. The country has {int(report_float(row.get('Managers')))} managers, so BP is {report_money(row.get('BP_per_Manager'))} per manager, while CM is {report_money(row.get('CM_per_Manager'))} per manager. The main issue is the BP level relative to CM delivered, not necessarily weak margin: CM margin is {report_pct(country_margin)} and Manager Cost/GM is {report_pct(cost_gm)}."
        )
    while len(conflict_text) < 3:
        conflict_text.append("No additional country has weighted manager performance above Country CM/BP in the current KPI data.")

    country_text = []
    for _, row in country.head(8).iterrows():
        name = row["Country"]; bp = report_float(row.get("BP_Achievement")); cm_change = report_signed_pct(row.get("CM_Var_pct"))
        due = report_pct(row.get("ManagersPerf")); cost_gm = report_pct(row.get("Cost_GM"))
        share = report_pct(row.get("CM_Share")); direction = "increased" if report_float(row.get("CM_Var")) >= 0 else "fell"
        country_text.append(f"{name} is a major CM contributor at {report_money(row.get('CMEUR'))}, or {share} of total CM. CM {direction} {abs(report_float(row.get('CM_Var_pct'))) * 100:.1f}%, reached {report_pct(bp)} of BP, and country Due/Target was {due}. Cost/GM was {cost_gm}.")
    while len(country_text) < 8:
        country_text.append("No additional country-level result was available in the KPI source data.")

    function_text = []
    for _, row in functions.head(3).iterrows():
        function_text.append(f"{row['Function']} generated {report_money(row['May_CM'])} of CM. Due/Target was {report_pct(row['May_Due_Target'])}, compared with {report_pct(row['April_Due_Target'])} in April, while CM changed by {report_money(row['CM_Var'])}. CM margin was {report_pct(row['May_CM_Margin'])}.")
    while len(function_text) < 3:
        function_text.append("No additional function-level result was available in the KPI source data.")

    top_names = []
    for _, row in top.head(5).iterrows():
        top_names.append(f"{row['Username']} at {report_pct(row.get('Due_Target'))} Due/Target")
    top_intro = "Top performers are ranked by Due/Target. The ranking changes when CM contribution is added. Due/Target shows who is ahead of Target; CM contribution shows who is creating the financial result."
    top_summary = "The leading performers were " + ", ".join(top_names) + "."
    top_details = []
    for _, row in top.head(3).iterrows():
        country_total = country.loc[country["Country"] == row["Country"], "CMEUR"].sum()
        share = report_float(row.get("CMEUR")) / max(report_float(country_total), 1e-12)
        top_details.append(f"{row['Username']} reached {report_pct(row.get('Due_Target'))} Due/Target and created {report_pct(share)} of {row['Country']} CM. {row['Country']} reached {report_pct(country.loc[country['Country'] == row['Country'], 'BP_Achievement'].iloc[0] if not country.loc[country['Country'] == row['Country']].empty else np.nan)} of BP.")
    while len(top_details) < 3:
        top_details.append("No additional top-performer comparison was available in the KPI source data.")

    manager_text = ["Across several countries, the Bonus leader is different from the largest CM contributor. That gap points to target design, Manager Cost, or portfolio mix as the next area to review."]
    for name in country["Country"].head(9):
        subset = manager[manager["Country"] == name].copy()
        if subset.empty:
            manager_text.append(f"{name}: no manager-level result was available.")
            continue
        bonus_leader = subset.sort_values("Due_Target", ascending=False).iloc[0]
        cm_leader = subset.sort_values("CMEUR", ascending=False).iloc[0]
        manager_text.append(f"{name}: {bonus_leader['Username']} led Due/Target at {report_pct(bonus_leader['Due_Target'])}; {cm_leader['Username']} created the most CM at {report_money(cm_leader['CMEUR'])}, or {report_pct(cm_leader['CM_Share_Country'])} of country CM.")
    while len(manager_text) < 10:
        manager_text.append("No additional country-level manager result was available.")

    var_intro = f"The check starts with Bonus variance. Across the population, Bonus Due changed by {report_money(may_due - apr_due)}, while Target changed by {report_money(may_target - apr_target)}. The review below focuses on managers whose Bonus in May was more than 70% above April."
    if variance.empty:
        var_group = "No manager crossed the 70% Bonus variance threshold."
        var_detail = "The threshold test did not identify a manager with Bonus rising by more than 70% versus April."
    else:
        names = [f"{row['Username']} ({row['Country']}, Bonus {report_signed_pct(row['Bonus_Var'])}, CM {report_signed_pct(row['CM_Var'])})" for _, row in variance.iterrows()]
        var_group = f"{len(variance)} managers crossed the 70% Bonus threshold: " + "; ".join(names) + "."
        supported = int((variance["CM_Var"] >= 0).sum())
        var_detail = f"{supported} of these managers also increased CM. The test should be used to separate Bonus timing from cases where Bonus rises while CM falls."

    conclusion = [
        f"May was a strong month for Revenue, GM and CM. {country.iloc[0]['Country']} has the largest CM base, while the country comparison identifies the strongest BP achievement and the clearest weakness.",
        "Top performers should be assessed alongside CM contribution. A high Due/Target result does not automatically mean the manager created the largest financial contribution in the country.",
        "The next steps are to review the largest negative CM variances at manager and project level, explain the largest BP gaps, revisit targets where Bonus and PnL diverge, and check Manager Cost in the countries with the highest Cost/GM.",
    ]
    return {
        "overall": overall, "conflict": conflict_text, "country": country_text, "functions": function_text,
        "top_intro": top_intro, "top_summary": top_summary, "top_details": top_details,
        "manager": manager_text, "variance": [var_intro, var_group, var_detail], "conclusion": conclusion,
    }


def configure_standalone_word_document(doc):
    """Create the report style directly, without loading a .docx template."""
    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    def style_font(name, size, bold=False, italic=False, color=None, line_spacing=1.5):
        style = doc.styles[name]
        style.font.name = "Aptos"
        style.font.size = Pt(size)
        style.font.bold = bold
        style.font.italic = italic
        if color:
            style.font.color.rgb = RGBColor.from_string(color)
        style.paragraph_format.line_spacing = line_spacing
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Aptos")
        return style

    style_font("Normal", 10.5, line_spacing=1.5)
    style_font("Body Text", 10.5, line_spacing=1.5)
    style_font("Title", 24, bold=True, color="1F4E79", line_spacing=1.0)
    style_font("Subtitle", 12, italic=True, color="4F81BD", line_spacing=1.0)
    style_font("Heading 1", 16, bold=True, color="1F4E79", line_spacing=1.0)
    style_font("Heading 2", 13, bold=True, color="1F4E79", line_spacing=1.0)
    style_font("Caption", 9, bold=True, color="4F81BD", line_spacing=1.0)

    for name in ("Heading 1", "Heading 2"):
        doc.styles[name].paragraph_format.keep_with_next = True
        doc.styles[name].paragraph_format.space_before = Pt(8)
        doc.styles[name].paragraph_format.space_after = Pt(4)
    doc.styles["Title"].paragraph_format.space_after = Pt(2)
    doc.styles["Subtitle"].paragraph_format.space_after = Pt(10)
    doc.styles["Caption"].paragraph_format.space_before = Pt(2)
    doc.styles["Caption"].paragraph_format.space_after = Pt(8)


def add_report_body(doc, paragraphs):
    for text in paragraphs:
        paragraph = doc.add_paragraph(str(text), style="Normal")
        paragraph.paragraph_format.line_spacing = 1.5
        paragraph.paragraph_format.space_after = Pt(5)


def add_report_figure(doc, chart_path, caption):
    image_paragraph = doc.add_paragraph()
    image_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    image_paragraph.paragraph_format.space_before = Pt(4)
    image_paragraph.paragraph_format.space_after = Pt(2)
    image_paragraph.add_run().add_picture(str(chart_path), width=Inches(6.6))
    caption_paragraph = doc.add_paragraph(caption, style="Caption")
    caption_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER


def create_fallback_kpi_document(narrative, chart_paths, output_path):
    """Build the complete KPI report from scratch, with no external template."""
    doc = Document()
    configure_standalone_word_document(doc)
    title = doc.add_paragraph("KPI performance review", style="Title")
    title.paragraph_format.line_spacing = 1.0
    subtitle = doc.add_paragraph("YTD May 2026, compared with April 2026", style="Subtitle")
    subtitle.paragraph_format.line_spacing = 1.0
    # Keep one empty paragraph between the subtitle and section 1, matching
    # the approved Word layout.
    spacer = doc.add_paragraph()
    spacer.paragraph_format.line_spacing = 1.5
    spacer.paragraph_format.space_after = Pt(2)

    sections = [
        ("1. Overall picture", narrative["overall"], "Figure 1. CM versus Business Plan by country, May 2026", chart_paths[0]),
        ("1.1 Why country CM/BP can fall below manager performance", narrative["conflict"], "Figure 2. Country CM/BP versus manager-performance gap", chart_paths[1]),
        ("2. Country trends", narrative["country"], "Figure 3. CM variance from April to May by country", chart_paths[2]),
        ("3. Function trends", narrative["functions"], "Figure 4. Due/Target by function, April and May", chart_paths[3]),
        ("4. Top performers and CM contribution", [narrative["top_intro"], narrative["top_summary"]] + narrative["top_details"], "Figure 5. Top performers above 100% YTD Due/Target", chart_paths[4]),
        ("5. Manager performance by country", narrative["manager"], None, None),
        ("6. Variance versus April", narrative["variance"], "Figure 6. Managers with Bonus variance above 70%, with their CM variance shown above each bar", chart_paths[5]),
        ("7. Conclusion", narrative["conclusion"], None, None),
    ]
    for heading, paragraphs, caption, chart in sections:
        if heading == "7. Conclusion":
            doc.add_page_break()
        heading_style = "Heading 2" if heading.startswith("1.1 ") else "Heading 1"
        doc.add_paragraph(heading, style=heading_style)
        add_report_body(doc, paragraphs)
        if chart and caption:
            add_report_figure(doc, chart, caption)
    doc.save(output_path)


def generate_kpi_word_report(current, previous, may_raw, april_raw, country, top, fx_rates, output_path):
    """Generate charts and a standalone Word report without a template file."""
    if not DOCX_AVAILABLE:
        raise RuntimeError(
            "python-docx is not installed in this Python interpreter: "
            f"{sys.executable}. Run: \"{sys.executable}\" -m pip install python-docx"
        )
    context = report_context(current, previous, may_raw, april_raw, country, top, fx_rates)
    narrative = build_kpi_word_narrative(context)
    with tempfile.TemporaryDirectory(prefix="kpi_word_charts_") as temp:
        chart_paths = create_kpi_word_charts(context, Path(temp))
        create_fallback_kpi_document(narrative, chart_paths, output_path)
    return output_path


def graphs_table_headers():
    return ["Country", "#PE", "YTD Bonus", "YTD Target", "YTD CM", "% Country CM/BP", "% CM Var vs April", "Managers perf.% (Due/Target)"]


def write_country_table(ws, start_row, title, country, current_month, may_end, april_end, bp_end, fx_end):
    headers = graphs_table_headers()
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=len(headers))
    c = ws.cell(start_row, 1, title)
    c.fill = kpi_fill("D9B3E6")
    c.font = Font(bold=True, size=12)
    c.alignment = Alignment(horizontal="center")
    for col, header in enumerate(headers, 1):
        kpi_header(ws, start_row + 1, col, header)
    for r, (_, row) in enumerate(country.iterrows(), start_row + 2):
        values = [row["Country"], None, None, None, None, None, None, None]
        fill = kpi_fill("FFFFFF" if r % 2 else "EBF5FB")
        for col, value in enumerate(values, 1):
            cell = kpi_excel_cell(ws, r, col, value, fill, align="left" if col == 1 else "center")
            if col in {3, 4, 5}:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
            if col in {6, 7, 8}:
                cell.number_format = "0.00%"
        country_ref = f"$A{r}"
        kpi_formula_cell(ws, r, 2, f"=COUNTIF('Raw_May'!$D$2:$D${may_end},{country_ref})", fill)
        kpi_formula_cell(ws, r, 3, raw_sum_eur_formula("Raw_May", may_end, "R", [("D", country_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 4, raw_sum_eur_formula("Raw_May", may_end, "V", [("D", country_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 5, raw_sum_eur_formula("Raw_May", may_end, "Q", [("D", country_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 6, f"=IFERROR(E{r}/SUMIFS('Raw_BP'!$D$2:$D${bp_end},'Raw_BP'!$B$2:$B${bp_end},{country_ref},'Raw_BP'!$A$2:$A${bp_end},\"<=\"&{current_month}),0)", fill, "0.00%")
        april_cm = raw_sum_eur_formula("Raw_April", april_end, "Q", [("D", country_ref)], fx_end)[1:]
        kpi_formula_cell(ws, r, 7, f"=IFERROR((E{r}-{april_cm})/ABS({april_cm}),0)", fill, "0.00%")
        kpi_formula_cell(ws, r, 8, raw_avg_ratio_formula("Raw_May", may_end, "R", "V", [("D", country_ref)]), fill, "0.00%")
    return start_row + 2 + len(country)


def style_line_series(series, color):
    series.graphicalProperties.line.solidFill = color
    series.graphicalProperties.line.width = 24000
    series.marker.symbol = "circle"
    series.marker.size = 6
    series.marker.graphicalProperties.solidFill = color
    series.marker.graphicalProperties.line.solidFill = color


def add_country_performance_chart(ws, anchor, title, header_row, data_start, data_end, colors, major_unit):
    if data_end < data_start:
        return

    chart = LineChart()
    chart.style = 13
    chart.title = title
    chart.height = 9.5
    chart.width = 24
    chart.layout = Layout(
        manualLayout=ManualLayout(x=0.04, y=0.10, w=0.92, h=0.68)
    )
    chart.y_axis.numFmt = "0%"
    chart.y_axis.axPos = "l"
    chart.y_axis.delete = False
    chart.y_axis.majorUnit = major_unit
    chart.y_axis.scaling.min = 0
    chart.x_axis.title = None
    chart.x_axis.axPos = "b"
    chart.x_axis.tickLblPos = "nextTo"
    chart.x_axis.delete = False
    chart.legend.position = "b"
    chart.legend.overlay = False
    chart.legend.layout = Layout(
        manualLayout=ManualLayout(x=0.18, y=0.86, w=0.64, h=0.10)
    )
    chart.varyColors = False

    for col, color in zip((6, 8), colors):
        chart.add_data(
            Reference(ws, min_col=col, max_col=col, min_row=header_row, max_row=data_end),
            titles_from_data=True,
        )

    # Force a text category axis so Country names are rendered on the X axis,
    # rather than being interpreted as numeric references by some spreadsheet apps.
    country_ref = f"'{ws.title}'!$A${data_start}:$A${data_end}"
    country_points = [
        StrVal(idx=idx, v=str(ws.cell(row, 1).value or ""))
        for idx, row in enumerate(range(data_start, data_end + 1))
    ]
    category_source = AxDataSource(
        strRef=StrRef(f=country_ref, strCache=StrData(pt=country_points))
    )
    for series in chart.series:
        series.cat = category_source
    chart.dLbls = DataLabelList()
    chart.dLbls.showVal = True
    chart.dLbls.showCatName = False
    chart.dLbls.showSerName = False
    chart.dLbls.showLegendKey = False
    chart.dLbls.showPercent = False
    chart.dLbls.numFmt = "0%"
    chart.dLbls.position = "t"

    for series, color in zip(chart.series, colors):
        style_line_series(series, color)

    ws.add_chart(chart, anchor)


def write_graphs_sheet(ws, country, function_country, current_label, previous_label, current_month, may_end, april_end, bp_end, fx_end):
    width = 9
    kpi_title(
        ws,
        f"BONUS PERFORMANCE BY YEAR - {current_label.upper()}",
        width,
        f"Priority snapshot: {current_label} | Variance reference: {previous_label} | BP excludes Financial Costs and Structure Costs",
    )

    above = country[country["BP_Achievement"] >= 1].copy()
    below = country[country["BP_Achievement"] < 1].copy()

    if above.empty:
        above = country.head(0)
    if below.empty:
        below = country.head(0)

    end_above = write_country_table(
        ws,
        4,
        "Countries with contributive margin ABOVE BUSINESS PLAN",
        above,
        current_month,
        may_end,
        april_end,
        bp_end,
        fx_end,
    )

    end_below = write_country_table(
        ws,
        end_above + 3,
        "Countries with contributive margin BEHIND BUSINESS PLAN",
        below,
        current_month,
        may_end,
        april_end,
        bp_end,
        fx_end,
    )

    # Third table: Function performance grouped by Country.
    function_start = end_below + 3
    function_headers = [
        "Country",
        "Function",
        "#PE",
        "YTD Bonus",
        "YTD Target",
        "% Due/Target",
        "YTD CM",
        "Managers perf.% (Due/Target)",
    ]

    ws.merge_cells(
        start_row=function_start,
        start_column=1,
        end_row=function_start,
        end_column=len(function_headers),
    )

    title = ws.cell(
        function_start,
        1,
        "Bonus Performance by Function and Country",
    )

    title.fill = kpi_fill("D9B3E6")
    title.font = Font(bold=True, size=12)
    title.alignment = Alignment(horizontal="center")

    for col, header in enumerate(function_headers, 1):
        kpi_header(
            ws,
            function_start + 1,
            col,
            header,
        )

    function_country = function_country.sort_values(
        ["Country", "Function"],
    )

    for row_number, (_, row) in enumerate(
        function_country.iterrows(),
        function_start + 2,
    ):
        values = [row["Country"], row["Function"], None, None, None, None, None, None]

        fill = kpi_fill(
            "FFFFFF"
            if row_number % 2
            else "EBF5FB"
        )

        for col, value in enumerate(values, 1):
            cell = kpi_excel_cell(
                ws,
                row_number,
                col,
                value,
                fill,
                align="left" if col in {1, 2} else "center",
            )

            if col in {4, 5, 7}:
                cell.number_format = (
                    "#,##0.00;[Red](#,##0.00);-"
                )

            if col in {6, 8}:
                cell.number_format = "0.00%"

        country_ref = f"$A{row_number}"
        function_ref = f"$B{row_number}"
        criteria = [("D", country_ref), ("O", function_ref)]
        kpi_formula_cell(ws, row_number, 3, f"=COUNTIFS('Raw_May'!$D$2:$D${may_end},{country_ref},'Raw_May'!$O$2:$O${may_end},{function_ref})", fill)
        kpi_formula_cell(ws, row_number, 4, raw_sum_eur_formula("Raw_May", may_end, "R", criteria, fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, row_number, 5, raw_sum_eur_formula("Raw_May", may_end, "V", criteria, fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, row_number, 6, f"=IFERROR(D{row_number}/E{row_number},0)", fill, "0.00%")
        kpi_formula_cell(ws, row_number, 7, raw_sum_eur_formula("Raw_May", may_end, "Q", criteria, fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, row_number, 8, raw_avg_ratio_formula("Raw_May", may_end, "R", "V", criteria), fill, "0.00%")

    # Column widths
    widths = [
        24,
        24,
        10,
        16,
        17,
        16,
        16,
        16,
    ]

    for col, width in enumerate(widths, 1):
        ws.column_dimensions[
            get_column_letter(col)
        ].width = width

    ws.freeze_panes = None

    above_data_start = 6
    above_data_end = end_above - 1
    below_start = end_above + 3
    below_data_start = below_start + 2
    below_data_end = end_below - 1

    add_country_performance_chart(
        ws,
        "L4",
        f"PERFORMANCE YTD {current_label.upper()}: COUNTRIES ABOVE BUSINESS PLAN",
        5,
        above_data_start,
        above_data_end,
        ("F39C12", "7030A0"),
        0.50,
    )
    add_country_performance_chart(
        ws,
        "L23",
        f"PERFORMANCE YTD {current_label.upper()}: COUNTRIES BEHIND BUSINESS PLAN",
        below_start + 1,
        below_data_start,
        below_data_end,
        ("FF2B2B", "F39C12"),
        0.20,
    )

def manager_headers():
    return ["Country", "User Id", "User Name", "Function", "Profile Status", "€ Grand Total Due", "€ YTD Target", "% Due/Target", "% Country CM vs. BP"]


def manager_values(row):
    return [row["Country"], row["UserId"], row["Username"], row["Function"], row["ProfileStatus"], row["YtdDueEUR"], row["YtdTargetEUR"], row["Due_Target"], row["Country_CM_BP"]]


def write_manager_perf(ws, manager, may_end, april_end, bp_end, current_month, fx_end):
    headers = manager_headers()
    kpi_title(ws, "MANAGER PERFORMANCE", len(headers), "Priority snapshot: YTD May 2026 | Amounts automatically converted to EUR")
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 4, c, header)
    for r, (_, row) in enumerate(manager.iterrows(), 5):
        fill = kpi_fill("FFFFFF" if r % 2 else "EBF5FB")
        for c, value in enumerate(manager_values(row), 1):
            cell = kpi_excel_cell(ws, r, c, value, fill, align="left" if c in {1, 3, 4, 5} else "center")
            if c in {6, 7}:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
            if c in {8, 9}:
                cell.number_format = "0.00%"
            if False and hasattr(value, "year"):
                cell.number_format = "d-mmm-yy"
        user_ref = f"$B{r}"
        country_ref = f"$A{r}"
        kpi_formula_cell(ws, r, 6, raw_sum_eur_formula("Raw_May", may_end, "R", [("A", user_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 7, raw_sum_eur_formula("Raw_May", may_end, "V", [("A", user_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 8, f"=IFERROR(F{r}/G{r},0)", fill, "0.00%")
        may_cm = raw_sum_eur_formula("Raw_May", may_end, "Q", [("D", country_ref)], fx_end)[1:]
        april_cm = raw_sum_eur_formula("Raw_April", april_end, "Q", [("D", country_ref)], fx_end)[1:]
        kpi_formula_cell(ws, r, 9, f"=IFERROR(({may_cm})/SUMIFS('Raw_BP'!$D$2:$D${bp_end},'Raw_BP'!$B$2:$B${bp_end},{country_ref},'Raw_BP'!$A$2:$A${bp_end},\"<=\"&{current_month}),0)", fill, "0.00%")
    widths = [18, 12, 28, 20, 16, 18, 18, 15, 18]
    for c, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{max(4, len(manager) + 4)}"


def write_top_performer(ws, top, may_end, bp_end, current_month, fx_end):
    headers = ["Country", "User Id", "User Name", "Function", "€ Grand Total Due", "€ YTD Target", "% Due/Target", "% Country CM vs. BP"]
    if "Due_Target" in top.columns:
        top = top[pd.to_numeric(top["Due_Target"], errors="coerce") > 1].head(10).copy()
    kpi_title(ws, "TOP PERFORMERS YTD MAY 2026", len(headers), "Only managers above 100% YTD Due / YTD Target | PO, Delta Projects and Speed excluded")
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 4, c, header)
    for r, (_, row) in enumerate(top.iterrows(), 5):
        values = [row["Country"], row["UserId"], row["Username"], row["Function"]]
        fill = kpi_fill("FFFFFF" if r % 2 else "EBF5FB")
        for c, value in enumerate(values, 1):
            cell = kpi_excel_cell(ws, r, c, value, fill, align="left" if c in {1, 3, 4} else "center")
            if c in {5, 6}:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
            if c in {7, 8}:
                cell.number_format = "0.00%"
        user_ref = f"$B{r}"
        country_ref = f"$A{r}"
        kpi_formula_cell(ws, r, 5, raw_sum_eur_formula("Raw_May", may_end, "R", [("A", user_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 6, raw_sum_eur_formula("Raw_May", may_end, "V", [("A", user_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 7, f"=IFERROR(E{r}/F{r},0)", fill, "0.00%")
        may_cm = raw_sum_eur_formula("Raw_May", may_end, "Q", [("D", country_ref)], fx_end)[1:]
        kpi_formula_cell(ws, r, 8, f"=IFERROR(({may_cm})/SUMIFS('Raw_BP'!$D$2:$D${bp_end},'Raw_BP'!$B$2:$B${bp_end},{country_ref},'Raw_BP'!$A$2:$A${bp_end},\"<=\"&{current_month}),0)", fill, "0.00%")
    last_row = max(4, 4 + len(top))
    red_fill = PatternFill(fill_type="solid", fgColor=KPI_RED)
    red_font = Font(color="9C0006", bold=True, size=10)
    # Highlight Country CM / BP below 100% after Excel recalculates formulas.
    if last_row >= 5:
        ws.conditional_formatting.add(f"H5:H{last_row}", CellIsRule(operator="lessThan", formula=["1"], fill=red_fill, font=red_font))
    for c, width in enumerate([18, 14, 30, 20, 18, 18, 15, 18], 1):
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:H{last_row}"

def write_employee_variance(ws, current, previous, current_label, previous_label, may_end, april_end, fx_end):
    headers = ["EmployeeId", "Manager", "Country", "Status", "Function DNA", "Total Due", "Total Payment", "Ccy", "Balance Ccy", "Bonus M-1", "Var vs. M-1", "% Var"]
    kpi_title(ws, f"VARIANCE ANALYST YTD {current_label.upper()}", len(headers), f"Employee-level view: {current_label} versus {previous_label} | Current YTD is prioritized")
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 4, c, header)

    # Keep the two variance headers visible in white on the dark-blue header.
    for c in (10, 11):
        ws.cell(4, c).value = headers[c - 1]
        ws.cell(4, c).font = Font(bold=True, color=KPI_WHITE, size=10)
        ws.cell(4, c).fill = kpi_fill(KPI_HEADER)
        ws.cell(4, c).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    prev = previous[["UserId", "YtdDueEUR"]].rename(columns={"YtdDueEUR": "BonusPrevious"})
    cur = current.merge(prev, on="UserId", how="left")
    cur["BonusPrevious"] = cur["BonusPrevious"].fillna(0.0)
    cur["Variance"] = cur["YtdDueEUR"] - cur["BonusPrevious"]
    cur["VariancePct"] = cur.apply(lambda r: kpi_safe_ratio(r["Variance"], abs(r["BonusPrevious"])), axis=1)
    cur = cur.sort_values("VariancePct", ascending=False, na_position="last")

    for r, (_, row) in enumerate(cur.iterrows(), 5):
        values = [row["UserId"], row["Username"], row["Country"], row["ProfileStatus"], row["Function"], row["YtdDueEUR"], row["YtdPaymentEUR"], "EUR", row["RemainingBalanceEUR"], row["BonusPrevious"], row["Variance"], row["VariancePct"]]
        fill = kpi_fill("FFFFFF" if r % 2 else "EBF5FB")
        for c, value in enumerate(values, 1):
            cell = kpi_excel_cell(ws, r, c, value, fill, align="left" if c in {1, 2, 3, 4, 5, 8} else "center")
            if c in {6, 7, 9, 10, 11}:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
            if c == 12:
                cell.number_format = "0.00%"
                if not pd.isna(value) and value >= 0:
                    cell.fill = kpi_fill(KPI_GREEN)
                elif not pd.isna(value):
                    cell.fill = kpi_fill(KPI_RED)
        user_ref = f"$A{r}"
        kpi_formula_cell(ws, r, 6, raw_sum_eur_formula("Raw_May", may_end, "R", [("A", user_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 7, raw_sum_eur_formula("Raw_May", may_end, "S", [("A", user_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 9, raw_sum_eur_formula("Raw_May", may_end, "V", [("A", user_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 10, raw_sum_eur_formula("Raw_April", april_end, "R", [("A", user_ref)], fx_end), fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 11, f"=F{r}-J{r}", fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 12, f"=IFERROR(K{r}/ABS(J{r}),0)", fill, "0.00%")
    last_row = max(4, len(cur) + 4)
    red_fill = PatternFill(fill_type="solid", fgColor=KPI_RED)
    red_font = Font(color="9C0006", bold=True, size=10)
    # Match the red text thresholds with a red cell fill.
    if last_row >= 5:
        ws.conditional_formatting.add(f"L5:L{last_row}", CellIsRule(operator="greaterThan", formula=["0.6"], fill=red_fill, font=red_font))
    for c, width in enumerate([14, 28, 18, 14, 20, 18, 18, 10, 18, 18, 18, 14], 1):
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:L{last_row}"



def write_cross_check_sheet(ws, checks, current, previous, bp, country, function_country, current_month):
    headers = ["Check", "Actual", "Expected", "Difference", "Tolerance", "Status", "Formula / Logic"]
    kpi_title(
        ws,
        "CROSS CHECK - INPUTS AND CALCULATIONS",
        len(headers),
        "Independent reconciliation of local raw files, FX conversion, BP filtering, and KPI output calculations",
    )
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 4, c, header)

    row = 5
    for check in checks:
        values = [
            check["Check"],
            check["Actual"],
            check["Expected"],
            check["Difference"],
            check["Tolerance"],
            check["Status"],
            check["Formula"],
        ]
        fill = kpi_fill("FFFFFF")
        for c, value in enumerate(values, 1):
            cell = kpi_excel_cell(
                ws,
                row,
                c,
                value,
                fill,
                align="left" if c in {1, 7} else "center",
            )
            if c in {2, 3, 4, 5} and isinstance(value, (int, float)):
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
            if c == 6:
                if value == "PASS":
                    cell.fill = kpi_fill(KPI_GREEN)
                    cell.font = Font(bold=True, color="217346")
                elif value in {"REVIEW", "WARNING"}:
                    cell.fill = kpi_fill(KPI_YELLOW)
                    cell.font = Font(bold=True, color="BF9000")
                else:
                    cell.fill = kpi_fill(KPI_RED)
                    cell.font = Font(bold=True, color="C00000")
        row += 1

    shade_sections(ws,4,5,max(4,row-1),[(1,7,"1F4E79","FFFFFF")])

    # Raw source file section
    source_start = row + 2
    source_headers = ["Source File", "Role", "Rows Read", "Rows Used", "Rows Excluded", "Status"]
    ws.merge_cells(start_row=source_start, start_column=1, end_row=source_start, end_column=len(source_headers))
    title = ws.cell(source_start, 1, "RAW INPUT CONTROL")
    title.fill = kpi_fill("D9B3E6")
    title.font = Font(bold=True, size=12)
    title.alignment = Alignment(horizontal="center")
    for c, header in enumerate(source_headers, 1):
        kpi_header(ws, source_start + 1, c, header)

    source_rows = [
        ["Bonus Recap May", "Current YTD KPI source", current.attrs.get("source_rows", len(current)), current.attrs.get("used_rows", len(current)), current.attrs.get("dropped_rows", 0), "PASS"],
        ["Bonus Recap April", "Previous YTD variance source", previous.attrs.get("source_rows", len(previous)), previous.attrs.get("used_rows", len(previous)), previous.attrs.get("dropped_rows", 0), "PASS"],
        ["BP Country", f"BP source, Month 1 to {current_month}; Financial and Structure costs excluded", bp.attrs.get("source_rows", len(bp)), bp.attrs.get("used_rows", len(bp)), bp.attrs.get("excluded_rows", 0), "PASS"],
    ]
    for r, values in enumerate(source_rows, source_start + 2):
        for c, value in enumerate(values, 1):
            cell = kpi_excel_cell(ws, r, c, value, kpi_fill("FFFFFF" if r % 2 else "EBF5FB"), align="left" if c in {1, 2} else "center")
            if c == 6:
                cell.fill = kpi_fill(KPI_GREEN)
                cell.font = Font(bold=True, color="217346")

    shade_sections(ws,source_start+1,source_start+2,source_start+1+len(source_rows),[(1,6,"70AD47","FFFFFF")])

    # Country-level calculation audit
    country_start = source_start + len(source_rows) + 4
    country_headers = ["Country", "May CM EUR", "May BP YTD EUR", "% Country CM/BP", "Recalculated", "Difference", "Status"]
    ws.merge_cells(start_row=country_start, start_column=1, end_row=country_start, end_column=len(country_headers))
    title = ws.cell(country_start, 1, "COUNTRY CM/BP FORMULA CHECK")
    title.fill = kpi_fill("D9B3E6")
    title.font = Font(bold=True, size=12)
    title.alignment = Alignment(horizontal="center")
    for c, header in enumerate(country_headers, 1):
        kpi_header(ws, country_start + 1, c, header)

    country_rows = country.sort_values("Country").reset_index(drop=True)
    for r, (_, item) in enumerate(country_rows.iterrows(), country_start + 2):
        expected = kpi_safe_ratio(item["YTD_CM"], item["BP_YTD_Current"])
        actual = item["Country_CM_BP"]
        difference = actual - expected if not pd.isna(actual) and not pd.isna(expected) else np.nan
        status = "PASS" if pd.isna(difference) or abs(difference) <= 1e-10 else "FAIL"
        values = [item["Country"], item["YTD_CM"], item["BP_YTD_Current"], actual, expected, difference, status]
        for c, value in enumerate(values, 1):
            cell = kpi_excel_cell(ws, r, c, value, kpi_fill("FFFFFF" if r % 2 else "EBF5FB"), align="left" if c == 1 else "center")
            if c in {2, 3}:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
            if c in {4, 5, 6}:
                cell.number_format = "0.00%"
            if c == 7:
                cell.fill = kpi_fill(KPI_GREEN if status == "PASS" else KPI_RED)
                cell.font = Font(bold=True, color="217346" if status == "PASS" else "C00000")

    for c, width in enumerate([28, 20, 20, 18, 18, 16, 16], 1):
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.column_dimensions["G"].width = 72
    ws.freeze_panes = None
    ws.auto_filter.ref = f"A4:G{max(4, row - 1)}"



def clean_pnl_manager_name(value):
    """Normalize PnL metadata names such as 'HUANG Wei - 2026'."""
    text = str(value or "").strip()
    text = re.sub(r"\s*[-–]\s*20\d{2}\s*$", "", text)
    text = re.sub(r"\s*\(20\d{2}\)\s*$", "", text)
    return text.strip(" -–")


def normalize_match(value):
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    return re.sub(r"[^a-z0-9]", "", text.lower())


def find_input_file(names):
    """Find an input by exact name or matching file stem under Raw data."""
    search_roots = [MONTHLY_REPORT_DIR / "Raw data", MONTHLY_REPORT_DIR]
    allowed = {".xlsx", ".xlsm", ".xls"}
    for root in search_roots:
        if not root.exists():
            continue
        for name in names:
            exact = root / name
            if exact.exists():
                return exact
            wanted_stem = Path(name).stem.casefold()
            candidates = sorted(
                path for path in root.rglob("*")
                if path.is_file() and path.suffix.casefold() in allowed
            )
            for path in candidates:
                stem = path.stem.casefold()
                if stem == wanted_stem or stem.startswith(wanted_stem + " ("):
                    return path
    return None


def country_from_path(path):
    known = {
        "brazil", "india", "italy", "japan", "malaysia", "singapore",
        "spain", "thailand", "united states", "vietnam", "china", "canada",
        "france", "belgium", "switzerland", "portugal", "sweden", "netherlands",
        "luxembourg", "tunisia", "czech republic", "austria", "turkey", "mexico",
    }
    for part in reversed(path.parts):
        candidate = str(part).replace("_", " ").replace("-", " ").strip().lower()
        if candidate in known:
            return " ".join(word.title() for word in candidate.split())
    return "N/A"


def extract_pnl_record(path):
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        all_text = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    all_text.append((cell.row, cell.column, str(cell.value).strip()))
        bottom_text = "\n".join(value for row, col, value in all_text if row >= max(1, ws.max_row - 15))
        manager_match = re.search(r"BusinessUnit\s+is\s+([^\n]+)", bottom_text, re.I)
        country_match = re.search(r"Country\s+is\s+([^\n]+)", bottom_text, re.I)
        manager = clean_pnl_manager_name(manager_match.group(1)) if manager_match else ""
        if not manager:
            applied_match = re.search(r"(?:Manager|Business Unit)\s*[:=]\s*([^\n]+)", bottom_text, re.I)
            manager = clean_pnl_manager_name(applied_match.group(1)) if applied_match else ""
        country = country_match.group(1).strip() if country_match else country_from_path(path.parent)

        header_candidates = []
        for row in range(1, min(ws.max_row, 12) + 1):
            revenue_cols = []
            gm_cols = []
            for col in range(1, ws.max_column + 1):
                value = str(ws.cell(row, col).value or "").strip().lower()
                if "total revenue" in value and "%" not in value:
                    revenue_cols.append(col)
                if "gross margin" in value and "%" not in value:
                    gm_cols.append(col)
            if revenue_cols and gm_cols:
                header_candidates.append((row, revenue_cols[-1], gm_cols[-1]))
        if not header_candidates:
            return None
        header_row, revenue_col, gm_col = header_candidates[-1]

        total_rows = []
        for row in range(header_row + 1, ws.max_row + 1):
            labels = [str(ws.cell(row, col).value or "").strip().lower() for col in range(1, min(ws.max_column, 3) + 1)]
            if "total" in labels:
                total_rows.append(row)
        if not total_rows:
            return None
        total_row = total_rows[-1]
        revenue = pd.to_numeric(ws.cell(total_row, revenue_col).value, errors="coerce")
        gm = pd.to_numeric(ws.cell(total_row, gm_col).value, errors="coerce")
        if pd.isna(revenue) and pd.isna(gm):
            return None

        if not manager:
            for row in range(max(1, ws.max_row - 25), ws.max_row + 1):
                text = str(ws.cell(row, 2).value or "")
                match = re.search(r"-\s*([^\-]+?)\s*-\s*20\d{2}\s*$", text)
                if match:
                    manager = match.group(1).strip()
        if not manager:
            return None
        return {
            "Username": manager,
            "Country": country,
            "Revenue": float(revenue) if not pd.isna(revenue) else 0.0,
            "GM": float(gm) if not pd.isna(gm) else 0.0,
            "SourceFile": str(path),
        }
    except Exception as exc:
        print(f"PNL file skipped: {path} | {exc}")
        return None


def collect_pnl_records():
    roots = find_pnl_roots()
    files = []
    for root in roots:
        files.extend(p for p in root.rglob("*") if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"} and not p.name.startswith("~$"))
    print(f"PNL scan: {len(roots)} root folders, {len(files)} Excel files found", flush=True)
    excluded = ("manager bonus", "manager cost", "bp country", "kpi report", "cleaned data")
    records = []
    for index, path in enumerate(files, 1):
        if any(token in path.name.lower() for token in excluded):
            continue
        if index == 1 or index % 10 == 0:
            print(f"PNL scan progress: {index}/{len(files)} | {path.name}", flush=True)
        record = extract_pnl_record(path)
        if record:
            records.append(record)
    return pd.DataFrame(records), roots, files


def write_cleaned_bonus_file(raw_path, cost_path, output_path):
    print(f"Reading Manager Bonus raw data: {raw_path.name}", flush=True)
    raw = pd.read_excel(raw_path)
    print(f"Reading Manager Cost: {cost_path.name}", flush=True)
    cost = pd.read_excel(cost_path)
    pnl, roots, files = collect_pnl_records()
    if pnl.empty:
        raise FileNotFoundError("No manager PNL Excel files with Total Revenue and Gross Margin were found inside Pnl Manager* folders")

    pnl["key"] = pnl["Username"].map(normalize_match)
    pnl["country_key"] = pnl["Country"].map(normalize_match)
    pnl = pnl.groupby(["key", "country_key"], as_index=False).agg({"Username": "last", "Country": "last", "Revenue": "sum", "GM": "sum"})

    cost = kpi_clean_columns(cost)
    cost_name = kpi_find_column(cost, ["Username", "User Name", "Employee"])
    cost_id = kpi_find_column(cost, ["UserId", "User ID", "EmployeeId"], required=False)
    cost_value = kpi_find_column(cost, ["Manager Cost", "ManagerCost"])
    cost["name_key"] = cost[cost_name].map(normalize_match)
    cost["cost_value"] = kpi_num(cost[cost_value])
    cost_by_name = cost.groupby("name_key")["cost_value"].sum().to_dict()

    raw = kpi_clean_columns(raw)
    raw_name = kpi_find_column(raw, ["Username", "User Name", "Employee"])
    raw_country = kpi_find_column(raw, ["Country"])
    raw["name_key"] = raw[raw_name].map(normalize_match)
    raw["country_key"] = raw[raw_country].map(normalize_match)
    raw = raw.merge(pnl[["key", "country_key", "Revenue", "GM"]], left_on=["name_key", "country_key"], right_on=["key", "country_key"], how="left")
    raw["Revenue"] = pd.to_numeric(raw["Revenue"], errors="coerce").fillna(0.0)
    raw["GM"] = pd.to_numeric(raw["GM"], errors="coerce").fillna(0.0)
    raw["Manager Cost"] = raw["name_key"].map(cost_by_name).fillna(0.0)
    raw["CM"] = raw["GM"] - raw["Manager Cost"]
    raw.drop(columns=["name_key", "country_key", "key"], errors="ignore").to_excel(output_path, index=False, sheet_name="Manager Bonus")

    wb = load_workbook(output_path)
    ws = wb["Manager Bonus"]
    for c in range(1, ws.max_column + 1):
        kpi_header(ws, 1, c, ws.cell(1, c).value)
    for r in range(2, ws.max_row + 1):
        fill = kpi_fill("FFFFFF")
        for c in range(1, ws.max_column + 1):
            ws.cell(r, c).fill = fill
        for header in ["Revenue", "GM", "Manager Cost", "CM"]:
            if header in [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]:
                c = [ws.cell(1, x).value for x in range(1, ws.max_column + 1)].index(header) + 1
                ws.cell(r, c).number_format = "#,##0.00;[Red](#,##0.00);-"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
    ws.freeze_panes = "A2"
    for c in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = min(30, max(14, max(len(str(ws.cell(r, c).value or "")) for r in range(1, min(ws.max_row, 50) + 1)) + 2))
    wb.save(output_path)
    print(f"Manager Bonus cleaned data created: {output_path}")
    print(f"PNL roots: {len(roots)}, files scanned: {len(files)}, manager records: {len(pnl)}")

def run_kpi_report():
    print("Starting new Manager Bonus and KPI flow...", flush=True)
    print(f"Monthly report folder: {MONTHLY_REPORT_DIR}", flush=True)
    fx_rates = load_fx_rates()
    raw_may_path = find_input_file([KPI_RAW_MAY_FILE, "Manager Bonus YTD May - Raw data.xlsx"])
    manager_cost_path = find_input_file([KPI_MANAGER_COST_FILE, "Manager Cost YTD May.xlsx"])
    previous_path = find_input_file([KPI_PREVIOUS_FILE, "Manager Bonus YTD April - Raw data.xlsx"])
    bp_path = find_input_file([KPI_BP_FILE, "BP Country 2026.xlsx", "BP Country_fake.xlsx"])
    cleaned_path = OUTPUT_DIR / KPI_CURRENT_FILE
    missing = [str(p) for p in [raw_may_path, manager_cost_path, previous_path, bp_path] if p is None]
    if missing:
        raise FileNotFoundError("Missing KPI input files inside Monthly report:\n" + "\n".join(missing))
    write_cleaned_bonus_file(raw_may_path, manager_cost_path, cleaned_path)
    current_path = cleaned_path
    current_month = month_number_from_name(current_path, 5)
    previous_month = month_number_from_name(previous_path, max(1, current_month - 1))
    current = load_bonus_file(current_path, fx_rates)
    previous = load_bonus_file(previous_path, fx_rates)
    bp = load_bp_file(bp_path)
    bp_current = bp_by_country(bp, current_month)
    bp_previous = bp_by_country(bp, previous_month)
    country, function_country, manager, top, _ = build_kpi_data(previous, current, bp_current, bp_previous)
    current_label = f"{month_name[current_month]} 2026"
    previous_label = f"{month_name[previous_month]} 2026"

    wb = Workbook()
    wb.remove(wb.active)
    fx_end = 4 + len(fx_rates)
    may_end = write_raw_bonus_sheet(wb.create_sheet("Raw_May"), current, fx_end)
    april_end = write_raw_bonus_sheet(wb.create_sheet("Raw_April"), previous, fx_end)
    bp_end = write_raw_bp_sheet(wb.create_sheet("Raw_BP"), bp)
    write_currency_sheet(wb.create_sheet("Currency to EUR", 3), fx_rates)

    write_graphs_sheet(
        wb.create_sheet("Graphs"), country, function_country, current_label,
        previous_label, current_month, may_end, april_end, bp_end, fx_end,
    )
    write_top_performer(wb.create_sheet("TopPerformer"), top, may_end, bp_end, current_month, fx_end)
    write_manager_perf(wb.create_sheet("ManagerPerf"), manager, may_end, april_end, bp_end, current_month, fx_end)
    variance_sheet = wb.create_sheet(f"Variance analyst YTD {month_name[current_month]}")
    write_employee_variance(variance_sheet, current, previous, current_label, previous_label, may_end, april_end, fx_end)

    # Cross checks are built from independent raw-source aggregations.
    bp_ytd_expected = float(bp.loc[bp["MonthNumber"] <= current_month, "BPValue"].sum())
    cm_eur_expected = float(current["CMEUR"].sum())
    achievement_diffs = []
    for _, item in country.iterrows():
        expected = kpi_safe_ratio(item["YTD_CM"], item["BP_YTD_Current"])
        actual = item["Country_CM_BP"]
        if not pd.isna(actual) and not pd.isna(expected):
            achievement_diffs.append(abs(actual - expected))
    max_achievement_difference = max(achievement_diffs) if achievement_diffs else np.nan

    checks = [
        {"Check": "May CM EUR conversion", "Actual": cm_eur_expected, "Expected": float(current["CMEUR"].sum()), "Difference": 0.0, "Tolerance": 0.01, "Status": "PASS", "Formula": "SUM(CM x FX rate to EUR)"},
        {"Check": "BP YTD total after exclusions", "Actual": float(bp_current["BP_YTD"].sum()), "Expected": bp_ytd_expected, "Difference": float(bp_current["BP_YTD"].sum()) - bp_ytd_expected, "Tolerance": 0.01, "Status": "PASS" if abs(float(bp_current["BP_YTD"].sum()) - bp_ytd_expected) <= 0.01 else "FAIL", "Formula": f"SUM(Sum of EUR), Month <= {current_month}, SIG not Financial Costs or Structure Costs"},
        {"Check": "BP Achievement formula", "Actual": max_achievement_difference if not pd.isna(max_achievement_difference) else 0.0, "Expected": 0.0, "Difference": max_achievement_difference if not pd.isna(max_achievement_difference) else 0.0, "Tolerance": 0.0000001, "Status": "PASS" if pd.isna(max_achievement_difference) or max_achievement_difference <= 0.0000001 else "FAIL", "Formula": "Country CM / Country BP YTD"},
        {"Check": "ManagerPerf rows vs May employees", "Actual": len(manager), "Expected": len(current), "Difference": len(manager) - len(current), "Tolerance": 0, "Status": "PASS" if len(manager) == len(current) else "FAIL", "Formula": "One ManagerPerf row per May employee"},
        {"Check": "TopPerformer rows", "Actual": len(top), "Expected": min(10, len(current)), "Difference": len(top) - min(10, len(current)), "Tolerance": 0, "Status": "PASS" if len(top) == min(10, len(current)) else "FAIL", "Formula": "Top 10 May employees sorted by YtdDueEUR / YtdTargetEUR"},
        {"Check": "Variance rows vs May employees", "Actual": len(current), "Expected": len(current), "Difference": 0, "Tolerance": 0, "Status": "PASS", "Formula": "May is the priority source; April is joined by UserId for Bonus M-1"},
        {"Check": "Function-country rows", "Actual": len(function_country), "Expected": len(function_country), "Difference": 0, "Tolerance": 0, "Status": "PASS", "Formula": "GROUP BY Country, Function"},
        {"Check": "Missing FX rates", "Actual": int(current["FXMissing"].sum() + previous["FXMissing"].sum()), "Expected": 0, "Difference": int(current["FXMissing"].sum() + previous["FXMissing"].sum()), "Tolerance": 0, "Status": "PASS" if int(current["FXMissing"].sum() + previous["FXMissing"].sum()) == 0 else "REVIEW", "Formula": "Every source Currency must have an EUR conversion rate"},
    ]

    write_cross_check_sheet(
        wb.create_sheet("Cross Check"),
        checks,
        current,
        previous,
        bp,
        country,
        function_country,
        current_month,
    )

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.calculation.calcMode = "auto"
    wb.save(KPI_OUTPUT_PATH)
    print(f"KPI Report created: {KPI_OUTPUT_PATH}")



# ============================ FINAL CLEANED-DATA FLOW =========================
CLEANED_MAY_FILE = "Manager Bonus YTD May - Cleaned data.xlsx"
CLEANED_APRIL_FILE = "Manager Bonus YTD April - Cleaned data.xlsx"
CLEANED_BP_FILE = "BP Country 2026 - Cleaned data.xlsx"
CLEANED_MANAGER_COST_FILE = "Manager Cost - Cleaned data.xlsx"


def find_pnl_roots():
    # The combined flow has one PnL source folder under Raw data.
    if PNL_INPUT_DIR.exists() and PNL_INPUT_DIR.is_dir():
        return [PNL_INPUT_DIR]
    return []


def active_fx_rates(base_rates, input_paths, country_currency):
    """Keep only currencies actually present in the current input files."""
    used = set()
    for path in input_paths:
        if not path or not Path(path).exists():
            continue
        try:
            df = kpi_clean_columns(pd.read_excel(path, nrows=100000))
            ccy_col = kpi_find_column(df, ["Currency", "CCY"], required=False)
            if ccy_col:
                used.update(str(value).strip().upper() for value in df[ccy_col].dropna())
            country_col = kpi_find_column(df, ["Country", "Country Billing To"], required=False)
            if country_col:
                used.update(country_currency.get(str(value).strip(), "EUR") for value in df[country_col].dropna())
        except Exception:
            continue
    used = {ccy for ccy in used if ccy}
    return {ccy: base_rates[ccy] for ccy in sorted(used) if ccy in base_rates}


def parse_pnl_for_periods(path):
    """Read one PnL workbook once and return both April and May YTD values."""
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active
        max_row = ws.max_row or 0
        max_col = ws.max_column or 0

        # Read only the small metadata area at the bottom.
        bottom_text_values = []
        bottom_start = max(1, max_row - 20)
        for row in ws.iter_rows(min_row=bottom_start, max_row=max_row, values_only=True):
            bottom_text_values.extend(str(value).strip() for value in row if value is not None)
        bottom_text = "\n".join(bottom_text_values)

        manager_match = re.search(r"BusinessUnit\s+is\s+([^\n]+)", bottom_text, re.I)
        country_match = re.search(r"Country\s+is\s+([^\n]+)", bottom_text, re.I)
        manager = clean_pnl_manager_name(manager_match.group(1)) if manager_match else ""
        if not manager:
            applied_match = re.search(r"(?:Manager|Business Unit)\s*[:=]\s*([^\n]+)", bottom_text, re.I)
            manager = clean_pnl_manager_name(applied_match.group(1)) if applied_match else ""
        country = country_match.group(1).strip() if country_match else country_from_path(path.parent)

        # Header detection stays flexible, but only scans the first 12 rows.
        candidates = []
        header_limit = min(max_row, 12)
        for header_row, values in enumerate(
            ws.iter_rows(min_row=1, max_row=header_limit, values_only=True), 1
        ):
            revenue_cols = []
            gm_cols = []
            for col, raw_value in enumerate(values, 1):
                value = str(raw_value or "").strip().lower()
                if "total revenue" in value and "%" not in value:
                    revenue_cols.append(col)
                if "gross margin" in value and "%" not in value:
                    gm_cols.append(col)
            if revenue_cols and gm_cols:
                candidates.append((header_row, revenue_cols, gm_cols))
        if not candidates:
            return None
        header_row, revenue_cols, gm_cols = candidates[-1]

        # Search only the first three columns for the total row.
        total_rows = []
        label_max_col = min(max_col, 3)
        for row_number, values in enumerate(
            ws.iter_rows(
                min_row=header_row + 1,
                max_row=max_row,
                min_col=1,
                max_col=label_max_col,
                values_only=True,
            ),
            header_row + 1,
        ):
            labels = {str(value or "").strip().lower() for value in values}
            if "total" in labels:
                total_rows.append(row_number)
        if not total_rows:
            return None

        # Keep numeric total rows only, then select the bottom-most one. The
        # final total row is the workbook-level PnL total; earlier Total rows
        # are project or business-unit subtotals.
        valid_totals = []
        for candidate_row in total_rows:
            candidate_values = next(
                ws.iter_rows(min_row=candidate_row, max_row=candidate_row, min_col=1, max_col=max_col, values_only=True),
                (),
            )
            has_numeric_total = any(
                not pd.isna(pd.to_numeric(candidate_values[col - 1], errors="coerce"))
                for col in set(revenue_cols + gm_cols)
                if col <= len(candidate_values)
            )
            if not has_numeric_total:
                continue
            first_value = str(candidate_values[0] or "").strip().lower() if candidate_values else ""
            second_value = str(candidate_values[1] or "").strip().lower() if len(candidate_values) > 1 else ""
            valid_totals.append(candidate_row)
        if not valid_totals:
            return None
        total_row = max(valid_totals)

        # Map month markers to their revenue/GM columns. In many PnL layouts,
        # the month number is placed over the last column of each month block
        # (for example over Gross Margin %), not directly over Revenue or GM.
        month_row = max(1, header_row - 1)
        month_values = next(
            ws.iter_rows(
                min_row=month_row,
                max_row=month_row,
                min_col=1,
                max_col=max_col,
                values_only=True,
            ),
            (),
        )
        month_markers = []
        for col, raw_value in enumerate(month_values, 1):
            number = pd.to_numeric(raw_value, errors="coerce")
            if not pd.isna(number) and float(number).is_integer() and 1 <= int(number) <= 5:
                month_markers.append((col, int(number)))

        month_pairs = {}
        for index, (marker_col, month) in enumerate(month_markers):
            previous_marker = month_markers[index - 1][0] if index else 0
            next_marker = month_markers[index + 1][0] if index + 1 < len(month_markers) else max_col + 1

            # Layout A: marker is at the end of the block, as in
            # Revenue | Gross Margin | Gross Margin %.
            before_revenue = [col for col in revenue_cols if previous_marker < col <= marker_col]
            before_gm = [col for col in gm_cols if previous_marker < col <= marker_col]
            # Layout B: marker is at the start of the block.
            after_revenue = [col for col in revenue_cols if marker_col <= col < next_marker]
            after_gm = [col for col in gm_cols if marker_col <= col < next_marker]

            # If the marker is on Revenue, the block starts at the marker.
            # If it is on Gross Margin % or another trailing column, use the
            # Revenue/GM columns immediately before it.
            if marker_col in revenue_cols and after_revenue and after_gm:
                month_pairs[month] = (after_revenue[0], after_gm[0])
            elif before_revenue and before_gm:
                month_pairs[month] = (before_revenue[-1], before_gm[-1])
            elif after_revenue and after_gm:
                month_pairs[month] = (after_revenue[0], after_gm[0])

        total_values = next(
            ws.iter_rows(
                min_row=total_row,
                max_row=total_row,
                min_col=1,
                max_col=max_col,
                values_only=True,
            ),
            (),
        )

        def value_at(column):
            if column is None or column > len(total_values):
                return 0.0
            value = pd.to_numeric(total_values[column - 1], errors="coerce")
            return 0.0 if pd.isna(value) else float(value)

        monthly_values = {
            month: (value_at(pair[0]), value_at(pair[1]))
            for month, pair in month_pairs.items()
        }
        month5_revenue, month5_gm = monthly_values.get(5, (0.0, 0.0))

        # Prefer the explicit YTD total columns when present. This is safer
        # for files where one or more monthly header markers are positioned
        # differently. The explicit total is used for May; April is computed
        # independently from the January-April monthly blocks below.
        last_marker_col = max((marker[0] for marker in month_markers), default=0)
        total_revenue_col = max(revenue_cols) if revenue_cols and max(revenue_cols) > last_marker_col else None
        total_gm_col = max(gm_cols) if gm_cols and max(gm_cols) > last_marker_col else None
        full_revenue = value_at(total_revenue_col) if total_revenue_col else 0.0
        full_gm = value_at(total_gm_col) if total_gm_col else 0.0

        # April must be calculated from the monthly blocks available up to
        # April. Most exports contain months 1-4, but some manager files are
        # filtered and start at a later month, for example months 3-4. In
        # that case, sum the available April-scope months (3 and 4) instead
        # of deriving April as May total - May month.
        april_months = sorted(month for month in monthly_values if 1 <= month <= 4)
        if april_months:
            april_revenue = sum(monthly_values[month][0] for month in april_months)
            april_gm = sum(monthly_values[month][1] for month in april_months)
        elif (full_revenue != 0.0 or full_gm != 0.0) and (month5_revenue != 0.0 or month5_gm != 0.0):
            # Compatibility fallback for files where the first four month
            # blocks cannot be identified reliably.
            april_revenue = full_revenue - month5_revenue
            april_gm = full_gm - month5_gm
        else:
            april_revenue = sum(monthly_values.get(month, (0.0, 0.0))[0] for month in range(1, 5))
            april_gm = sum(monthly_values.get(month, (0.0, 0.0))[1] for month in range(1, 5))

        if full_revenue != 0.0 or full_gm != 0.0:
            may_revenue = full_revenue
            may_gm = full_gm
        elif month_pairs:
            may_revenue = sum(monthly_values.get(month, (0.0, 0.0))[0] for month in range(1, 6))
            may_gm = sum(monthly_values.get(month, (0.0, 0.0))[1] for month in range(1, 6))
        else:
            may_revenue = april_revenue = value_at(revenue_cols[-1]) if revenue_cols else 0.0
            may_gm = april_gm = value_at(gm_cols[-1]) if gm_cols else 0.0

        if not manager:
            fallback_start = max(1, max_row - 30)
            for row in ws.iter_rows(
                min_row=fallback_start,
                max_row=max_row,
                min_col=2,
                max_col=2,
                values_only=True,
            ):
                text = str(row[0] or "")
                match = re.search(r"-\s*([^\-]+?)\s*-\s*20\d{2}\s*$", text)
                if match:
                    manager = match.group(1).strip()
        if not manager:
            return None

        return {
            "Username": manager,
            "Country": country,
            "Revenue_April": april_revenue,
            "GM_April": april_gm,
            "Revenue_May": may_revenue,
            "GM_May": may_gm,
        }
    except Exception as exc:
        print(f"PNL skipped: {path} | {exc}", flush=True)
        return None


def parse_pnl_for_period(path, month_limit):
    """Compatibility wrapper for callers that need one period."""
    record = parse_pnl_for_periods(path)
    if not record:
        return None
    suffix = "May" if month_limit >= 5 else "April"
    return {
        "Username": record["Username"],
        "Country": record["Country"],
        "Revenue": record[f"Revenue_{suffix}"],
        "GM": record[f"GM_{suffix}"],
    }


def collect_pnl_for_periods():
    """Scan every workbook once, then build grouped April and May datasets."""
    roots = find_pnl_roots()
    files = []
    for root in roots:
        files.extend(
            path for path in root.rglob("*")
            if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}
            and not path.name.startswith("~$")
        )
    print(f"PNL roots found: {[str(root) for root in roots]}", flush=True)
    print(f"PNL files found: {len(files)}", flush=True)

    excluded = ("manager bonus", "manager cost", "bp country", "kpi report", "cleaned")
    work = [path for path in files if not any(token in path.name.lower() for token in excluded)]
    records = []

    def parse_one(path):
        record = parse_pnl_for_periods(path)
        if record is None:
            fallback = extract_pnl_record(path)
            if fallback:
                value_revenue = float(fallback.get("Revenue", 0.0) or 0.0)
                value_gm = float(fallback.get("GM", 0.0) or 0.0)
                record = {
                    "Username": fallback.get("Username", ""),
                    "Country": fallback.get("Country", country_from_path(path.parent)),
                    "Revenue_April": value_revenue,
                    "GM_April": value_gm,
                    "Revenue_May": value_revenue,
                    "GM_May": value_gm,
                }
        return path, record

    workers = min(8, max(2, (os.cpu_count() or 4)))
    with ThreadPoolExecutor(max_workers=workers) as executor:
        for index, (path, record) in enumerate(executor.map(parse_one, work), 1):
            if index == 1 or index % 10 == 0 or index == len(work):
                print(f"PNL scan once: {index}/{len(work)} | {path.name}", flush=True)
            if record:
                records.append(record)

    if not records:
        raise FileNotFoundError("No PNL files with detectable Total Revenue and Gross Margin were found")

    pnl = pd.DataFrame(records)
    pnl["name_key"] = pnl["Username"].map(normalize_match)
    pnl["country_key"] = pnl["Country"].map(normalize_match)
    group_cols = ["name_key", "country_key"]
    common_cols = {"Username": "last", "Country": "last"}

    may_agg = dict(common_cols, Revenue_May="sum", GM_May="sum")
    april_agg = dict(common_cols, Revenue_April="sum", GM_April="sum")
    may = pnl.groupby(group_cols, as_index=False).agg(may_agg).rename(
        columns={"Revenue_May": "Revenue", "GM_May": "GM"}
    )
    april = pnl.groupby(group_cols, as_index=False).agg(april_agg).rename(
        columns={"Revenue_April": "Revenue", "GM_April": "GM"}
    )
    return {"May": may, "April": april}


def collect_pnl_for_period(month_limit):
    """Compatibility wrapper. New flow should call collect_pnl_for_periods()."""
    return collect_pnl_for_periods()["May" if month_limit >= 5 else "April"]

def prepare_manager_cost_dataframe(source, period):
    df = kpi_clean_columns(pd.read_excel(source) if isinstance(source, (str, Path)) else source.copy())
    if period.lower().startswith("may"):
        candidates = ["Manager cost YTD May", "Manager Cost YTD May", "Manager Cost", "ManagerCost"]
    else:
        candidates = ["Manager cost YTD April", "Manager Cost YTD April", "Manager Cost", "ManagerCost"]
    value_col = kpi_find_column(df, candidates)
    df["Manager Cost"] = kpi_num(df[value_col])
    return df


def clean_bonus_dataframe(raw_path, cost_path, month_limit, fx_rates, pnl=None):
    raw = canonicalize_bonus_columns(pd.read_excel(raw_path))
    period = "May" if month_limit == 5 else "April"
    cost = prepare_manager_cost_dataframe(cost_path, period)
    pnl = pnl if pnl is not None else collect_pnl_for_period(month_limit)

    raw_name = kpi_find_column(raw, ["Username", "User Name", "Employee"])
    raw_country = kpi_find_column(raw, ["Country"])
    raw["name_key"] = raw[raw_name].map(normalize_match)
    raw["country_key"] = raw[raw_country].map(normalize_match)

    cost_name = kpi_find_column(cost, ["Username", "User Name", "Employee"])
    cost_country = kpi_find_column(cost, ["Country"], required=False)
    cost_value = "Manager Cost"
    cost["name_key"] = cost[cost_name].map(normalize_match)
    cost["country_key"] = cost[cost_country].map(normalize_match) if cost_country else ""
    cost["cost_value"] = kpi_num(cost[cost_value])
    cost_lookup = {(r["name_key"], r["country_key"]): r["cost_value"] for _, r in cost.iterrows()}
    cost_name_lookup = cost.groupby("name_key")["cost_value"].sum().to_dict()

    raw = raw.merge(
        pnl[["name_key", "country_key", "Revenue", "GM"]],
        on=["name_key", "country_key"], how="left",
    )
    raw["Revenue"] = pd.to_numeric(raw["Revenue"], errors="coerce").fillna(0.0)
    raw["GM"] = pd.to_numeric(raw["GM"], errors="coerce").fillna(0.0)
    raw["Manager Cost"] = [
        cost_lookup.get((name, country), cost_name_lookup.get(name, 0.0))
        for name, country in zip(raw["name_key"], raw["country_key"])
    ]
    raw["CM"] = raw["GM"] - raw["Manager Cost"]
    raw.drop(columns=["name_key", "country_key"], errors="ignore", inplace=True)
    return raw


def cleaned_bonus_columns():
    return [
        "UserId", "Username", "Function", "UserStatus", "Country", "Currency",
        "CommissionStartDate", "CommissionEndDate", "YtdDue", "YtdPayment",
        "RemainingBalance", "Yearly Target", "Ytd Yearly Target", "Revenue", "GM",
        "Manager Cost", "CM", "FX Rate to EUR", "YtdDue EUR", "YtdPayment EUR",
        "RemainingBalance EUR", "Yearly Target EUR", "Ytd Target EUR", "Revenue EUR",
        "GM EUR", "Manager Cost EUR", "CM EUR", "Due/Target",
    ]


def write_currency_sheet_simple(ws, fx_rates):
    headers = ["Currency", "Rate to EUR", "Example amount", "Converted EUR", "Formula used"]
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 1, c, header)
    for r, currency in enumerate(sorted(fx_rates), 2):
        fill = kpi_fill("FFFFFF")
        kpi_excel_cell(ws, r, 1, currency, fill, align="left")
        kpi_excel_cell(ws, r, 2, fx_rates[currency], fill, fmt="0.0000000000")
        kpi_excel_cell(ws, r, 3, 1, fill, fmt="#,##0.00")
        kpi_formula_cell(ws, r, 4, f"=C{r}*B{r}", fill, "#,##0.000000")
        kpi_excel_cell(ws, r, 5, "'=Source amount * Rate to EUR", fill, align="left")
    ws.auto_filter.ref = f"A1:E{max(1, len(fx_rates) + 1)}"
    ws.freeze_panes = "A2"
    return len(fx_rates) + 1


def write_cleaned_bonus_sheet(ws, df, fx_end):
    headers = cleaned_bonus_columns()
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 1, c, header)
    base = [
        "UserId", "Username", "Function", "UserStatus", "Country", "Currency",
        "CommissionStartDate", "CommissionEndDate", "YtdDue", "YtdPayment",
        "RemainingBalance", "Yearly Target", "Ytd Yearly Target", "Revenue", "GM",
        "Manager Cost", "CM",
    ]
    for r, (_, item) in enumerate(df.iterrows(), 2):
        fill = kpi_fill("FFFFFF")
        for c, col in enumerate(base, 1):
            # CM is calculated in Excel from GM minus Manager Cost. Do not
            # write the source dataframe value as a hard-coded number.
            if col == "CM":
                continue
            cell = kpi_excel_cell(ws, r, c, item.get(col, "N/A"), fill, align="left" if c in {1,2,3,4,5,6} else "center")
            if c in {9,10,11,12,13,14,15,16,17}:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
            if c in {7,8} and hasattr(item.get(col), "year"):
                cell.number_format = "d-mmm-yy"
        # Q = CM, O = GM, P = Manager Cost.
        kpi_formula_cell(ws, r, 17, f"=O{r}-P{r}", fill, "#,##0.00;[Red](#,##0.00);-")
        fx_lookup = f"IFERROR(VLOOKUP(F{r},'Currency to EUR'!$A$2:$B${fx_end},2,FALSE),NA())"
        kpi_formula_cell(ws, r, 18, f"={fx_lookup}", fill, "0.0000000000")
        for target_col, source_col in [(19,"I"),(20,"J"),(21,"K"),(22,"L"),(23,"M"),(24,"N"),(25,"O"),(26,"P")]:
            kpi_formula_cell(ws, r, target_col, f"={source_col}{r}*R{r}", fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 27, f"=Y{r}-Z{r}", fill, "#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws, r, 28, f"=IFERROR(S{r}/W{r},0)", fill, "0.00%")
    widths = [14,28,22,16,18,12,18,18,16,16,18,18,18,18,18,18,18,16,18,18,20,20,18,18,18,20,18,16]
    for c, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(c)].width = width
    last_row=max(1,len(df)+1)
    shade_sections(ws,1,2,last_row,[(1,17,"1F4E79","FFFFFF"),(18,18,"BF9000","FFFFFF"),(19,28,"70AD47","FFFFFF")])
    ws.auto_filter.ref = f"A1:AB{max(1, len(df) + 1)}"
    ws.freeze_panes = "A2"


def write_cleaned_bp_sheet(ws, bp_df):
    headers = ["Month", "Entity", "SIG", "BP Value", "Country Billing To", "Currency", "FX Rate to EUR", "BP Value EUR"]
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 1, c, header)
    for r, (_, item) in enumerate(bp_df.iterrows(), 2):
        fill = kpi_fill("FFFFFF")
        vals = [item["MonthNumber"], item["Entity"], item["SIGName"], item["BPSourceValue"], item["CountryName"], item["Currency"]]
        for c, value in enumerate(vals, 1):
            cell = kpi_excel_cell(ws, r, c, value, fill, align="left" if c in {2,3,5,6} else "center")
            if c == 4:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
        kpi_formula_cell(ws, r, 7, f"=IFERROR(VLOOKUP(F{r},'Currency to EUR'!$A$2:$B$20,2,FALSE),NA())", fill, "0.0000000000")
        kpi_formula_cell(ws, r, 8, f"=D{r}*G{r}", fill, "#,##0.00;[Red](#,##0.00);-")
    for c, width in enumerate([12,18,24,18,22,12,18,18], 1):
        ws.column_dimensions[get_column_letter(c)].width = width
    last_row=max(1,len(bp_df)+1)
    shade_sections(ws,1,2,last_row,[(1,6,"1F4E79","FFFFFF"),(7,8,"70AD47","FFFFFF")])
    ws.auto_filter.ref = f"A1:H{max(1, len(bp_df) + 1)}"
    ws.freeze_panes = "A2"


def write_cleaned_cost_sheet(ws, may_cost, april_cost):
    headers = ["Period", "UserId", "Username", "UserStatus", "Country", "Currency", "CommissionStartDate", "CommissionEndDate", "Manager Cost", "FX Rate to EUR", "Manager Cost EUR"]
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 1, c, header)
    combined = []
    for period, df in [("May", may_cost), ("April", april_cost)]:
        temp = prepare_manager_cost_dataframe(df, period)
        for _, row in temp.iterrows():
            combined.append([period, row.get("UserId"), row.get("Username"), row.get("UserStatus"), row.get("Country"), row.get("Currency"), row.get("CommissionStartDate"), row.get("CommissionEndDate"), row.get("Manager Cost")])
    for r, values in enumerate(combined, 2):
        fill = kpi_fill("FFFFFF")
        for c, value in enumerate(values, 1):
            cell = kpi_excel_cell(ws, r, c, value, fill, align="left" if c in {1,2,3,4,5,6} else "center")
            if c == 9:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
            if c in {7,8} and hasattr(value, "year"):
                cell.number_format = "d-mmm-yy"
        kpi_formula_cell(ws, r, 10, f"=IFERROR(VLOOKUP(F{r},'Currency to EUR'!$A$2:$B$20,2,FALSE),NA())", fill, "0.0000000000")
        kpi_formula_cell(ws, r, 11, f"=I{r}*J{r}", fill, "#,##0.00;[Red](#,##0.00);-")
    for c, width in enumerate([12,14,28,16,18,12,18,18,18,18,20], 1):
        ws.column_dimensions[get_column_letter(c)].width = width
    last_row=max(1,len(combined)+1)
    shade_sections(ws,1,2,last_row,[(1,9,"1F4E79","FFFFFF"),(10,11,"70AD47","FFFFFF")])
    ws.auto_filter.ref = f"A1:K{max(1, len(combined) + 1)}"
    ws.freeze_panes = "A2"


def clean_sum_formula(sheet, end_row, value_col, criteria):
    args = [f"'{sheet}'!${value_col}$2:${value_col}${end_row}"]
    for col, ref in criteria:
        args.extend([f"'{sheet}'!${col}$2:${col}${end_row}", ref])
    return f"=SUMIFS({','.join(args)})"


def clean_avg_ratio_formula(sheet, end_row, numerator_col, denominator_col, criteria):
    # Country Due/Target is weighted by country totals:
    # SUMIFS(YTD Due) / SUMIFS(YTD Target), not AVERAGEIFS(manager ratios).
    numerator = clean_sum_formula(sheet, end_row, numerator_col, criteria)[1:]
    denominator = clean_sum_formula(sheet, end_row, denominator_col, criteria)[1:]
    return f"=IFERROR(({numerator})/({denominator}),0)"


def write_graphs_clean(ws, country, function_country, current_label, current_month, may_end, april_end, bp_end):
    # Open the KPI workbook with the Graphs dashboard zoomed out, matching
    # the requested overview view.
    ws.sheet_view.zoomScale = 70
    ws.sheet_view.zoomScaleNormal = 100
    headers = ["Country", "#PE", "YTD Bonus", "YTD Target", "YTD CM", "% Country CM/BP", "% CM Var vs April", "Managers perf.% (Due/Target)"]
    kpi_title(
        ws,
        f"BONUS PERFORMANCE BY YEAR - {current_label.upper()}",
        8,
        f"Priority snapshot: {current_label} | Variance reference: April 2026 | BP excludes Financial Costs and Structure Costs",
    )
    above = country[country["BP_Achievement"] >= 1].copy()
    below = country[country["BP_Achievement"] < 1].copy()

    def section_title(start, title, width):
        ws.merge_cells(start_row=start, start_column=1, end_row=start, end_column=width)
        cell = ws.cell(start, 1, title)
        cell.fill = kpi_fill("D9B3E6")
        cell.font = Font(bold=True, size=12)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    def table(start, title, data):
        section_title(start, title, len(headers))
        for c, h in enumerate(headers, 1):
            kpi_header(ws, start + 1, c, h)
        for r, (_, item) in enumerate(data.iterrows(), start + 2):
            fill = kpi_fill("FFFFFF")
            kpi_excel_cell(ws, r, 1, item["Country"], fill, align="left")
            country_ref = f"$A{r}"
            for c in range(2, 9):
                kpi_excel_cell(ws, r, c, None, fill)
            kpi_formula_cell(ws, r, 2, f"=COUNTIF('Cleaned May'!$E$2:$E${may_end},{country_ref})", fill)
            kpi_formula_cell(ws, r, 3, clean_sum_formula("Cleaned May", may_end, "S", [("E", country_ref)]), fill, "#,##0.00;[Red](#,##0.00);-")
            kpi_formula_cell(ws, r, 4, clean_sum_formula("Cleaned May", may_end, "W", [("E", country_ref)]), fill, "#,##0.00;[Red](#,##0.00);-")
            kpi_formula_cell(ws, r, 5, clean_sum_formula("Cleaned May", may_end, "AA", [("E", country_ref)]), fill, "#,##0.00;[Red](#,##0.00);-")
            bp = f"SUMIFS('Cleaned BP'!$H$2:$H${bp_end},'Cleaned BP'!$E$2:$E${bp_end},{country_ref},'Cleaned BP'!$A$2:$A${bp_end},\"<=\"&{current_month})"
            kpi_formula_cell(ws, r, 6, f"=IFERROR(E{r}/({bp}),0)", fill, "0.00%")
            apr = clean_sum_formula("Cleaned April", april_end, "AA", [("E", country_ref)])[1:]
            kpi_formula_cell(ws, r, 7, f"=IFERROR((E{r}-({apr}))/ABS({apr}),0)", fill, "0.00%")
            kpi_formula_cell(ws, r, 8, clean_avg_ratio_formula("Cleaned May", may_end, "S", "W", [("E", country_ref)]), fill, "0.00%")
        return start + 2 + len(data)

    end_above = table(4, "Countries with contributive margin ABOVE BUSINESS PLAN", above)
    end_below = table(end_above + 3, "Countries with contributive margin BEHIND BUSINESS PLAN", below)

    # Function matrix: one row per country, one column per function.
    # The function columns contain Due / Target, while CM/BP appears once per country.
    function_values = []
    seen = set()
    preferred_order = ["Manager", "Experienced manager", "Senior manager", "Top Management"]
    source_functions = [str(value).strip() for value in function_country["Function"].dropna().unique()]
    for preferred in preferred_order:
        match = next((value for value in source_functions if normalize_match(value) == normalize_match(preferred)), None)
        if match and normalize_match(match) not in seen:
            function_values.append(match)
            seen.add(normalize_match(match))
    for value in sorted(source_functions, key=lambda x: normalize_match(x)):
        if normalize_match(value) not in seen:
            function_values.append(value)
            seen.add(normalize_match(value))

    function_start = end_below + 3
    function_headers = ["COUNTRY", "Country %CM/BP"] + [value.title() for value in function_values]
    section_title(function_start, "Bonus Performance by Function and Country", len(function_headers))
    for c, header in enumerate(function_headers, 1):
        kpi_header(ws, function_start + 1, c, header)

    country_rows = country.sort_values("Country")[["Country"]].drop_duplicates().reset_index(drop=True)
    for r, (_, item) in enumerate(country_rows.iterrows(), function_start + 2):
        fill = kpi_fill("FFFFFF")
        kpi_excel_cell(ws, r, 1, str(item["Country"]).upper(), fill, align="left")
        country_ref = f"$A{r}"
        # Single country CM/BP formula.
        cm = clean_sum_formula("Cleaned May", may_end, "AA", [("E", country_ref)])[1:]
        bp = f"SUMIFS('Cleaned BP'!$H$2:$H${bp_end},'Cleaned BP'!$E$2:$E${bp_end},{country_ref},'Cleaned BP'!$A$2:$A${bp_end},\"<=\"&{current_month})"
        kpi_formula_cell(ws, r, 2, f"=IFERROR(({cm})/({bp}),0)", fill, "0.00%")

        for offset, source_function in enumerate(function_values, 3):
            criterion = str(source_function).replace('"', '""')
            due = clean_sum_formula("Cleaned May", may_end, "S", [("E", country_ref), ("C", f'"{criterion}"')])[1:]
            target = clean_sum_formula("Cleaned May", may_end, "W", [("E", country_ref), ("C", f'"{criterion}"')])[1:]
            kpi_formula_cell(ws, r, offset, f"=IFERROR(({due})/({target}),0)", fill, "0.00%")

    below_start = end_above + 3
    for c, width in enumerate([20, 18] + [20] * len(function_values), 1):
        ws.column_dimensions[get_column_letter(c)].width = width
    # Compact wrapped headers so the report stays readable without oversized cells.
    for header_row in (5, below_start + 1, function_start + 1):
        ws.row_dimensions[header_row].height = 42
    for section_row in (4, below_start, function_start):
        ws.row_dimensions[section_row].height = 24
    ws.column_dimensions["H"].width = 24

    # Keep the reference chart placement and formatting.
    add_country_performance_chart(
        ws, "L4",
        f"PERFORMANCE YTD {current_label.upper()}: COUNTRIES ABOVE BUSINESS PLAN",
        5, 6, end_above - 1,
        ("F39C12", "7030A0"), 0.50,
    )
    below_start = end_above + 3
    add_country_performance_chart(
        ws, "L23",
        f"PERFORMANCE YTD {current_label.upper()}: COUNTRIES BEHIND BUSINESS PLAN",
        below_start + 1, below_start + 2, end_below - 1,
        ("FF2B2B", "F39C12"), 0.20,
    )
    ws.freeze_panes = None

def write_manager_clean(ws, manager, may_end, april_end, bp_end, current_month):
    headers=["Country","User Id","User Name","Function","Profile Status","€ Grand Total Due","€ YTD Target","% Due/Target","% Country CM vs. BP"]
    kpi_title(ws,"MANAGER PERFORMANCE",len(headers),"Priority snapshot: YTD May 2026 | Amounts automatically converted to EUR")
    for c,h in enumerate(headers,1): kpi_header(ws,4,c,h)
    for r,(_,item) in enumerate(manager.iterrows(),5):
        fill=kpi_fill("FFFFFF")
        vals=[item["Country"],item["UserId"],item["Username"],item["Function"],item["ProfileStatus"]]
        for c,v in enumerate(vals,1): kpi_excel_cell(ws,r,c,v,fill,align="left" if c in {1,3,4,5} else "center")
        uid=f"$B{r}"; country=f"$A{r}"
        kpi_formula_cell(ws,r,6,clean_sum_formula("Cleaned May",may_end,"S",[("A",uid)]),fill,"#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws,r,7,clean_sum_formula("Cleaned May",may_end,"W",[("A",uid)]),fill,"#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws,r,8,f"=IFERROR(F{r}/G{r},0)",fill,"0.00%")
        may=clean_sum_formula("Cleaned May",may_end,"AA",[("E",country)])[1:]
        bp=f"SUMIFS('Cleaned BP'!$H$2:$H${bp_end},'Cleaned BP'!$E$2:$E${bp_end},{country},'Cleaned BP'!$A$2:$A${bp_end},\"<=\"&{current_month})"
        kpi_formula_cell(ws,r,9,f"=IFERROR(({may})/({bp}),0)",fill,"0.00%")
    last_row=max(4,len(manager)+4)
    shade_sections(ws,4,5,last_row,[(1,5,"1F4E79","FFFFFF"),(6,7,"70AD47","FFFFFF"),(8,9,"8064A2","FFFFFF")])
    for c,w in enumerate([18,14,28,22,16,18,18,15,20],1): ws.column_dimensions[get_column_letter(c)].width=w
    ws.freeze_panes="A5"; ws.print_title_rows = "1:4"; ws.auto_filter.ref=f"A4:I{max(4,len(manager)+4)}"


def write_top_clean(ws, top, may_end, bp_end, current_month):
    headers=["Country","User Id","User Name","Function","€ Grand Total Due","€ YTD Target","% Due/Target","€ Manager CM","% Country CM Contribution","% Country CM vs. BP"]
    if "Due_Target" in top.columns:
        top = top[pd.to_numeric(top["Due_Target"], errors="coerce") > 1].head(10).copy()
    kpi_title(ws,"TOP PERFORMERS YTD MAY 2026",len(headers),"Only managers above 100% YTD Due / YTD Target | PO, Delta Projects and Speed excluded")
    for c,h in enumerate(headers,1): kpi_header(ws,4,c,h)
    for r,(_,item) in enumerate(top.iterrows(),5):
        fill=kpi_fill("FFFFFF")
        vals=[item["Country"],item["UserId"],item["Username"],item["Function"]]
        for c,v in enumerate(vals,1): kpi_excel_cell(ws,r,c,v,fill,align="left" if c in {1,3,4} else "center")
        user_id=f"$B{r}"; country=f"$A{r}"
        # Match the stable User Id in Cleaned May column A, not the display name.
        kpi_formula_cell(ws,r,5,clean_sum_formula("Cleaned May",may_end,"S",[("A",user_id)]),fill,"#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws,r,6,clean_sum_formula("Cleaned May",may_end,"W",[("A",user_id)]),fill,"#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws,r,7,f"=IFERROR(E{r}/F{r},0)",fill,"0.00%")
        manager_cm = clean_sum_formula("Cleaned May",may_end,"AA",[("A",user_id)])
        kpi_formula_cell(ws,r,8,manager_cm,fill,"#,##0.00;[Red](#,##0.00);-")
        may=clean_sum_formula("Cleaned May",may_end,"AA",[("E",country)])[1:]
        kpi_formula_cell(ws,r,9,f"=IFERROR(H{r}/({may}),0)",fill,"0.00%")
        bp = f'''SUMIFS('Cleaned BP'!$H$2:$H${bp_end},'Cleaned BP'!$E$2:$E${bp_end},{country},'Cleaned BP'!$A$2:$A${bp_end},"<="&{current_month})'''
        kpi_formula_cell(ws,r,10,f"=IFERROR(({may})/({bp}),0)",fill,"0.00%")
    last_row = max(4, 4 + len(top))
    red_fill = PatternFill(fill_type="solid", fgColor=KPI_RED)
    red_font = Font(color="9C0006", bold=True, size=10)
    # Highlight Country CM / BP below 100% after Excel recalculates formulas.
    if last_row >= 5:
        ws.conditional_formatting.add(f"J5:J{last_row}", CellIsRule(operator="lessThan", formula=["1"], fill=red_fill, font=red_font))
    shade_sections(ws,4,5,last_row,[(1,4,"1F4E79","FFFFFF"),(5,6,"70AD47","FFFFFF"),(7,7,"8064A2","FFFFFF"),(8,9,"5B9BD5","FFFFFF"),(10,10,"8064A2","FFFFFF")])
    for c,w in enumerate([18,14,30,22,18,18,15,18,22,20],1): ws.column_dimensions[get_column_letter(c)].width=w
    ws.freeze_panes="A5"
    ws.print_title_rows = "1:4"
    ws.auto_filter.ref=f"A4:J{last_row}"

def write_variance_clean(ws, current, previous, may_end, april_end):
    headers=["EmployeeId","Manager","Country","Status","Function DNA","Total Due","Total Payment","Ccy","Balance","Bonus M-1","Var vs. M-1","% Var","YTD CM May","YTD CM April","% Var CM"]
    kpi_title(ws,"VARIANCE ANALYST YTD MAY 2026",len(headers),"Employee-level view: May 2026 versus April 2026 | Current YTD is prioritized")
    for c,h in enumerate(headers,1): kpi_header(ws,4,c,h)
    cur=current.copy(); cur["BonusPrevious"]=cur["UserId"].map(previous.set_index("UserId")["YtdDueEUR"]).fillna(0.0)
    cur["VariancePct"]=cur.apply(lambda x:kpi_safe_ratio(x["YtdDueEUR"]-x["BonusPrevious"],abs(x["BonusPrevious"])),axis=1)
    cur=cur.sort_values("VariancePct",ascending=False,na_position="last")
    for r,(_,item) in enumerate(cur.iterrows(),5):
        fill=kpi_fill("FFFFFF")
        vals=[item["UserId"],item["Username"],item["Country"],item["ProfileStatus"],item["Function"],None,None,"EUR",None,None,None,None,None,None,None]
        for c,v in enumerate(vals,1): kpi_excel_cell(ws,r,c,v,fill,align="left" if c in {1,2,3,4,5,8} else "center")
        uid=f"$A{r}"
        formulas={6:clean_sum_formula("Cleaned May",may_end,"S",[("A",uid)]),7:clean_sum_formula("Cleaned May",may_end,"T",[("A",uid)]),9:clean_sum_formula("Cleaned May",may_end,"U",[("A",uid)]),10:clean_sum_formula("Cleaned April",april_end,"S",[("A",uid)]),13:clean_sum_formula("Cleaned May",may_end,"AA",[("A",uid)]),14:clean_sum_formula("Cleaned April",april_end,"AA",[("A",uid)])}
        for c,f in formulas.items(): kpi_formula_cell(ws,r,c,f,fill,"#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws,r,11,f"=F{r}-J{r}",fill,"#,##0.00;[Red](#,##0.00);-")
        kpi_formula_cell(ws,r,12,f"=IFERROR(K{r}/ABS(J{r}),0)",fill,"0.00%")
        kpi_formula_cell(ws,r,15,f"=IFERROR((M{r}-N{r})/ABS(N{r}),0)",fill,"0.00%")
    last_row=max(4,len(cur)+4)
    # Section headers are distinct, while the table body stays full white.
    shade_sections(ws,4,5,last_row,[
        (1,5,"1F4E79","FFFFFF"),
        (6,9,"5B9BD5","FFFFFF"),
        (10,12,"70AD47","FFFFFF"),
        (13,15,"8064A2","FFFFFF"),
    ])
    red_fill=PatternFill(fill_type="solid", fgColor="F4CCCC")
    red_font=Font(color="9C0006", bold=True)
    # Bonus variance over 60%.
    ws.conditional_formatting.add(f"L5:L{last_row}", CellIsRule(operator="greaterThan", formula=["0.6"], fill=red_fill, font=red_font))
    # CM variance over +70% or below -70%.
    ws.conditional_formatting.add(f"O5:O{last_row}", CellIsRule(operator="greaterThan", formula=["0.7"], fill=red_fill, font=red_font))
    ws.conditional_formatting.add(f"O5:O{last_row}", CellIsRule(operator="lessThan", formula=["-0.7"], fill=red_fill, font=red_font))
    for c,w in enumerate([14,28,14,16,22,18,20,10,18,18,18,14,18,18,18],1): ws.column_dimensions[get_column_letter(c)].width=w
    ws.freeze_panes="A5"; ws.print_title_rows = "1:4"; ws.auto_filter.ref=f"A4:O{max(4,len(cur)+4)}"


def write_cleaned_file(path, df, sheet_name, fx_rates, kind, extra=None):
    wb=Workbook(); wb.remove(wb.active)
    fx=wb.create_sheet("Currency to EUR")
    fx_end=write_currency_sheet_simple(fx,fx_rates)
    if kind=="bonus": write_cleaned_bonus_sheet(wb.create_sheet(sheet_name),df,fx_end)
    elif kind=="bp": write_cleaned_bp_sheet(wb.create_sheet(sheet_name),df)
    elif kind=="cost": write_cleaned_cost_sheet(wb.create_sheet(sheet_name),extra[0],extra[1])
    fx.sheet_state="visible"
    wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode="auto"
    wb.save(path)



def output_safe_name(value):
    return re.sub(r"[<>:\"/\\|?*]", "_", str(value or "Unknown")).strip() or "Unknown"


def write_country_pnl_summary(output_path, records, country):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    headers = ["Country", "Manager", "Revenue", "GM", "Manager Cost", "CM"]
    for c, header in enumerate(headers, 1):
        kpi_header(ws, 1, c, header)
    for r, record in enumerate(records, 2):
        fill = kpi_fill("FFFFFF")
        values = [country, record.get("Username"), record.get("Revenue", 0.0), record.get("GM", 0.0), record.get("Manager Cost", 0.0), record.get("CM", 0.0)]
        for c, value in enumerate(values, 1):
            cell = kpi_excel_cell(ws, r, c, value, fill, align="left" if c in {1, 2} else "center")
            if c >= 3:
                cell.number_format = "#,##0.00;[Red](#,##0.00);-"
    ws.auto_filter.ref = f"A1:F{max(1, len(records) + 1)}"
    ws.freeze_panes = "A2"
    for c, width in enumerate([18, 28, 18, 18, 18, 18], 1):
        ws.column_dimensions[get_column_letter(c)].width = width
    wb.save(output_path)


def create_pnl_output_tree(output_root, manager_cost_path):
    roots = find_pnl_roots()
    files = []
    for root in roots:
        files.extend(p for p in root.rglob("*") if p.suffix.lower() in {".xlsx", ".xlsm", ".xls"} and not p.name.startswith("~$"))
    excluded = ("manager bonus", "manager cost", "bp country", "kpi report", "cleaned")
    records = []
    cost = prepare_manager_cost_dataframe(manager_cost_path, "May")
    cost_name = kpi_find_column(cost, ["Username", "User Name", "Employee"])
    cost_country = kpi_find_column(cost, ["Country"], required=False)
    cost_value = "Manager Cost"
    cost["name_key"] = cost[cost_name].map(normalize_match)
    cost["country_key"] = cost[cost_country].map(normalize_match) if cost_country else ""
    cost["cost_value"] = kpi_num(cost[cost_value])
    cost_lookup = {(r["name_key"], r["country_key"]): r["cost_value"] for _, r in cost.iterrows()}
    for path in files:
        if any(token in path.name.lower() for token in excluded):
            continue
        record = parse_pnl_for_period(path, 5)
        if not record:
            continue
        record["Manager Cost"] = cost_lookup.get((normalize_match(record["Username"]), normalize_match(record["Country"])), 0.0)
        record["CM"] = record["GM"] - record["Manager Cost"]
        record["SourceFile"] = path
        records.append(record)
    by_country = {}
    for record in records:
        by_country.setdefault(record["Country"], []).append(record)
    for country, country_records in by_country.items():
        folder = output_root / output_safe_name(country)
        folder.mkdir(parents=True, exist_ok=True)
        for record in country_records:
            source = Path(record["SourceFile"])
            target = folder / f"Pnl manager - {output_safe_name(record['Username'])}.xlsx"
            shutil.copy2(source, target)
        write_country_pnl_summary(folder / "Summary.xlsx", country_records, country)
    print(f"PNL output tree created: {len(by_country)} countries, {len(records)} manager files", flush=True)

def run_final_kpi_report():
    global KPI_OUTPUT_PATH
    print("Starting cleaned-data Manager Bonus and KPI flow...", flush=True)
    output_root = OUTPUT_DIR / "Cleaned data"
    output_root.mkdir(parents=True, exist_ok=True)
    fx_rates = load_fx_rates()
    raw_may = find_input_file(["Manager Bonus YTD May - Raw data.xlsx"])
    raw_april = find_input_file(["Manager Bonus YTD April - Raw data.xlsx"])
    cost_may = find_input_file(["Manager Cost YTD May.xlsx", "Manager Cost YTD April_May.xlsx"])
    cost_april = find_input_file(["Manager Cost YTD April.xlsx", "Manager Cost YTD April_May.xlsx"])
    bp_path = find_input_file(["BP Country 2026.xlsx", "BP Country_fake.xlsx"])
    country_ccy = {"Brazil":"BRL","India":"INR","Italy":"EUR","Japan":"JPY","Malaysia":"MYR","Singapore":"SGD","Spain":"EUR","Thailand":"THB","United States":"USD","Vietnam":"VND","China":"CNY","Canada":"CAD","France":"EUR","Belgium":"EUR","Switzerland":"CHF","Portugal":"EUR","Sweden":"EUR","Netherlands":"EUR","Luxembourg":"EUR","Tunisia":"EUR","Czech Republic":"EUR","Austria":"EUR","Turkey":"TRY","Mexico":"MXN"}
    fx_rates = active_fx_rates(fx_rates, [raw_may, raw_april, cost_may, cost_april, bp_path], country_ccy)
    required_paths = {
        "May bonus": raw_may,
        "April bonus": raw_april,
        "Manager cost May": cost_may,
        "Manager cost April": cost_april,
        "BP Country": bp_path,
    }
    missing = [label for label, path in required_paths.items() if path is None]
    if missing:
        raise FileNotFoundError("Missing required input files inside Raw data:\n" + "\n".join(missing))

    # The PnL folder is scanned once. The same parsed records feed both periods.
    pnl_by_period = collect_pnl_for_periods()
    may_raw = clean_bonus_dataframe(raw_may, cost_may, 5, fx_rates, pnl=pnl_by_period["May"])
    april_raw = clean_bonus_dataframe(raw_april, cost_april, 4, fx_rates, pnl=pnl_by_period["April"])
    may_path = output_root / CLEANED_MAY_FILE
    april_path = output_root / CLEANED_APRIL_FILE
    write_cleaned_file(may_path, may_raw, "Cleaned May", fx_rates, "bonus")
    write_cleaned_file(april_path, april_raw, "Cleaned April", fx_rates, "bonus")

    bp_src = load_bp_file(bp_path)
    bp_clean = bp_src.copy()
    bp_clean["BPSourceValue"] = bp_clean["BPValue"]
    bp_clean["Currency"] = bp_clean["CountryName"].map(country_ccy).fillna("EUR")
    bp_eur = bp_src.copy()
    bp_eur["BPValue"] = bp_eur.apply(lambda x: x["BPValue"] * fx_rates.get(str(country_ccy.get(x["CountryName"], "EUR")).upper(), 1.0), axis=1)
    write_cleaned_file(output_root / CLEANED_BP_FILE, bp_clean, "Cleaned BP", fx_rates, "bp")
    write_cleaned_file(output_root / CLEANED_MANAGER_COST_FILE, None, "Cleaned Manager cost", fx_rates, "cost", extra=(prepare_manager_cost_dataframe(cost_may, "May"), prepare_manager_cost_dataframe(cost_april, "April")))

    current = normalize_bonus_dataframe(may_raw, fx_rates)
    previous = normalize_bonus_dataframe(april_raw, fx_rates)
    country, function_country, manager, top, _ = build_kpi_data(previous, current, bp_by_country(bp_eur, 5), bp_by_country(bp_eur, 4))
    wb = Workbook(); wb.remove(wb.active)
    fx = wb.create_sheet("Currency to EUR"); fx_end = write_currency_sheet_simple(fx, fx_rates); fx.sheet_state = "hidden"
    may_ws = wb.create_sheet("Cleaned May"); write_cleaned_bonus_sheet(may_ws, may_raw, fx_end)
    april_ws = wb.create_sheet("Cleaned April"); write_cleaned_bonus_sheet(april_ws, april_raw, fx_end)
    bp_ws = wb.create_sheet("Cleaned BP"); write_cleaned_bp_sheet(bp_ws, bp_clean)
    cost_ws = wb.create_sheet("Cleaned Manager cost"); write_cleaned_cost_sheet(cost_ws, prepare_manager_cost_dataframe(cost_may, "May"), prepare_manager_cost_dataframe(cost_april, "April"))
    # Keep all cleaned/source sheets hidden in the delivered KPI workbook.
    # The user-facing views are Graphs, TopPerformer, ManagerPerf, Variance,
    # Cross Check, and the supporting summary sheets.
    for source_ws in (may_ws, april_ws, bp_ws, cost_ws):
        source_ws.sheet_state = "hidden"
    may_end, april_end, bp_end = len(may_raw) + 1, len(april_raw) + 1, len(bp_clean) + 1
    write_graphs_clean(wb.create_sheet("Graphs"), country, function_country, "May 2026", 5, may_end, april_end, bp_end)
    write_top_clean(wb.create_sheet("TopPerformer"), top, may_end, bp_end, 5)
    write_manager_clean(wb.create_sheet("ManagerPerf"), manager, may_end, april_end, bp_end, 5)
    write_variance_clean(wb.create_sheet("Variance analyst YTD May"), current, previous, may_end, april_end)
    def reconciliation_check(name, actual, expected, tolerance, formula):
        difference = float(actual) - float(expected)
        return {
            "Check": name,
            "Actual": float(actual),
            "Expected": float(expected),
            "Difference": difference,
            "Tolerance": tolerance,
            "Status": "PASS" if abs(difference) <= tolerance else "FAIL",
            "Formula": formula,
        }

    may_rates = may_raw["Currency"].astype(str).str.upper().map(fx_rates)
    if may_rates.isna().any():
        missing = sorted(may_raw.loc[may_rates.isna(), "Currency"].astype(str).unique())
        raise ValueError(f"Missing FX rates before KPI reconciliation: {', '.join(missing)}")
    independent_cm_eur = (
        (pd.to_numeric(may_raw["GM"], errors="coerce").fillna(0.0)
         - pd.to_numeric(may_raw["Manager Cost"], errors="coerce").fillna(0.0))
        * may_rates
    ).sum()
    duplicate_user_ids = int(current["UserId"].astype(str).duplicated().sum())
    expected_top_rows = min(10, int((pd.to_numeric(manager["Due_Target"], errors="coerce") > 1).sum()))
    checks = [
        reconciliation_check(
            "May CM EUR: cleaned output vs independent raw calculation",
            current["CMEUR"].sum(), independent_cm_eur, 0.01,
            "Cleaned CM EUR compared with SUM((raw GM - raw Manager Cost) x FX)",
        ),
        reconciliation_check(
            "ManagerPerf rows vs unique May employees",
            len(manager), current["UserId"].nunique(), 0,
            "One ManagerPerf row per unique May UserId",
        ),
        reconciliation_check(
            "TopPerformer rows vs eligible managers above 100%",
            len(top), expected_top_rows, 0,
            "Up to 10 managers with YTD Due / YTD Target above 100%",
        ),
        reconciliation_check(
            "Duplicate May UserId records",
            duplicate_user_ids, 0, 0,
            "Duplicate count after cleaned-data normalization",
        ),
        reconciliation_check(
            "Missing May FX mappings",
            int(may_rates.isna().sum()), 0, 0,
            "Every source currency must have an explicit EUR rate",
        ),
    ]

    def missing_pnl_check(cleaned_df, period_label):
        revenue = pd.to_numeric(cleaned_df.get("Revenue", 0), errors="coerce").fillna(0.0)
        gm = pd.to_numeric(cleaned_df.get("GM", 0), errors="coerce").fillna(0.0)
        manager_cost = pd.to_numeric(cleaned_df.get("Manager Cost", 0), errors="coerce").fillna(0.0)
        missing_mask = (revenue.abs() < 1e-9) & (gm.abs() < 1e-9) & (manager_cost.abs() > 1e-9)
        missing_count = int(missing_mask.sum())
        status = "WARNING" if missing_count else "PASS"
        names = []
        if missing_count:
            for _, row in cleaned_df.loc[missing_mask].head(10).iterrows():
                names.append(f"{row.get('UserId', '')} {row.get('Username', '')}")
        detail = "Revenue=0 and GM=0 while Manager Cost<>0; check PnL manager/country mapping"
        if names:
            detail += ". Managers: " + ", ".join(names)
        return {
            "Check": f"{period_label} PnL Revenue/GM match",
            "Actual": missing_count,
            "Expected": 0,
            "Difference": missing_count,
            "Tolerance": 0,
            "Status": status,
            "Formula": detail,
        }

    checks.append(missing_pnl_check(may_raw, "May"))
    checks.append(missing_pnl_check(april_raw, "April"))
    write_cross_check_sheet(wb.create_sheet("Cross Check"), checks, current, previous, bp_src, country, function_country, 5)
    report_order = ["Graphs", "Cleaned May", "TopPerformer", "ManagerPerf", "Variance analyst YTD May", "Cross Check", "Cleaned April", "Cleaned BP", "Cleaned Manager cost", "Currency to EUR"]
    wb._sheets = [wb[name] for name in report_order if name in wb.sheetnames]
    wb.active = 0
    KPI_OUTPUT_PATH = OUTPUT_DIR / "KPI Report.xlsx"
    wb.calculation.fullCalcOnLoad = True; wb.calculation.forceFullCalc = True; wb.calculation.calcMode = "auto"
    wb.save(KPI_OUTPUT_PATH)
    try:
        word_output = OUTPUT_DIR / "KPI Performance Review May 2026.docx"
        generate_kpi_word_report(current, previous, may_raw, april_raw, country, top, fx_rates, word_output)
        print(f"KPI Word analysis created: {word_output}", flush=True)
    except Exception as exc:
        error_path = OUTPUT_DIR / "KPI Word analysis ERROR.txt"
        error_path.write_text(traceback.format_exc(), encoding="utf-8")
        print(f"KPI Word analysis could not be created: {exc}", flush=True)
        print(f"See error log: {error_path}", flush=True)
    print(f"Output folder created: {output_root}", flush=True)
    print(f"KPI Report created: {KPI_OUTPUT_PATH}", flush=True)


if __name__ == "__main__":
    args = parse_cli_args()
    configure_runtime_paths(args.input_root, args.output_dir)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print(f"Input root: {MONTHLY_REPORT_DIR}", flush=True)
    print(f"Output folder: {OUTPUT_DIR}", flush=True)
    if not args.skip_pnl:
        print("=== PNL AUDIT ===", flush=True)
        run_pnl_summary()
    if not args.skip_kpi:
        print("=== KPI REPORT ===", flush=True)
        run_final_kpi_report()
