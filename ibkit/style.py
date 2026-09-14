"""
ibkit.style — one design system for every workbook in the IB portfolio.

Usage
-----
    from ibkit.style import Theme, THEMES, Book
    book = Book(theme=THEMES["apple"], project_no=1, project="Three-Statement Operating Model", company="Apple Inc.")
    ws = book.sheet("Model", ncols=12)                 # title band + subtitle already drawn, data starts at row 5
    book.section(ws, 6, "INCOME STATEMENT", ncols=12)
    book.year_header(ws, 5, ["FY2024A", "FY2025E"], first_col=3)
    ...
    book.cover(toc=[("Inputs", "Assumptions and scenario selector"), ...], highlights=[("Revenue FY2030E", "=Model!K7", "#,##0.0")])
    book.finish()                                     # print setup, gridlines off, footers, freeze panes
    book.save(path); export_pdf(path)

Colour conventions inside the grid follow banking practice and never change with
the brand: blue text = hard-coded input, black = formula, green = cross-sheet
link. The brand palette is used for title bands, section headers, the cover,
fills and borders, so the model reads as the company's document without
breaking the input/formula convention an analyst expects.
"""
from __future__ import annotations

import os
import subprocess
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.properties import PageSetupProperties

FONT = "Arial"


@dataclass(frozen=True)
class Theme:
    key: str
    name: str
    primary: str       # title bands, section headers, cover sidebar
    secondary: str     # subtitles, secondary headers
    accent: str        # highlights, key-lever fills, chart accents
    light: str         # light tint for sub-sections and table stripes
    text_on_primary: str = "FFFFFF"
    tagline: str = ""


# Brand-associated palettes (approximations of each company's public identity).
THEMES: Dict[str, Theme] = {
    "apple":      Theme("apple", "Apple Inc.", "1D1D1F", "6E6E73", "0071E3", "F5F5F7", tagline="Cupertino, California · NASDAQ: AAPL"),
    "microsoft":  Theme("microsoft", "Microsoft", "0F3C7C", "5E5E5E", "00A4EF", "EAF3FB", tagline="NASDAQ: MSFT"),
    "tcs":        Theme("tcs", "Tata Consultancy Services", "0F2B5B", "5B6B85", "E31837", "EEF2F8", tagline="NSE: TCS"),
    "reliance":   Theme("reliance", "Reliance Industries", "0A2E8A", "5A6A8A", "E4002B", "EDF1FA", tagline="NSE: RELIANCE"),
    "hdfc":       Theme("hdfc", "HDFC Bank", "004C8F", "5C6F86", "ED232A", "E9F0F7", tagline="NSE: HDFCBANK"),
    "enbd":       Theme("enbd", "Emirates NBD", "0B2B57", "5C6E8A", "5BA4D9", "EAF0F8", tagline="DFM: EMIRATESNBD"),
    "emaar":      Theme("emaar", "Emaar Properties", "1C1C1C", "6B6B6B", "C8A951", "F4F1E8", tagline="DFM: EMAAR"),
    "dewa":       Theme("dewa", "DEWA", "005B96", "5D7A8C", "6CB33F", "E8F2F8", tagline="DFM: DEWA"),
    "salik":      Theme("salik", "Salik", "0B3B6F", "5D6F84", "F58220", "EAF0F7", tagline="DFM: SALIK"),
    "aramco":     Theme("aramco", "Saudi Aramco", "00477A", "5E7A8E", "84BD00", "E6F1F7", tagline="Tadawul: 2222"),
    "neutral":    Theme("neutral", "IB Portfolio", "1F3A5F", "5F6B7A", "C0392B", "ECF0F4", tagline="Jainam Mehta"),
}

# Grid conventions (never brand-dependent)
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_FORMULA = Font(name=FONT, size=10, color="000000")
F_LINK = Font(name=FONT, size=10, color="008000")
F_TEXT = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_NOTE = Font(name=FONT, size=9, italic=True, color="7F7F7F")
NUM = '#,##0.0;(#,##0.0);"-"'
NUM0 = '#,##0;(#,##0);"-"'
NUM2 = '#,##0.00;(#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
MULT = '0.0x'
TOTAL_BORDER = Border(top=Side(style="thin", color="000000"))
DOUBLE_BORDER = Border(top=Side(style="thin", color="000000"), bottom=Side(style="double", color="000000"))


class Book:
    """A themed workbook with consistent sheets, cover and print setup."""

    def __init__(self, theme: Theme, project_no: int, project: str, company: str,
                 units: str = "US$ billions unless stated", as_of: str = "", author: str = "Jainam Mehta · github.com/jainammehta1215/ib-portfolio"):
        self.theme, self.project_no, self.project, self.company = theme, project_no, project, company
        self.units, self.as_of, self.author = units, as_of, author
        self.wb = Workbook(); self.wb.remove(self.wb.active)
        self.sheets: Dict[str, Tuple] = {}
        self.fill_primary = PatternFill("solid", fgColor=theme.primary)
        self.fill_secondary = PatternFill("solid", fgColor=theme.secondary)
        self.fill_accent = PatternFill("solid", fgColor=theme.accent)
        self.fill_light = PatternFill("solid", fgColor=theme.light)
        self.f_band = Font(name=FONT, size=16, bold=True, color=theme.text_on_primary)
        self.f_band_sub = Font(name=FONT, size=10, color=theme.text_on_primary)
        self.f_section = Font(name=FONT, size=10, bold=True, color=theme.text_on_primary)
        self.f_header = Font(name=FONT, size=10, bold=True, color=theme.text_on_primary)
        self.f_subsection = Font(name=FONT, size=10, bold=True, color=theme.primary)
        self.f_kicker = Font(name=FONT, size=9, bold=True, color=theme.accent)

    # -- sheet scaffolding ----------------------------------------------------
    def sheet(self, name: str, ncols: int, label_width: float = 46, col_width: float = 12,
              subtitle: Optional[str] = None, first_col: int = 3):
        ws = self.wb.create_sheet(name)
        ws.column_dimensions["A"].width = 2.5
        ws.column_dimensions["B"].width = label_width
        for j in range(first_col, ncols + 1):
            ws.column_dimensions[L(j)].width = col_width
        # title band rows 1-2
        for r in (1, 2):
            for j in range(1, ncols + 1):
                ws.cell(r, j).fill = self.fill_primary
        ws.row_dimensions[1].height = 30; ws.row_dimensions[2].height = 18
        ws["B1"] = f"{self.company} — {self.project}"; ws["B1"].font = self.f_band; ws["B1"].alignment = Alignment(vertical="center")
        ws["B2"] = subtitle or f"{name}  ·  {self.units}  ·  Project {self.project_no} of 20"; ws["B2"].font = self.f_band_sub
        c = ws.cell(1, ncols, self.theme.name.upper()); c.font = Font(name=FONT, size=9, bold=True, color=self.theme.text_on_primary); c.alignment = Alignment(horizontal="right", vertical="center")
        ws.sheet_view.showGridLines = False
        self.sheets[name] = (ws, ncols)
        return ws

    def year_header(self, ws, row: int, labels: Sequence[str], first_col: int = 3, label: str = ""):
        ws.cell(row, 2, label).font = F_BOLD
        for j, lab in enumerate(labels):
            c = ws.cell(row, first_col + j, lab); c.font = self.f_header; c.fill = self.fill_secondary
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row].height = 18

    def section(self, ws, row: int, text: str, ncols: int, first_col: int = 2):
        for j in range(first_col, ncols + 1):
            ws.cell(row, j).fill = self.fill_primary
        c = ws.cell(row, first_col, text); c.font = self.f_section
        ws.row_dimensions[row].height = 16

    def subsection(self, ws, row: int, text: str, ncols: int, first_col: int = 2):
        for j in range(first_col, ncols + 1):
            ws.cell(row, j).fill = self.fill_light
        c = ws.cell(row, first_col, text); c.font = self.f_subsection

    def note(self, ws, row: int, text: str, ncols: int, height: float = 26, first_col: int = 2):
        c = ws.cell(row, first_col, text); c.font = F_NOTE; c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=row, start_column=first_col, end_row=row, end_column=ncols)
        ws.row_dimensions[row].height = height

    def key_cell(self, cell, value, fmt: Optional[str] = None):
        cell.value = value; cell.font = Font(name=FONT, size=11, bold=True, color="0000FF"); cell.fill = self.fill_accent
        if fmt: cell.number_format = fmt
        cell.alignment = Alignment(horizontal="center")

    # -- cover ----------------------------------------------------------------
    def cover(self, toc: List[Tuple[str, str]], highlights: List[Tuple[str, str, str]] = (),
              blurb: str = "", method: List[str] = (), sources: List[str] = (), legend: bool = True):
        ws = self.wb.create_sheet("Cover", 0)
        ws.sheet_view.showGridLines = False
        widths = {"A": 3, "B": 30, "C": 62, "D": 3, "E": 24, "F": 14}
        for k, v in widths.items(): ws.column_dimensions[k].width = v
        # sidebar in brand primary
        for r in range(1, 46):
            ws.cell(r, 1).fill = self.fill_primary
        ws.row_dimensions[1].height = 8
        ws["B3"] = f"IB PORTFOLIO · PROJECT {self.project_no:02d} OF 20"; ws["B3"].font = self.f_kicker
        ws["B4"] = self.company; ws["B4"].font = Font(name=FONT, size=24, bold=True, color=self.theme.primary)
        ws["B5"] = self.project; ws["B5"].font = Font(name=FONT, size=15, color=self.theme.secondary)
        ws["B6"] = f"{self.theme.tagline}   ·   {self.units}" + (f"   ·   as of {self.as_of}" if self.as_of else ""); ws["B6"].font = F_NOTE
        ws.row_dimensions[4].height = 32; ws.row_dimensions[5].height = 22
        # accent rule
        for j in range(2, 7): ws.cell(7, j).fill = self.fill_accent
        ws.row_dimensions[7].height = 3
        r = 9
        if blurb:
            ws.cell(r, 2, "What this is").font = self.f_subsection
            c = ws.cell(r, 3, blurb); c.font = F_TEXT; c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6); ws.row_dimensions[r].height = 15 * max(2, len(blurb) // 95 + 1); r += 2
        if method:
            ws.cell(r, 2, "How it is built").font = self.f_subsection
            for i, m in enumerate(method):
                c = ws.cell(r + i, 3, "•  " + m); c.font = F_TEXT; c.alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=r + i, start_column=3, end_row=r + i, end_column=6); ws.row_dimensions[r + i].height = 15 * max(1, len(m) // 95 + 1)
            r += len(method) + 1
        ws.cell(r, 2, "Contents").font = self.f_subsection
        for i, (nm, desc) in enumerate(toc):
            ws.cell(r + i, 3, nm).font = F_BOLD; ws.cell(r + i, 3).hyperlink = f"#'{nm}'!A1"
            d = ws.cell(r + i, 3); d.value = f"{nm}  —  {desc}"; d.font = F_TEXT
            ws.merge_cells(start_row=r + i, start_column=3, end_row=r + i, end_column=6)
        r += len(toc) + 1
        if highlights:
            ws.cell(r, 2, "Key outputs (live)").font = self.f_subsection
            for i, (lab, formula, fmt) in enumerate(highlights):
                ws.cell(r + i, 3, lab).font = F_TEXT
                c = ws.cell(r + i, 5, formula); c.font = F_LINK; c.number_format = fmt; c.alignment = Alignment(horizontal="right")
                for j in (3, 4, 5): ws.cell(r + i, j).fill = self.fill_light
            r += len(highlights) + 1
        if legend:
            ws.cell(r, 2, "Colour code").font = self.f_subsection
            ws.cell(r, 3, "Blue = hard-coded input").font = F_INPUT
            ws.cell(r + 1, 3, "Black = formula").font = F_FORMULA
            ws.cell(r + 2, 3, "Green = link to another sheet").font = F_LINK
            c = ws.cell(r + 3, 3, "Accent fill = scenario selector / key levers"); c.font = F_TEXT; c.fill = self.fill_accent
            r += 5
        if sources:
            ws.cell(r, 2, "Sources").font = self.f_subsection
            for i, s in enumerate(sources):
                c = ws.cell(r + i, 3, s); c.font = F_NOTE; c.alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=r + i, start_column=3, end_row=r + i, end_column=6); ws.row_dimensions[r + i].height = 14 * max(1, len(s) // 95 + 1)
            r += len(sources) + 1
        ws.cell(r, 2, "Prepared by").font = self.f_subsection; ws.cell(r, 3, self.author).font = F_TEXT
        self.sheets["Cover"] = (ws, 6)
        return ws

    # -- finishing ------------------------------------------------------------
    def finish(self, freeze: Dict[str, str] = None, repeat_rows: Dict[str, str] = None, portrait: Iterable[str] = ("Cover",)):
        freeze = freeze or {}; repeat_rows = repeat_rows or {}
        for name, (ws, ncols) in self.sheets.items():
            ws.page_setup.orientation = "portrait" if name in portrait else "landscape"
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
            ws.page_margins.left = ws.page_margins.right = 0.4; ws.page_margins.top = 0.5; ws.page_margins.bottom = 0.55
            ws.print_options.horizontalCentered = True
            ws.oddHeader.left.text = f"&\"Arial,Bold\"&8{self.company} — {self.project}"
            ws.oddHeader.right.text = f"&\"Arial\"&8{name}"
            ws.oddFooter.left.text = f"&\"Arial\"&8{self.author}"
            ws.oddFooter.center.text = "&\"Arial\"&8Blue = input · Black = formula · Green = link"
            ws.oddFooter.right.text = "&\"Arial\"&8Page &P of &N"
            max_row = ws.max_row; ws.print_area = f"A1:{L(ncols)}{max_row}"
            if name in freeze: ws.freeze_panes = freeze[name]
            if name in repeat_rows: ws.print_title_rows = repeat_rows[name]
        return self.wb

    def save(self, path: str):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        self.wb.save(path); return path


# --------------------------------------------------------------------------- #
# Export helpers
# --------------------------------------------------------------------------- #
def recalc(path: str, timeout: int = 90) -> dict:
    import json
    out = subprocess.run(["python", "/mnt/skills/public/xlsx/scripts/recalc.py", path, str(timeout)], capture_output=True, text=True)
    try:
        return json.loads(out.stdout)
    except Exception:
        return {"status": "unknown", "raw": out.stdout[-500:], "err": out.stderr[-500:]}


def export_pdf(path: str) -> str:
    outdir = os.path.dirname(path) or "."
    subprocess.run(["python", "/mnt/skills/public/xlsx/scripts/office/soffice.py", "--headless", "--convert-to", "pdf", "--outdir", outdir, path],
                   capture_output=True, text=True, timeout=180)
    return os.path.splitext(path)[0] + ".pdf"


def preview_png(pdf_path: str, page: int = 0, dpi: int = 80) -> str:
    import pymupdf
    doc = pymupdf.open(pdf_path); out = os.path.splitext(pdf_path)[0] + f"_p{page + 1}.png"
    doc[page].get_pixmap(dpi=dpi).save(out); return out
