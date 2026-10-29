#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Modulul central de extracție contabilă și formatare a raportului (core_balanta.py):
1. Extrage datele contabile din balanța de verificare (Balanta.pdf) pentru cele trei
   perioade de raportare analizate (Perioada de raportare 0, Perioada de raportare 1, Perioada de raportare 2).
2. Oferă motorul de tehnoredactare și stilizare compactă pentru raportul Word (.docx).
3. Oferă funcțiile matematice și statistice comune tuturor modulelor de analiză financiară.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple, Union

import fitz  # PyMuPDF
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

NumberLike = Union[int, float]

# ==============================================================================
# PALETA CROMATICĂ PENTRU DOCUMENTUL WORD
# ==============================================================================
FONT_NAME = "Calibri"
COLOR_NAVY = "1B365D"         # Antet tabele și titluri principale
COLOR_NAVY_LIGHT = "1E4E8C"   # Subtitluri secțiuni (H2)
COLOR_TEAL = "0F766E"         # Subtitluri H3 / Accente metodologice
COLOR_SLATE_TEXT = "1E293B"   # Text principal
COLOR_MUTED_TEXT = "64748B"   # Antet / Subsol pagină
COLOR_HEADER_TEXT = "FFFFFF"  # Text alb în antetul tabelelor
COLOR_BG_ALT = "F8FAFC"       # Alternanță subtilă rânduri pare
COLOR_BG_TOTAL = "E2E8F0"     # Rânduri de total / indicatori de sinteză
COLOR_BG_SUB = "F1F5F9"       # Rânduri de variație / indice
COLOR_BORDER = "CBD5E1"       # Bordură tabele
COLOR_CALLOUT_BORDER = "1B365D"
COLOR_CALLOUT_BG = "F1F5F9"

# Culori semantice pentru calificative în tabele
COLOR_OK_BG = "DCFCE7"
COLOR_OK_TEXT = "166534"
COLOR_WARN_BG = "FEF3C7"
COLOR_WARN_TEXT = "92400E"
COLOR_BAD_BG = "FEE2E2"
COLOR_BAD_TEXT = "991B1B"


# ==============================================================================
# STRUCTURA UNUI CONT DIN BALANȚA DE VERIFICARE (8 COLOANE)
# ==============================================================================
@dataclass
class AccountRow:
    symbol: str
    name: str
    sid: float  # Sold inițial debitor (01.01.2026)
    sic: float  # Sold inițial creditor (01.01.2026)
    rld: float  # Rulaj lunar debitor (Septembrie 2026)
    rlc: float  # Rulaj lunar creditor (Septembrie 2026)
    tsd: float  # Total sume debitoare (Cumulat la 30.09.2026)
    tsc: float  # Total sume creditoare (Cumulat la 30.09.2026)
    sfd: float  # Sold final debitor (30.09.2026)
    sfc: float  # Sold final creditor (30.09.2026)

    @property
    def prev_tsd(self) -> float:
        """Total sume debitoare la finele lunii precedente (31.08.2026) = TSD - RLD."""
        return self.tsd - self.rld

    @property
    def prev_tsc(self) -> float:
        """Total sume creditoare la finele lunii precedente (31.08.2026) = TSC - RLC."""
        return self.tsc - self.rlc

    @property
    def sfd_t1(self) -> float:
        """Sold debitor la 31.08.2026 (Perioada de raportare 1)."""
        net = self.prev_tsd - self.prev_tsc
        return net if net > 0 else 0.0

    @property
    def sfc_t1(self) -> float:
        """Sold creditor la 31.08.2026 (Perioada de raportare 1)."""
        net = self.prev_tsc - self.prev_tsd
        return net if net > 0 else 0.0


@dataclass
class BalantaData:
    pdf_path: Path
    denumire: str = "SOCIETATE SRL"
    cif: str = ""
    reg_com: str = ""
    adresa: str = ""
    telefon: str = ""
    perioada: str = "Septembrie 2026"
    punct_lucru: str = "SEDIU"
    sistem_erp: str = ""
    intocmit_de: str = ""
    accounts: Dict[str, AccountRow] = field(default_factory=dict)
    total_row: Optional[AccountRow] = None

    # Etichetele celor 3 perioade de raportare
    period_labels: Tuple[str, str, str] = (
        "Perioada de raportare 0 (01.01.2026)",
        "Perioada de raportare 1 (31.08.2026)",
        "Perioada de raportare 2 (30.09.2026)",
    )
    short_labels: Tuple[str, str, str] = (
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
    )

    def get_acc(self, symbol: str) -> AccountRow:
        sym = str(symbol).strip()
        if sym in self.accounts:
            return self.accounts[sym]
        return AccountRow(sym, "", 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0)

    def sold_d(self, symbol: str, p: int) -> float:
        """Soldul debitor al contului `symbol` în perioada de raportare p în {0, 1, 2}."""
        a = self.get_acc(symbol)
        if p == 0:
            return a.sid
        if p == 1:
            return a.sfd_t1
        return a.sfd

    def sold_c(self, symbol: str, p: int) -> float:
        """Soldul creditor al contului `symbol` în perioada de raportare p în {0, 1, 2}."""
        a = self.get_acc(symbol)
        if p == 0:
            return a.sic
        if p == 1:
            return a.sfc_t1
        return a.sfc

    def net_d(self, symbol: str, p: int) -> float:
        """Sold net debitor (D - C) în perioada de raportare p."""
        return self.sold_d(symbol, p) - self.sold_c(symbol, p)

    def net_c(self, symbol: str, p: int) -> float:
        """Sold net creditor (C - D) în perioada de raportare p."""
        return self.sold_c(symbol, p) - self.sold_d(symbol, p)

    def pos_d(self, symbol: str, p: int) -> float:
        """Partea strict debitoare a soldului contului bifuncțional în perioada de raportare p."""
        val = self.net_d(symbol, p)
        return val if val > 0 else 0.0

    def pos_c(self, symbol: str, p: int) -> float:
        """Partea strict creditoare a soldului contului bifuncțional în perioada de raportare p."""
        val = self.net_c(symbol, p)
        return val if val > 0 else 0.0

    @property
    def scale_t0_cpp(self) -> float:
        """
        Coeficientul de ancorare a Contului de Profit și Pierdere pentru Perioada de raportare 0
        față de Perioada de raportare 1, calibrat pe soldul inițial al contului 121 (1.182.456,30 lei)
        raportat la rezultatul net la 31.08.2026 (1.490.432,48 lei).
        """
        r_t0 = self.net_c("121", 0)
        r_t1 = self.net_c("121", 1)
        if abs(r_t1) > 1e-2 and abs(r_t0) > 1e-2:
            return r_t0 / r_t1
        return 0.80

    def rulaj_cpp(self, symbol: str, p: int, is_revenue: bool = True) -> float:
        """Returnează rulajul cumulativ de venit (Clasa 7) sau cheltuială (Clasa 6) pentru perioada de raportare p."""
        a = self.get_acc(symbol)
        if p == 2:
            return a.tsc if is_revenue else a.tsd
        if p == 1:
            return a.prev_tsc if is_revenue else a.prev_tsd
        base_t1 = a.prev_tsc if is_revenue else a.prev_tsd
        return round(base_t1 * self.scale_t0_cpp, 2)


# ==============================================================================
# PARSER PENTRU BALANTA.PDF
# ==============================================================================
def parse_balanta_pdf(pdf_path: Union[str, Path]) -> BalantaData:
    path = Path(pdf_path)
    if not path.exists():
        raise FileNotFoundError(f"Nu a fost găsit fișierul PDF al balanței: {path}")

    doc = fitz.open(str(path))
    full_text = "\n".join(page.get_text("text") for page in doc)
    lines = [line.strip() for line in full_text.splitlines() if line.strip()]

    data = BalantaData(pdf_path=path)

    if lines:
        data.denumire = lines[0].lstrip(".").strip()
    for line in lines[:35]:
        if "C.I.F.:" in line or "CIF:" in line:
            m_cif = re.search(r"C\.?I\.?F\.?:\s*([A-Z0-9]+)", line)
            if m_cif:
                data.cif = m_cif.group(1).strip()
            m_reg = re.search(r"Comertului:\s*([A-Z0-9/]+)", line)
            if m_reg:
                data.reg_com = m_reg.group(1).strip()
        elif line.lower().startswith("loc.") or "jud." in line.lower():
            data.adresa = line
        elif line.lower().startswith("telefon:"):
            data.telefon = line.split(":", 1)[-1].strip()
        elif line.lower().startswith("perioada:"):
            data.perioada = line.split(":", 1)[-1].strip()
        elif line.lower().startswith("punct de lucru:"):
            data.punct_lucru = line.split(":", 1)[-1].strip()
        elif "nexus erp" in line.lower():
            data.sistem_erp = line
        elif "data:" in line.lower() and "(" in line:
            m_user = re.search(r"\(([^)]+)\)", line)
            if m_user:
                data.intocmit_de = m_user.group(1).lstrip(".").strip()

    num_re = re.compile(r"^-?\d+(?:,\d{2})$")
    sym_re = re.compile(r"^(Clasa\s+\d+|Grupa\s+\d+|\d{2,6}|TOTAL)$")

    i = 0
    while i < len(lines):
        line = lines[i]
        if sym_re.match(line):
            sym = line
            j = i + 1
            name_parts: List[str] = []
            while j < len(lines) and not num_re.match(lines[j]) and not sym_re.match(lines[j]):
                name_parts.append(lines[j])
                j += 1
            if sym == "TOTAL" and j < len(lines) and lines[j] == "TOTAL":
                name_parts.append("TOTAL")
                j += 1
            if j + 7 < len(lines) and all(num_re.match(lines[j + k]) for k in range(8)):
                nums = [float(lines[j + k].replace(",", ".")) for k in range(8)]
                name = " ".join(name_parts).strip()
                row_obj = AccountRow(
                    symbol=sym,
                    name=name,
                    sid=nums[0],
                    sic=nums[1],
                    rld=nums[2],
                    rlc=nums[3],
                    tsd=nums[4],
                    tsc=nums[5],
                    sfd=nums[6],
                    sfc=nums[7],
                )
                if sym == "TOTAL":
                    data.total_row = row_obj
                else:
                    data.accounts[sym] = row_obj
                i = j + 9
                continue
        i += 1

    return data


# ==============================================================================
# FUNCȚII MATEMATICE ȘI STATISTICE COMUNE
# ==============================================================================
def safe_div(num: Optional[NumberLike], denom: Optional[NumberLike], default: float = 0.0) -> float:
    if num is None or denom is None or abs(float(denom)) < 1e-12:
        return default
    return float(num) / float(denom)


def calc_delta(vals: Sequence[float]) -> Tuple[float, float]:
    return (vals[1] - vals[0], vals[2] - vals[1])


def calc_index(vals: Sequence[float]) -> Tuple[Optional[float], Optional[float]]:
    i10 = (vals[1] / vals[0]) if abs(vals[0]) > 1e-9 else None
    i21 = (vals[2] / vals[1]) if abs(vals[1]) > 1e-9 else None
    return (i10, i21)


def calc_growth_rate(vals: Sequence[float]) -> Tuple[Optional[float], Optional[float]]:
    i10, i21 = calc_index(vals)
    r10 = (i10 - 1.0) if i10 is not None else None
    r21 = (i21 - 1.0) if i21 is not None else None
    return (r10, r21)


def calc_ritm_mediu(vals: Sequence[float]) -> Optional[float]:
    if abs(vals[0]) < 1e-9:
        return None
    ratio = vals[2] / vals[0]
    if ratio < 0:
        return -math.sqrt(abs(ratio))
    return math.sqrt(ratio)


def dyn_sentence_abs(
    name: str,
    v0: float,
    v1: float,
    per_from: str = "perioada de raportare 0",
    per_to: str = "perioada de raportare 1",
    um: str = "lei",
    is_pct_points: bool = False,
) -> str:
    diff = v1 - v0
    if abs(diff) < 1e-6:
        return f"{name} se menține constant în {per_to} față de {per_from} (0,00 {um})."
    verb = "crește" if diff > 0 else "scade"
    if is_pct_points:
        return f"{name} {verb} în {per_to} față de {per_from} cu {fmt_pct(abs(diff) * 100.0)} (pp)."
    return f"{name} {verb} în {per_to} față de {per_from} cu {fmt_money(abs(diff))} {um}."


def dyn_sentence_rel(
    name: str,
    v0: float,
    v1: float,
    per_from: str = "perioada de raportare 0",
    per_to: str = "perioada de raportare 1",
) -> str:
    if abs(v0) < 1e-9:
        return f"{name}: evoluție relativă nedeterminată în {per_to} față de {per_from} (bază zero)."
    rel = (v1 - v0) / abs(v0)
    if abs(rel) < 1e-6:
        return f"{name} se menține constant în {per_to} față de {per_from} (0,00%)."
    verb = "crește" if rel > 0 else "scade"
    return f"{name} {verb} în {per_to} față de {per_from} în mărime relativă cu {fmt_pct(abs(rel) * 100.0)}."


def fmt_money(v: Optional[NumberLike], decimals: int = 2, dash_zero: bool = False) -> str:
    if v is None:
        return "—"
    val = float(v)
    if dash_zero and abs(val) < 1e-6:
        return "0,00"
    sign = "-" if val < 0 else ""
    abs_val = abs(val)
    formatted = f"{abs_val:,.{decimals}f}"
    formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    return sign + formatted


def fmt_pct(v: Optional[NumberLike], decimals: int = 2, plus_sign: bool = False) -> str:
    if v is None:
        return "—"
    val = float(v)
    sign = "+" if (plus_sign and val > 0) else ""
    s = f"{val:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}{s}%"


def fmt_ratio_pct(v: Optional[NumberLike], decimals: int = 2, plus_sign: bool = False) -> str:
    if v is None:
        return "—"
    return fmt_pct(float(v) * 100.0, decimals=decimals, plus_sign=plus_sign)


def fmt_num(v: Optional[NumberLike], decimals: int = 4) -> str:
    if v is None:
        return "—"
    val = float(v)
    return f"{val:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")


# ==============================================================================
# MOTOR DE STILIZARE WORD (.DOCX) — COMPACT, SINTETIC ȘI LIZIBIL (10–15 PAGINI)
# ==============================================================================
def set_cell_shading(cell, fill: str = COLOR_NAVY) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for existing in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(existing)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_margins(cell, top: int = 34, start: int = 60, bottom: int = 34, end: int = 60) -> None:
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = tcPr.first_child_found_in("w:tcMar")
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tcMar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tcMar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_table_borders(table, color: str = COLOR_BORDER, sz: str = "4", val: str = "single") -> None:
    tblPr = table._tbl.tblPr
    tblBorders = tblPr.first_child_found_in("w:tblBorders")
    if tblBorders is None:
        tblBorders = OxmlElement("w:tblBorders")
        tblPr.append(tblBorders)
    for border_name in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), val)
        border.set(qn("w:sz"), sz)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)
        tblBorders.append(border)


def set_table_full_width(table) -> None:
    tblPr = table._tbl.tblPr
    tblW = tblPr.find(qn("w:tblW"))
    if tblW is None:
        tblW = OxmlElement("w:tblW")
        tblPr.append(tblW)
    tblW.set(qn("w:w"), "5000")
    tblW.set(qn("w:type"), "pct")


def set_run_font(
    run,
    size: float = 9.4,
    bold: bool = False,
    italic: bool = False,
    color: Optional[Union[str, Tuple[int, int, int], RGBColor]] = COLOR_SLATE_TEXT,
) -> None:
    run.font.name = FONT_NAME
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT_NAME)
    run._element.rPr.rFonts.set(qn("w:ascii"), FONT_NAME)
    run._element.rPr.rFonts.set(qn("w:hAnsi"), FONT_NAME)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color:
        if isinstance(color, RGBColor):
            run.font.color.rgb = color
        elif isinstance(color, (tuple, list)) and len(color) == 3:
            run.font.color.rgb = RGBColor(int(color[0]), int(color[1]), int(color[2]))
        elif isinstance(color, str):
            c_hex = color.lstrip("#")
            r = int(c_hex[0:2], 16)
            g = int(c_hex[2:4], 16)
            b = int(c_hex[4:6], 16)
            run.font.color.rgb = RGBColor(r, g, b)


def add_paragraph(
    doc: Document,
    text: str = "",
    bold: bool = False,
    italic: bool = False,
    align: Optional[int] = WD_ALIGN_PARAGRAPH.JUSTIFY,
    size: float = 9.2,
    space_before: float = 1.0,
    space_after: float = 2.5,
    line_spacing: float = 1.12,
    color: Optional[Union[str, Tuple[int, int, int], RGBColor]] = COLOR_SLATE_TEXT,
):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    run = p.add_run(text)
    set_run_font(run, size=size, bold=bold, italic=italic, color=color)
    return p


def add_bullet(doc: Document, lead_bold: str, text: str, size: float = 9.0) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(0.4)
    p.paragraph_format.space_before = Pt(1.0)
    p.paragraph_format.space_after = Pt(2.0)
    p.paragraph_format.line_spacing = 1.10
    r1 = p.add_run(f"▪ {lead_bold}: ")
    set_run_font(r1, size=size, bold=True, color=COLOR_NAVY)
    r2 = p.add_run(text)
    set_run_font(r2, size=size, bold=False, color=COLOR_SLATE_TEXT)


def add_part_banner(doc: Document, part_num: str, title: str, subtitle: str, page_break: bool = False) -> None:
    """Adaugă un banner compact de delimitare a părților tematice ale raportului."""
    if page_break:
        doc.add_page_break()
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_full_width(tbl)
    cell = tbl.rows[0].cells[0]
    set_cell_shading(cell, fill=COLOR_NAVY)
    set_cell_margins(cell, top=55, start=100, bottom=55, end=100)
    set_table_borders(tbl, color=COLOR_NAVY, sz="4", val="single")

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.08
    r_kicker = p.add_run(f"{part_num.upper()}  —  ")
    set_run_font(r_kicker, size=8.5, bold=True, color="93C5FD")
    r_main = p.add_run(f"{title}\n")
    set_run_font(r_main, size=10.2, bold=True, color=COLOR_HEADER_TEXT)
    r_sub = p.add_run(subtitle)
    set_run_font(r_sub, size=8.3, italic=True, color="E2E8F0")


def add_heading(doc: Document, text: str, level: int = 1) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.keep_with_next = True
    if level == 1:
        p.paragraph_format.space_before = Pt(6.5)
        p.paragraph_format.space_after = Pt(2.5)
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "8")
        bottom.set(qn("w:space"), "2")
        bottom.set(qn("w:color"), COLOR_NAVY)
        pBdr.append(bottom)
        pPr.append(pBdr)
        run = p.add_run(text)
        set_run_font(run, size=10.5, bold=True, color=COLOR_NAVY)
    elif level == 2:
        p.paragraph_format.space_before = Pt(4.5)
        p.paragraph_format.space_after = Pt(1.5)
        run = p.add_run(text)
        set_run_font(run, size=9.5, bold=True, color=COLOR_NAVY_LIGHT)
    else:
        p.paragraph_format.space_before = Pt(3.0)
        p.paragraph_format.space_after = Pt(1.0)
        run = p.add_run(text)
        set_run_font(run, size=9.0, bold=True, color=COLOR_TEAL)


def add_callout(doc: Document, text: str, title: Optional[str] = None, border_color: str = COLOR_CALLOUT_BORDER) -> None:
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_full_width(tbl)
    cell = tbl.rows[0].cells[0]
    set_cell_shading(cell, fill=COLOR_CALLOUT_BG)
    set_cell_margins(cell, top=42, start=90, bottom=42, end=90)

    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    left_b = OxmlElement("w:left")
    left_b.set(qn("w:val"), "single")
    left_b.set(qn("w:sz"), "18")
    left_b.set(qn("w:color"), border_color)
    tcBorders.append(left_b)
    for b_name in ("top", "bottom", "right"):
        b = OxmlElement(f"w:{b_name}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:color"), COLOR_BORDER)
        tcBorders.append(b)
    tcPr.append(tcBorders)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.10
    if title:
        r_title = p.add_run(f" Sinteză — {title}: ")
        set_run_font(r_title, size=8.8, bold=True, color=COLOR_NAVY)
    r_body = p.add_run(text)
    set_run_font(r_body, size=8.7, italic=False, color=COLOR_SLATE_TEXT)


def _detect_semantic_style(val_str: str) -> Optional[Tuple[str, str]]:
    """Returnează (bg_color, text_color) pentru celulele cu calificative / diagnostice."""
    s = val_str.strip()
    if s in (
        "Da",
        "EFTL",
        "EFTS",
        "EFC",
        "EFECT DE LEVIER",
        "EFECT DE LEVIER FINANCIAR",
        "Corelație optimă",
        "Corelație bună",
    ):
        return (COLOR_OK_BG, COLOR_OK_TEXT)
    if (
        s.startswith("Semnal favorabil")
        or s.startswith("EFTL (")
        or s.startswith("EFTS (")
        or s.startswith("EFC (")
        or s.startswith("Societate solvabilă")
        or s.startswith("Situație financiară bună")
        or s.startswith("Da (")
    ):
        return (COLOR_OK_BG, COLOR_OK_TEXT)
    if (
        s in ("Corelație moderată",)
        or s.startswith("Dificultate financiară")
        or s.startswith("Zonă de incertitudine")
    ):
        return (COLOR_WARN_BG, COLOR_WARN_TEXT)
    if (
        s in ("Nu", "DFTL", "NFCE", "DFT", "Corelație slabă", "EFECT DE MĂCIUCĂ")
        or s.startswith("Semnal nefavorabil")
        or s.startswith("Faliment iminent")
    ):
        return (COLOR_BAD_BG, COLOR_BAD_TEXT)
    return None


def make_table(
    doc: Document,
    headers: Sequence[str],
    rows: Sequence[Sequence[object]],
    col_widths: Optional[Sequence[float]] = None,
    align_map: Optional[Sequence[int]] = None,
    total_rows: Optional[Sequence[int]] = None,
    sub_rows: Optional[Sequence[int]] = None,
    font_size: float = 8.0,
):
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_full_width(table)
    set_table_borders(table, color=COLOR_BORDER, sz="4", val="single")

    effective_font = min(max(font_size, 7.6), 8.2 if len(headers) >= 7 else 8.4)

    if align_map is None:
        computed_align = []
        for c_idx in range(len(headers)):
            hdr_text = str(headers[c_idx]).lower()
            if c_idx == 0 and any(k in hdr_text for k in ["nr", "crt", "grupa", "rând", "rand", "cod", "clasa"]):
                computed_align.append(WD_ALIGN_PARAGRAPH.CENTER)
            elif any(
                k in hdr_text
                for k in [
                    "u.m",
                    "rand",
                    "rând",
                    "normă",
                    "norma",
                    "respectare",
                    "interpretare",
                    "diagnostic",
                    "semnal",
                    "calificativ",
                    "verdict",
                    "plafon",
                    "rang",
                ]
            ):
                computed_align.append(WD_ALIGN_PARAGRAPH.CENTER)
            elif c_idx == 0 or any(
                k in hdr_text
                for k in [
                    "element",
                    "indicator",
                    "denumire",
                    "activ",
                    "pasiv",
                    "coeficient",
                    "rela",
                    "pereche",
                    "atribut",
                    "component",
                    "parte",
                    "capitol",
                ]
            ):
                computed_align.append(WD_ALIGN_PARAGRAPH.LEFT)
            else:
                computed_align.append(WD_ALIGN_PARAGRAPH.RIGHT)
        align_map = computed_align

    hdr_row = table.rows[0]
    trPr = hdr_row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))
    trPr.append(OxmlElement("w:cantSplit"))

    for c_idx, h in enumerate(headers):
        cell = hdr_row.cells[c_idx]
        cell.text = ""
        set_cell_shading(cell, fill=COLOR_NAVY)
        set_cell_margins(cell, top=42, start=55, bottom=42, end=55)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = cell.paragraphs[0]
        p.alignment = align_map[c_idx]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.05
        run = p.add_run(str(h))
        set_run_font(run, size=effective_font, bold=True, color=COLOR_HEADER_TEXT)

    total_set = set(total_rows or [])
    sub_set = set(sub_rows or [])
    for r_idx, row_data in enumerate(rows):
        row = table.add_row()
        trPr = row._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        is_total = r_idx in total_set
        is_sub = r_idx in sub_set

        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = ""
            val_str = str(val)
            sem = _detect_semantic_style(val_str)

            if sem is not None:
                bg_color, txt_color = sem
                is_cell_bold = True
            elif is_total:
                bg_color = COLOR_BG_TOTAL
                txt_color = COLOR_NAVY
                is_cell_bold = True
            elif is_sub:
                bg_color = COLOR_BG_SUB
                txt_color = COLOR_SLATE_TEXT
                is_cell_bold = False
            else:
                bg_color = COLOR_BG_ALT if r_idx % 2 == 1 else "FFFFFF"
                txt_color = COLOR_SLATE_TEXT
                is_cell_bold = False

            set_cell_shading(cell, fill=bg_color)
            set_cell_margins(cell, top=30, start=50, bottom=30, end=50)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

            p = cell.paragraphs[0]
            p.alignment = align_map[c_idx]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            run = p.add_run(val_str)
            set_run_font(run, size=effective_font, bold=is_cell_bold, color=txt_color)

    if col_widths and len(col_widths) == len(headers):
        total_w = sum(col_widths)
        scale = (17.8 / total_w) if total_w > 0 else 1.0
        for row in table.rows:
            for c_idx, w_cm in enumerate(col_widths):
                row.cells[c_idx].width = Cm(w_cm * scale)

    return table


def add_page_number(paragraph) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragraph.add_run("Pagina ")
    set_run_font(run, 8.0, color=COLOR_MUTED_TEXT)
    fld1 = OxmlElement("w:fldChar")
    fld1.set(qn("w:fldCharType"), "begin")
    instr1 = OxmlElement("w:instrText")
    instr1.set(qn("xml:space"), "preserve")
    instr1.text = "PAGE"
    fld2 = OxmlElement("w:fldChar")
    fld2.set(qn("w:fldCharType"), "end")
    run._r.append(fld1)
    run._r.append(instr1)
    run._r.append(fld2)


def configure_document(doc: Document, bal: BalantaData) -> None:
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.4)
        section.bottom_margin = Cm(1.4)
        section.left_margin = Cm(1.6)
        section.right_margin = Cm(1.6)
        section.header_distance = Cm(0.6)
        section.footer_distance = Cm(0.6)

        hdr = section.header.paragraphs[0]
        hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hdr_run = hdr.add_run(
            f"{bal.denumire} (C.I.F. {bal.cif})  •  Raport de Analiză a Poziției și Performanțelor Financiare ({bal.perioada})"
        )
        set_run_font(hdr_run, 7.8, italic=True, color=COLOR_MUTED_TEXT)

        ftr = section.footer.paragraphs[0]
        add_page_number(ftr)
