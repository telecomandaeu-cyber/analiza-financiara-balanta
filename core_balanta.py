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

import calendar
import hashlib
import math
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple, Union

import fitz  # PyMuPDF
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

NumberLike = Union[int, float]

# Valoare „nedisponibilă” (n.d.): se propagă automat prin operațiile aritmetice
# și este afișată ca „n.d.” de funcțiile de formatare. Se folosește pentru
# indicatorii de flux (CPP) în Perioada de raportare 0, care reprezintă doar
# soldurile de deschidere ale exercițiului și nu conține rulaje de venituri/cheltuieli.
ND = float("nan")


def is_nd(v: Optional[NumberLike]) -> bool:
    """True dacă valoarea lipsește (None) sau este nedisponibilă (NaN)."""
    if v is None:
        return True
    try:
        return math.isnan(float(v))
    except (TypeError, ValueError):
        return True


def all_ok(*vals: Optional[NumberLike]) -> bool:
    """True dacă toate valorile sunt disponibile (nici None, nici NaN)."""
    return not any(is_nd(v) for v in vals)


_LUNI_RO = {
    "ianuarie": 1, "februarie": 2, "martie": 3, "aprilie": 4, "mai": 5, "iunie": 6,
    "iulie": 7, "august": 8, "septembrie": 9, "octombrie": 10, "noiembrie": 11, "decembrie": 12,
}
_LUNI_RO_NUME = {v: k for k, v in _LUNI_RO.items()}


def _strip_diacritics(s: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFKD", s) if not unicodedata.combining(ch))


def parse_period_text(text: str) -> Optional[Tuple[int, int]]:
    """Extrage (luna, anul) dintr-un text de tipul „Septembrie 2026” / „Septembrie.2026” / „09.2026”."""
    t = _strip_diacritics(text or "").lower()
    m_year = re.search(r"(20\d{2}|19\d{2})", t)
    if not m_year:
        return None
    year = int(m_year.group(1))
    for name, num in _LUNI_RO.items():
        if name in t:
            return num, year
    m_num = re.search(r"\b(0?[1-9]|1[0-2])[./-](20\d{2}|19\d{2})", t)
    if m_num:
        return int(m_num.group(1)), int(m_num.group(2))
    return None

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
    """
    Un rând din balanța de verificare cu 8 coloane. „Total sume” include soldul inițial
    (TSD = SID + rulaje debitoare cumulate de la începutul anului), astfel încât
    SFD - SFC = TSD - TSC.
    """

    symbol: str
    name: str
    sid: float  # Sold inițial debitor (la începutul exercițiului)
    sic: float  # Sold inițial creditor (la începutul exercițiului)
    rld: float  # Rulaj debitor al lunii de raportare
    rlc: float  # Rulaj creditor al lunii de raportare
    tsd: float  # Total sume debitoare (cumulat de la începutul anului, inclusiv SID)
    tsc: float  # Total sume creditoare (cumulat de la începutul anului, inclusiv SIC)
    sfd: float  # Sold final debitor (la sfârșitul lunii de raportare)
    sfc: float  # Sold final creditor (la sfârșitul lunii de raportare)

    @property
    def prev_tsd(self) -> float:
        """Total sume debitoare la finele lunii precedente = TSD - RLD."""
        return self.tsd - self.rld

    @property
    def prev_tsc(self) -> float:
        """Total sume creditoare la finele lunii precedente = TSC - RLC."""
        return self.tsc - self.rlc

    @property
    def net_t1(self) -> float:
        """Sold net debitor (D - C) la finele lunii precedente (Perioada de raportare 1)."""
        return self.prev_tsd - self.prev_tsc

    @property
    def sfd_t1(self) -> float:
        """Sold debitor la finele lunii precedente (Perioada de raportare 1)."""
        net = self.net_t1
        return net if net > 0 else 0.0

    @property
    def sfc_t1(self) -> float:
        """Sold creditor la finele lunii precedente (Perioada de raportare 1)."""
        net = -self.net_t1
        return net if net > 0 else 0.0


@dataclass
class PeriodInfo:
    """Datele calendaristice ale celor trei momente de analiză, deduse din balanță."""

    year: int = 2026
    month: int = 9

    @property
    def months(self) -> Tuple[int, int, int]:
        """Numărul de luni cumulate în rulajele fiecărei perioade (P0 = 0 luni)."""
        return (0, max(self.month - 1, 0), self.month)

    @property
    def days(self) -> Tuple[int, int, int]:
        """Numărul de zile calendaristice acoperite de rulajele cumulate ale fiecărei perioade."""
        d1 = sum(calendar.monthrange(self.year, m)[1] for m in range(1, self.month))
        d2 = d1 + calendar.monthrange(self.year, self.month)[1]
        return (0, d1, d2)

    @property
    def days_in_year(self) -> int:
        return 366 if calendar.isleap(self.year) else 365

    def annual_factor(self, p: int) -> float:
        """Factorul de anualizare al rulajelor cumulate (365 / zile); NaN dacă perioada nu are rulaje."""
        d = self.days[p]
        return self.days_in_year / d if d > 0 else ND

    @property
    def dates(self) -> Tuple[str, str, str]:
        d0 = f"01.01.{self.year}"
        if self.month > 1:
            last1 = calendar.monthrange(self.year, self.month - 1)[1]
            d1 = f"{last1:02d}.{self.month - 1:02d}.{self.year}"
        else:
            d1 = d0
        last2 = calendar.monthrange(self.year, self.month)[1]
        d2 = f"{last2:02d}.{self.month:02d}.{self.year}"
        return (d0, d1, d2)

    @property
    def month_name(self) -> str:
        return _LUNI_RO_NUME.get(self.month, "").capitalize()

    @property
    def flow_labels(self) -> Tuple[str, str, str]:
        """Descrierea intervalului cumulat acoperit de rulajele fiecărei perioade."""
        d = self.dates
        return (
            "fără rulaje (solduri de deschidere)",
            f"01.01–{d[1][:5]} ({self.months[1]} luni)" if self.months[1] > 0 else "fără rulaje",
            f"01.01–{d[2][:5]} ({self.months[2]} luni)",
        )


@dataclass
class BalantaData:
    pdf_path: Path
    denumire: str = "SOCIETATE SRL"
    cif: str = ""
    reg_com: str = ""
    adresa: str = ""
    telefon: str = ""
    perioada: str = ""
    punct_lucru: str = "SEDIU"
    sistem_erp: str = ""
    intocmit_de: str = ""
    accounts: Dict[str, AccountRow] = field(default_factory=dict)
    total_row: Optional[AccountRow] = None
    period: PeriodInfo = field(default_factory=PeriodInfo)
    warnings: List[str] = field(default_factory=list)

    short_labels: Tuple[str, str, str] = (
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
    )

    @property
    def period_labels(self) -> Tuple[str, str, str]:
        d = self.period.dates
        return (
            f"Perioada de raportare 0 ({d[0]})",
            f"Perioada de raportare 1 ({d[1]})",
            f"Perioada de raportare 2 ({d[2]})",
        )

    @property
    def col_labels(self) -> Tuple[str, str, str]:
        """Etichete scurte de coloană pentru tabele: „P0 (01.01.2026)” etc."""
        d = self.period.dates
        return (f"P0 ({d[0]})", f"P1 ({d[1]})", f"P2 ({d[2]})")

    # ------------------------------------------------------------------ conturi
    def has(self, symbol: str) -> bool:
        return str(symbol).strip() in self.accounts

    def children(self, symbol: str) -> List[str]:
        """Conturile de grad imediat inferior (ex.: 442 -> 4423, 4424, 4428)."""
        sym = str(symbol).strip()
        n = len(sym) + 1
        return sorted(k for k in self.accounts if len(k) == n and k.isdigit() and k.startswith(sym))

    def get_acc(self, symbol: str) -> AccountRow:
        sym = str(symbol).strip()
        if sym in self.accounts:
            return self.accounts[sym]
        return AccountRow(sym, "", 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0)

    def net_d(self, symbol: str, p: int) -> float:
        """Sold net debitor (D - C) al contului în perioada de raportare p ∈ {0, 1, 2}."""
        sym = str(symbol).strip()
        if sym not in self.accounts:
            kids = self.children(sym)
            return sum(self.net_d(k, p) for k in kids) if kids else 0.0
        a = self.accounts[sym]
        if p == 0:
            return a.sid - a.sic
        if p == 1:
            return a.net_t1
        return a.sfd - a.sfc

    def net_c(self, symbol: str, p: int) -> float:
        """Sold net creditor (C - D) al contului în perioada de raportare p."""
        return -self.net_d(symbol, p)

    def sold_d(self, symbol: str, p: int) -> float:
        val = self.net_d(symbol, p)
        return val if val > 0 else 0.0

    def sold_c(self, symbol: str, p: int) -> float:
        val = self.net_c(symbol, p)
        return val if val > 0 else 0.0

    pos_d = sold_d
    pos_c = sold_c

    # ------------------------------------------------------------------ rulaje
    def flow_d(self, symbol: str, p: int) -> float:
        """
        Rulajul debitor cumulat de la începutul anului (fără soldul inițial) până la
        sfârșitul perioadei p. Pentru P0 (solduri de deschidere) rezultatul este n.d.
        Dacă un cont lipsește din balanță se însumează conturile de grad inferior.
        """
        if p == 0:
            return ND
        sym = str(symbol).strip()
        if sym not in self.accounts:
            kids = self.children(sym)
            return sum(self.flow_d(k, p) for k in kids) if kids else 0.0
        a = self.accounts[sym]
        base = a.tsd if p == 2 else a.prev_tsd
        return base - a.sid

    def flow_c(self, symbol: str, p: int) -> float:
        """Rulajul creditor cumulat de la începutul anului (fără soldul inițial) până la sfârșitul perioadei p."""
        if p == 0:
            return ND
        sym = str(symbol).strip()
        if sym not in self.accounts:
            kids = self.children(sym)
            return sum(self.flow_c(k, p) for k in kids) if kids else 0.0
        a = self.accounts[sym]
        base = a.tsc if p == 2 else a.prev_tsc
        return base - a.sic

    def synthetic_accounts(self) -> List[str]:
        """Conturile sintetice de gradul I (3 cifre), plus conturile de 4 cifre fără părinte în balanță."""
        syn = sorted(k for k in self.accounts if k.isdigit() and len(k) == 3)
        orphans = sorted(
            k for k in self.accounts if k.isdigit() and len(k) == 4 and k[:3] not in self.accounts
        )
        return syn + orphans

    def leaf_accounts(self) -> List[str]:
        """
        Conturile la nivelul cel mai detaliat care reconstituie integral soldul fiecărui
        cont sintetic. Conturile bifuncționale (ex. 442, 409, 518) trebuie împărțite pe
        analitice/sintetice de gradul II, deoarece soldul sintetic este compensat (net).
        """
        leaves: List[str] = []
        for syn in self.synthetic_accounts():
            kids = self.children(syn)
            if kids and len(syn) == 3:
                ok = all(
                    abs(sum(self.net_d(k, p) for k in kids) - self.net_d(syn, p)) < 0.05 for p in (0, 1, 2)
                )
                if ok:
                    leaves.extend(kids)
                    continue
            leaves.append(syn)
        return leaves


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
                # După cele 8 coloane numerice, balanța repetă simbolul contului (coloana din dreapta).
                skip = 8
                if j + 8 < len(lines) and lines[j + 8] in (sym, sym.split()[-1]):
                    skip = 9
                i = j + skip
                continue
        i += 1

    # Perioada de raportare (luna și anul) — din antet sau din titlul paginilor.
    per = parse_period_text(data.perioada)
    if per is None:
        for line in lines[:60]:
            if "balanta" in _strip_diacritics(line).lower():
                per = parse_period_text(line)
                if per:
                    break
    if per is None:
        data.warnings.append(
            "Perioada de raportare nu a putut fi identificată din antetul balanței; s-a presupus luna 12 a anului curent."
        )
        import datetime as _dt

        per = (12, _dt.date.today().year)
    data.period = PeriodInfo(year=per[1], month=per[0])
    if not data.perioada:
        data.perioada = f"{data.period.month_name} {data.period.year}"

    _validate_balanta(data)
    return data


def _validate_balanta(data: BalantaData) -> None:
    """Verifică egalitățile fundamentale ale balanței și consemnează eventualele abateri."""
    if not data.accounts:
        raise ValueError("Nu au fost identificate conturi în balanța de verificare (format PDF nerecunoscut).")
    t = data.total_row
    if t is not None:
        for lbl, d, c in (("solduri inițiale", t.sid, t.sic), ("total sume", t.tsd, t.tsc), ("solduri finale", t.sfd, t.sfc)):
            if abs(d - c) > 0.05:
                data.warnings.append(
                    f"Balanța nu este echilibrată la {lbl}: debit {fmt_money(d)} lei ≠ credit {fmt_money(c)} lei."
                )
    for sym, a in data.accounts.items():
        if abs((a.tsd - a.tsc) - (a.sfd - a.sfc)) > 0.05:
            data.warnings.append(
                f"Contul {sym}: soldul final nu corespunde cu totalul sumelor (TSD - TSC ≠ SFD - SFC)."
            )


# ==============================================================================
# FUNCȚII MATEMATICE ȘI STATISTICE COMUNE
# ==============================================================================
def safe_div(num: Optional[NumberLike], denom: Optional[NumberLike], default: float = 0.0) -> float:
    """Împărțire protejată: n.d. dacă un termen este n.d.; `default` dacă numitorul este zero."""
    if is_nd(num) or is_nd(denom):
        return ND
    if abs(float(denom)) < 1e-12:
        return default
    return float(num) / float(denom)


def ratio_or_nd(num: Optional[NumberLike], denom: Optional[NumberLike]) -> float:
    """Raport care devine n.d. (nu 0) când numitorul este zero — pentru rate fără sens economic la bază nulă."""
    return safe_div(num, denom, default=ND)


def tup3(fn) -> Tuple[float, float, float]:
    """Construiește un triplet (P0, P1, P2) aplicând fn(p)."""
    return (fn(0), fn(1), fn(2))


def calc_delta(vals: Sequence[float]) -> Tuple[float, float]:
    return (vals[1] - vals[0], vals[2] - vals[1])


def calc_index(vals: Sequence[float]) -> Tuple[Optional[float], Optional[float]]:
    """Indicii de dinamică I(P1/P0) și I(P2/P1); None dacă baza este nulă, negativă sau n.d."""

    def _idx(a: float, b: float) -> Optional[float]:
        if not all_ok(a, b) or abs(a) < 1e-9:
            return None
        # Un indice calculat pe o bază negativă nu are interpretare economică.
        if a < 0:
            return None
        return b / a

    return (_idx(vals[0], vals[1]), _idx(vals[1], vals[2]))


def calc_growth_rate(vals: Sequence[float]) -> Tuple[Optional[float], Optional[float]]:
    i10, i21 = calc_index(vals)
    r10 = (i10 - 1.0) if i10 is not None else None
    r21 = (i21 - 1.0) if i21 is not None else None
    return (r10, r21)


def calc_ritm_mediu(vals: Sequence[float]) -> Optional[float]:
    """Indicele mediu de dinamică (medie geometrică pe 2 intervale); None dacă nu este definit."""
    if not all_ok(*vals) or abs(vals[0]) < 1e-9:
        return None
    ratio = vals[2] / vals[0]
    if ratio <= 0:
        return None
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


ND_TEXT = "n.d."


def fmt_money(v: Optional[NumberLike], decimals: int = 2, dash_zero: bool = False) -> str:
    if v is None:
        return "—"
    if is_nd(v):
        return ND_TEXT
    val = float(v)
    if abs(val) < 0.5 * 10 ** (-decimals):
        val = 0.0
    sign = "-" if val < 0 else ""
    abs_val = abs(val)
    formatted = f"{abs_val:,.{decimals}f}"
    formatted = formatted.replace(",", "X").replace(".", ",").replace("X", ".")
    return sign + formatted


def fmt_pct(v: Optional[NumberLike], decimals: int = 2, plus_sign: bool = False) -> str:
    if v is None:
        return "—"
    if is_nd(v):
        return ND_TEXT
    val = float(v)
    if abs(val) < 0.5 * 10 ** (-decimals):
        val = 0.0
    sign = "+" if (plus_sign and val > 0) else ""
    s = f"{val:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}{s}%"


def fmt_ratio_pct(v: Optional[NumberLike], decimals: int = 2, plus_sign: bool = False) -> str:
    if v is None:
        return "—"
    if is_nd(v):
        return ND_TEXT
    return fmt_pct(float(v) * 100.0, decimals=decimals, plus_sign=plus_sign)


def fmt_pp(v: Optional[NumberLike], decimals: int = 2) -> str:
    """Diferență de rate exprimată în puncte procentuale (v este o fracție: 0,05 = 5 pp)."""
    if v is None:
        return "—"
    if is_nd(v):
        return ND_TEXT
    val = float(v) * 100.0
    if abs(val) < 0.5 * 10 ** (-decimals):
        val = 0.0
    sign = "+" if val > 0 else ""
    s = f"{val:,.{decimals}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{sign}{s} pp"


def fmt_num(v: Optional[NumberLike], decimals: int = 4) -> str:
    if v is None:
        return "—"
    if is_nd(v):
        return ND_TEXT
    val = float(v)
    if abs(val) < 0.5 * 10 ** (-decimals):
        val = 0.0
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
    if s.startswith("EFECT DE MĂCIUCĂ"):
        return (COLOR_BAD_BG, COLOR_BAD_TEXT)
    if (
        s.startswith("Semnal favorabil")
        or s.startswith("EFECT DE LEVIER")
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


# ==============================================================================
# MOTOR DE DIVERSIFICARE NARATIVĂ ȘI DIAGNOSTIC CONDIȚIONAT PE VALORI
# ==============================================================================
def company_seed(bal: BalantaData, salt: str = "") -> int:
    """
    Calculează o amprentă numerică deterministă bazată pe datele de identificare
    și sumele din balanța de verificare, astfel încât două balanțe diferite
    să genereze formulări și structuri sintactice distincte.
    """
    t = bal.total_row
    sig = (
        f"{bal.cif}|{bal.denumire}|{bal.perioada}|"
        f"{t.sid if t else 0:.2f}|{t.rld if t else 0:.2f}|"
        f"{t.tsd if t else 0:.2f}|{t.sfd if t else 0:.2f}|{salt}"
    )
    digest = hashlib.sha256(sig.encode("utf-8")).hexdigest()
    return int(digest[:12], 16)


def pick_variant(bal: BalantaData, salt: str, variants: Sequence[str]) -> str:
    """Selectează determinist o formulare din lista de variante pe baza amprentei balanței."""
    if not variants:
        return ""
    idx = company_seed(bal, salt) % len(variants)
    return variants[idx]


def describe_trajectory(
    vals: Sequence[float],
    bal: BalantaData,
    salt: str,
    label: str,
    is_ratio: bool = False,
) -> str:
    """
    Generează o descriere analitică diversificată a evoluției unui indicator pe
    cele trei perioade de raportare (0, 1 și 2), adaptată regimului de variație.
    """
    v0, v1, v2 = float(vals[0]), float(vals[1]), float(vals[2])
    f0 = fmt_ratio_pct(v0) if is_ratio else f"{fmt_money(v0)} lei"
    f1 = fmt_ratio_pct(v1) if is_ratio else f"{fmt_money(v1)} lei"
    f2 = fmt_ratio_pct(v2) if is_ratio else f"{fmt_money(v2)} lei"

    if is_nd(v1) or is_nd(v2):
        return f"{label} nu poate fi determinat(ă) pentru toate perioadele (valori n.d.)"
    if is_nd(v0):
        if abs(v2 - v1) <= max(abs(v1) * 0.015, 1e-4):
            return f"{label} rămâne practic constant(ă) între perioada de raportare 1 ({f1}) și perioada de raportare 2 ({f2})"
        verb = "crește" if v2 > v1 else "scade"
        return f"{label} {verb} de la {f1} în perioada de raportare 1 la {f2} în perioada de raportare 2 (valoarea pentru perioada 0 nu este disponibilă)"

    rel_total = safe_div(v2 - v0, abs(v0)) if abs(v0) > 1e-6 else 0.0
    rel_str = fmt_ratio_pct(rel_total, plus_sign=True) if abs(v0) > 1e-6 else "n/a"

    if abs(v2 - v0) <= max(abs(v0) * 0.015, 1e-4) and abs(v1 - v0) <= max(abs(v0) * 0.015, 1e-4):
        return pick_variant(
            bal,
            f"traj_stab_{salt}",
            [
                f"{label} prezintă o evoluție stabilă pe întregul orizont analizat ({f0} în perioada de raportare 0, {f1} în perioada de raportare 1 și {f2} în perioada de raportare 2)",
                f"{label} se menține la un palier cvasi-constant în cele trei perioade de raportare ({f0} -> {f1} -> {f2})",
                f"nivelul înregistrat de {label} rămâne echilibrat și lipsit de fluctuații semnificative ({f0} în perioada 0 și {f2} în perioada 2)",
                f"dinamica indicatorului {label} relevă conservarea volumului patrimonial în jurul valorii de {f2} ({f0} în perioada de raportare 0)",
            ],
        )

    if v0 < v1 < v2:
        if rel_total > 0.20:
            return pick_variant(
                bal,
                f"traj_up_strong_{salt}",
                [
                    f"{label} consemnează o expansiune susținută de {rel_str}, urcând progresiv de la {f0} în perioada de raportare 0 la {f1} în perioada de raportare 1 și {f2} în perioada de raportare 2",
                    f"evoluția indicatorului {label} este puternic ascendentă ({rel_str} pe ansamblul intervalului), avansând etapizat de la {f0} (perioada 0) la {f1} (perioada 1) și {f2} (perioada 2)",
                    f"se remarcă o consolidare accelerată pentru {label}, care își majorează nivelul cu {rel_str}, de la {f0} în perioada de raportare 0 până la {f2} în perioada de raportare 2 (via {f1} în perioada 1)",
                    f"traiectoria urmată de {label} indică o creștere continuă și viguroasă ({rel_str}), de la {f0} la începutul intervalului la {f1} în perioada de raportare 1 și {f2} la finele perioadei de raportare 2",
                ],
            )
        return pick_variant(
            bal,
            f"traj_up_mod_{salt}",
            [
                f"{label} urmează o tendință pozitivă graduală ({rel_str}), crescând de la {f0} în perioada de raportare 0 la {f1} în perioada de raportare 1 și {f2} în perioada de raportare 2",
                f"pe parcursul celor trei perioade de raportare, {label} înregistrează o apreciere moderată și constantă ({rel_str}), de la {f0} la {f1} și ulterior la {f2}",
                f"indicatorul {label} evoluează favorabil în ambele subperioade ({f0} -> {f1} -> {f2}, reprezentând o variație cumulată de {rel_str})",
                f"se observă o creștere organică a nivelului pentru {label}, care avansează de la {f0} în perioada de raportare 0 la {f2} în perioada de raportare 2 ({rel_str})",
            ],
        )

    if v1 < v0 and v2 > v1:
        if v2 >= v0:
            return pick_variant(
                bal,
                f"traj_v_full_{salt}",
                [
                    f"{label} traversează o ajustare temporară în perioada de raportare 1 ({f1} față de {f0} în perioada 0), urmată de o redresare completă în perioada de raportare 2 până la {f2} ({rel_str} față de perioada 0)",
                    f"după o contracție intermediară de la {f0} la {f1} în perioada de raportare 1, {label} revine pe creștere în perioada de raportare 2 și atinge {f2} ({rel_str} pe ansamblul intervalului)",
                    f"profilul dinamic al indicatorului {label} este unul în «V»: scădere tranzitorie la {f1} în perioada 1, urmată de o expansiune la {f2} în perioada 2, depășind nivelul inițial de {f0}",
                    f"deși în perioada de raportare 1 {label} s-a diminuat la {f1} (de la {f0}), revirimentul din perioada de raportare 2 ridică indicatorul la {f2} ({rel_str})",
                ],
            )
        return pick_variant(
            bal,
            f"traj_v_part_{salt}",
            [
                f"{label} scade de la {f0} în perioada de raportare 0 la {f1} în perioada de raportare 1, recuperând apoi parțial în perioada de raportare 2 până la nivelul de {f2} ({rel_str} față de perioada 0)",
                f"după reculul consemnat în perioada de raportare 1 ({f1} față de {f0}), {label} se redresează în perioada de raportare 2 la {f2} (variație cumulată de {rel_str})",
                f"evoluția pentru {label} marchează o diminuare în prima parte a intervalului ({f0} -> {f1}), atenuată de creșterea din perioada de raportare 2 până la {f2}",
                f"înregistrând un minim intermediar de {f1} în perioada de raportare 1, {label} își îmbunătățește poziția în perioada de raportare 2 la {f2} (față de {f0} în perioada 0)",
            ],
        )

    if v1 > v0 and v2 < v1:
        if v2 >= v0:
            return pick_variant(
                bal,
                f"traj_peak_pos_{salt}",
                [
                    f"{label} atinge un maxim în perioada de raportare 1 ({f1}, în urcare de la {f0}), după care se temperează la {f2} în perioada de raportare 2, menținându-se însă cu {rel_str} peste nivelul din perioada 0",
                    f"după vârful consemnat în perioada de raportare 1 ({f1}), {label} cunoaște o ajustare la {f2} în perioada de raportare 2, păstrând un avans net de {rel_str} față de perioada de raportare 0 ({f0})",
                    f"traiectoria indicatorului {label} evidențiază o expansiune pronunțată în perioada 1 ({f1}) urmată de o stabilizare la {f2} în perioada 2 (peste valoarea inițială de {f0})",
                    f"cu o creștere inițială de la {f0} la {f1} și o repliere marginală la {f2} în perioada de raportare 2, {label} încheie intervalul analizat pe un plus de {rel_str}",
                ],
            )
        return pick_variant(
            bal,
            f"traj_peak_neg_{salt}",
            [
                f"{label} crește de la {f0} în perioada de raportare 0 la {f1} în perioada de raportare 1, dar se contractă ulterior la {f2} în perioada de raportare 2 ({rel_str} față de perioada 0)",
                f"după o apreciere temporară în perioada de raportare 1 ({f1} față de {f0}), {label} coboară în perioada de raportare 2 la {f2} (variație totală {rel_str})",
                f"dinamica pentru {label} indică o inversare de tendință în partea a doua a intervalului: avans de la {f0} la {f1}, urmat de o reducere până la {f2} în perioada de raportare 2",
                f"atingând un nivel maxim de {f1} în perioada de raportare 1, {label} se ajustează în perioada de raportare 2 la {f2} ({rel_str} raportat la perioada 0)",
            ],
        )

    return pick_variant(
        bal,
        f"traj_down_{salt}",
        [
            f"{label} urmează o traiectorie descendentă ({rel_str}), diminuându-se de la {f0} în perioada de raportare 0 la {f1} în perioada de raportare 1 și {f2} în perioada de raportare 2",
            f"pe ansamblul celor trei perioade de raportare, {label} înregistrează o reducere progresivă de {fmt_ratio_pct(abs(rel_total)) if abs(v0) > 1e-6 else 'n/a'}, trecând de la {f0} (perioada 0) la {f1} (perioada 1) și {f2} (perioada 2)",
            f"se constată o contracție succesivă a indicatorului {label}, al cărui nivel coboară de la {f0} în perioada de raportare 0 până la {f2} în perioada de raportare 2 ({rel_str})",
            f"evoluția consemnată de {label} este una de restrângere treptată ({f0} -> {f1} -> {f2}, echivalentul unei modificări relative de {rel_str})",
        ],
    )


def describe_flow(vals: Sequence[float], bal: BalantaData, label: str) -> str:
    """
    Descrie un indicator de flux (rulaj cumulat de la începutul anului). Valorile din P1 și P2
    sunt cumulate pe intervale de lungimi diferite, deci nu se compară direct: se prezintă
    rulajul lunii curente (P2 - P1) și media lunară a celor două intervale.
    """
    v1, v2 = float(vals[1]), float(vals[2])
    m1, m2 = bal.period.months[1], bal.period.months[2]
    if is_nd(v1) or is_nd(v2):
        return f"{label} nu este disponibil(ă)"
    luna = v2 - v1
    luna_txt = bal.period.month_name.lower() or "luna curentă"
    if m1 <= 0:
        return f"{label} însumează {fmt_money(v2)} lei în {luna_txt} {bal.period.year}"
    avg1 = v1 / m1
    if abs(avg1) < 1e-6:
        cmp_txt = ""
    else:
        rel = (luna - avg1) / abs(avg1)
        if abs(rel) < 0.02:
            cmp_txt = f", în linie cu media lunară a perioadei precedente ({fmt_money(avg1)} lei/lună)"
        elif rel > 0:
            cmp_txt = f", peste media lunară a primelor {m1} luni ({fmt_money(avg1)} lei/lună, {fmt_ratio_pct(rel, plus_sign=True)})"
        else:
            cmp_txt = f", sub media lunară a primelor {m1} luni ({fmt_money(avg1)} lei/lună, {fmt_ratio_pct(rel, plus_sign=True)})"
    return (
        f"{label} cumulează {fmt_money(v1)} lei în primele {m1} luni și {fmt_money(v2)} lei în primele {m2} luni ale anului "
        f"{bal.period.year}, din care {fmt_money(luna)} lei în {luna_txt}{cmp_txt}"
    )


def extract_word_report_blocks(docx_path: Union[str, Path]) -> Dict[str, Any]:
    """
    Extrage integral conținutul structurat al documentului Word (.docx) generat
    (textul introductiv, cele 6 părți tematice, cele 21 de capitole, toate tabelele
    cu marcarea rândurilor de total și a celulelor semantice, paragrafele analitice
    și concluziile executive), pentru afișare identică în interfața web.
    """
    doc = Document(str(docx_path))
    intro_text = ""
    blocks: List[Dict[str, Any]] = []
    chapters_nav: List[Dict[str, str]] = []
    started_parts = False

    for child in doc.element.body:
        tag = child.tag.split("}")[-1]
        if tag == "tbl":
            tbl = Table(child, doc)
            if len(tbl.rows) == 1 and len(tbl.columns) == 1:
                cell_p = tbl.rows[0].cells[0].paragraphs[0]
                runs = cell_p.runs
                full_t = cell_p.text.strip()
                if "PARTEA " in full_t:
                    started_parts = True
                    kicker = runs[0].text.strip().rstrip("—").strip() if len(runs) > 0 else ""
                    title = runs[1].text.strip() if len(runs) > 1 else full_t
                    subtitle = runs[2].text.strip() if len(runs) > 2 else ""
                    blocks.append(
                        {
                            "type": "part_banner",
                            "kicker": kicker,
                            "title": title,
                            "subtitle": subtitle,
                        }
                    )
            else:
                if not started_parts:
                    continue
                headers = [c.text.strip() for c in tbl.rows[0].cells]
                rows_data = []
                for r in tbl.rows[1:]:
                    cells_info = []
                    is_total_row = False
                    for c_idx, c in enumerate(r.cells):
                        val_str = c.text.strip()
                        shd = c._tc.xpath(".//w:shd/@w:fill")
                        fill_hex = shd[0].upper() if shd else "FFFFFF"
                        if c_idx == 0 and fill_hex == COLOR_BG_TOTAL:
                            is_total_row = True
                        sem = _detect_semantic_style(val_str)
                        cells_info.append(
                            {
                                "text": val_str,
                                "bg": f"#{sem[0]}" if sem else None,
                                "fg": f"#{sem[1]}" if sem else None,
                            }
                        )
                    rows_data.append({"is_total": is_total_row, "cells": cells_info})
                blocks.append(
                    {
                        "type": "table",
                        "headers": headers,
                        "rows": rows_data,
                    }
                )
        elif tag == "p":
            p = Paragraph(child, doc)
            txt = p.text.strip()
            if not txt:
                continue
            if not started_parts:
                if not txt.startswith("0.1.") and not txt.startswith("0.2.") and not intro_text:
                    intro_text = txt
                continue

            has_bottom_bdr = bool(p._p.xpath(".//w:pBdr/w:bottom"))
            if has_bottom_bdr or any(txt.startswith(f"{i}. ") for i in range(1, 22)):
                cap_num = txt.split(".")[0].strip()
                cap_id = f"capitol-{cap_num}"
                chapters_nav.append({"id": cap_id, "num": cap_num, "title": txt})
                blocks.append(
                    {
                        "type": "heading",
                        "id": cap_id,
                        "text": txt,
                    }
                )
            elif txt.startswith("▪ "):
                parts = txt[2:].split(": ", 1)
                lead = parts[0].strip()
                body = parts[1].strip() if len(parts) > 1 else ""
                blocks.append(
                    {
                        "type": "bullet",
                        "lead": lead,
                        "body": body,
                    }
                )
            else:
                blocks.append(
                    {
                        "type": "paragraph",
                        "text": txt,
                    }
                )

    return {
        "intro_text": intro_text,
        "chapters_nav": chapters_nav,
        "blocks": blocks,
    }


