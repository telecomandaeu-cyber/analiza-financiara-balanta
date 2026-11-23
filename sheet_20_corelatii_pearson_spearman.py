#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 20: Analiza Corelațiilor Statistice (Pearson și Spearman).

Corelațiile se calculează doar pentru indicatorii de stoc (bilanțieri), care au valori în
toate cele trei momente (P0, P1, P2). Indicatorii de flux (cifra de afaceri, rezultatele,
scorurile de risc) nu au valori în P0 și sunt cumulați pe intervale diferite, deci nu pot
forma serii comparabile.

Cu doar n = 3 observații, coeficienții au valoare DESCRIPTIVĂ: testul t pentru r are un
singur grad de libertate, iar pragul de semnificație de 5% cere |r| ≥ 0,997. Rangurile
Spearman folosesc media rangurilor la egalitate (r_s = Pearson aplicat pe ranguri).
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import List, Optional, Sequence, Tuple

from docx import Document

from core_balanta import (
    BalantaData,
    add_heading,
    add_paragraph,
    all_ok,
    fmt_num,
    make_table,
)
from sheet_01_bilant import BilantResult
from sheet_02_cpp import CPPResult
from sheet_03_bilant_financiar import BilantFinanciarResult
from sheet_05_structura_financiara import StructuraFinanciaraResult
from sheet_07_echilibrul_financiar import EchilibruFinanciarResult
from sheet_08_lichiditatea import LichiditateResult
from sheet_09_solvabilitatea import SolvabilitateResult
from sheet_13_sig import SIGResult
from sheet_17_riscul_faliment import RiscFalimentResult

Triple = Tuple[float, float, float]


@dataclass
class CorrelationPairResult:
    idx: int
    name_x: str
    name_y: str
    pair_title: str
    x_vals: Triple
    y_vals: Triple
    pearson_r: Optional[float]
    calificativ: str
    ranks_x: Triple
    ranks_y: Triple
    spearman_rs: Optional[float]
    p_value: Optional[float]


@dataclass
class CorelatiiResult:
    pairs: List[CorrelationPairResult] = field(default_factory=list)


def _avg_ranks(vals: Sequence[float]) -> Triple:
    """Ranguri descrescătoare (1 = cea mai mare valoare), cu media rangurilor la egalitate."""
    ranks = []
    for v in vals:
        greater = sum(1 for o in vals if o > v + 1e-9)
        equal = sum(1 for o in vals if abs(o - v) <= 1e-9)
        ranks.append(greater + (equal + 1) / 2.0)
    return (ranks[0], ranks[1], ranks[2])


def _pearson(x: Sequence[float], y: Sequence[float]) -> Optional[float]:
    if not all_ok(*x, *y):
        return None
    n = len(x)
    mx, my = sum(x) / n, sum(y) / n
    num = sum((x[i] - mx) * (y[i] - my) for i in range(n))
    dx = math.sqrt(sum((v - mx) ** 2 for v in x))
    dy = math.sqrt(sum((v - my) ** 2 for v in y))
    if dx < 1e-9 * max(1.0, abs(mx)) or dy < 1e-9 * max(1.0, abs(my)):
        return None  # serie constantă: corelația nu este definită
    return max(-1.0, min(1.0, num / (dx * dy)))


def _p_value_n3(r: Optional[float]) -> Optional[float]:
    """p bilateral pentru H0: ρ = 0 cu n = 3 (t cu 1 grad de libertate = distribuția Cauchy)."""
    if r is None:
        return None
    if abs(r) >= 1.0 - 1e-12:
        return 0.0
    t = abs(r) / math.sqrt(1.0 - r * r)
    return 1.0 - (2.0 / math.pi) * math.atan(t)


def _calificativ(r: Optional[float]) -> str:
    if r is None:
        return "n.d. (serie constantă)"
    ar = abs(r)
    sens = "directă" if r >= 0 else "inversă"
    if ar >= 0.75:
        return f"Corelație puternică {sens}"
    if ar >= 0.50:
        return f"Corelație moderată {sens}"
    if ar >= 0.25:
        return f"Corelație slabă {sens}"
    return "Corelație neglijabilă"


def compute_sheet_20_corelatii(
    bil: BilantResult,
    cpp: CPPResult,
    bf: BilantFinanciarResult,
    sf: StructuraFinanciaraResult,
    ef: EchilibruFinanciarResult,
    liq: LichiditateResult,
    solv: SolvabilitateResult,
    sig: SIGResult,
    risc: RiscFalimentResult,
) -> CorelatiiResult:
    """Corelații Pearson și Spearman pe seriile bilanțiere complete (P0, P1, P2)."""
    res = CorelatiiResult()
    pair_defs = [
        ("Lichiditatea curentă", "Solvabilitatea generală", liq.l_curenta, solv.solv_generala),
        ("Lichiditatea rapidă", "Lichiditatea curentă", liq.l_rapida, liq.l_curenta),
        ("Lichiditatea imediată", "Lichiditatea rapidă", liq.l_imediata, liq.l_rapida),
        ("Fondul de rulment", "Necesarul de fond de rulment", ef.fr_sus, ef.nfr),
        ("Fondul de rulment", "Trezoreria netă", ef.fr_sus, ef.tn),
        ("Activele imobilizate", "Capitalurile permanente", bf.active_imobilizate, sf.cperm),
        ("Activele circulante", "Datoriile pe termen scurt", bf.active_circulante, sf.dts),
        ("Capitalurile proprii", "Activele imobilizate", sf.cp, bf.active_imobilizate),
        ("Capitalurile proprii", "Datoriile pe termen lung", sf.cp, sf.dtl),
        ("Capitalurile proprii", "Datoriile pe termen scurt", sf.cp, sf.dts),
        ("Datoriile pe termen lung", "Datoriile pe termen scurt", sf.dtl, sf.dts),
        ("Stocurile", "Datoriile pe termen scurt", bf.stocuri, sf.dts),
        ("Creanțele", "Disponibilitățile bănești", bf.creante, bf.disponibilitati),
    ]
    for i, (nx, ny, xv, yv) in enumerate(pair_defs, start=1):
        r = _pearson(xv, yv)
        rkx, rky = _avg_ranks(xv), _avg_ranks(yv)
        rs = _pearson(rkx, rky)
        res.pairs.append(
            CorrelationPairResult(
                idx=i,
                name_x=nx,
                name_y=ny,
                pair_title=f"{nx} și {ny}",
                x_vals=xv,
                y_vals=yv,
                pearson_r=r,
                calificativ=_calificativ(r),
                ranks_x=rkx,
                ranks_y=rky,
                spearman_rs=rs,
                p_value=_p_value_n3(r),
            )
        )
    return res


def _fmt_rank(v: float) -> str:
    return str(int(v)) if abs(v - round(v)) < 1e-9 else fmt_num(v, 1)


def render_sheet_20_corelatii(doc: Document, bal: BalantaData, res: CorelatiiResult) -> None:
    """Generează secțiunea Corelațiilor Statistice."""
    add_heading(doc, "20. Analiza Corelațiilor Statistice (Pearson și Spearman) între Indicatorii Bilanțieri", level=1)

    headers = ["Nr.", "Perechea de indicatori (X și Y)", "Ranguri X", "Ranguri Y", "Spearman r_s", "Pearson r", "p (n = 3)", "Intensitate"]
    rows = []
    for p in res.pairs:
        rows.append([
            f"{p.idx:02d}",
            p.pair_title,
            "(" + ", ".join(_fmt_rank(v) for v in p.ranks_x) + ")",
            "(" + ", ".join(_fmt_rank(v) for v in p.ranks_y) + ")",
            fmt_num(p.spearman_rs, 2),
            fmt_num(p.pearson_r, 4),
            fmt_num(p.p_value, 3),
            p.calificativ,
        ])
    make_table(doc, headers=headers, rows=rows, col_widths=[0.8, 6.0, 1.6, 1.6, 1.4, 1.5, 1.3, 2.6], font_size=7.6)

    valid = [p for p in res.pairs if p.pearson_r is not None]
    strong = [p for p in valid if abs(p.pearson_r) >= 0.75]
    signif = [p for p in valid if p.p_value is not None and p.p_value < 0.05]
    top = sorted(valid, key=lambda p: abs(p.pearson_r), reverse=True)[:2]
    top_txt = "; ".join(f"«{p.pair_title}» (r = {fmt_num(p.pearson_r, 4)})" for p in top) if top else "—"
    add_paragraph(
        doc,
        f"Sinteză statistică: din {len(res.pairs)} perechi de indicatori bilanțieri, {len(strong)} prezintă o corelație "
        f"puternică (|r| ≥ 0,75), cele mai intense fiind {top_txt}. Seria are însă doar trei observații (P0, P1, P2), astfel încât "
        f"{len(signif)} perechi ating pragul de semnificație statistică de 5% (|r| ≥ 0,997); coeficienții descriu sensul "
        f"evoluțiilor comune pe intervalul analizat și nu trebuie interpretați ca relații cauzale sau stabile. Indicatorii de flux "
        f"(cifra de afaceri, rezultatele, scorurile de risc) au fost excluși, deoarece nu au valori pentru perioada de raportare 0.",
    )
