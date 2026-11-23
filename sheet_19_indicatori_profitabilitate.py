#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 19: Analiza Indicatorilor de Profitabilitate și Rentabilitate.

Ratele de marjă (raportate la cifra de afaceri sau la cheltuieli) folosesc rulaje cumulate
din aceeași perioadă și nu necesită anualizare. Ratele care raportează un flux la un stoc
(ROA, ROE) se calculează pe rezultatul ANUALIZAT (× zile an / zile perioadă).
ROE nu se calculează când capitalurile proprii sunt negative sau nule.
Pentru P0 (solduri de deschidere, fără rulaje) ratele nu sunt disponibile.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Tuple

from docx import Document

from core_balanta import (
    ND,
    BalantaData,
    add_heading,
    add_paragraph,
    fmt_pct,
    fmt_pp,
    make_table,
    ratio_or_nd,
)
from sheet_02_cpp import CPPResult
from sheet_03_bilant_financiar import BilantFinanciarResult

Triple = Tuple[float, float, float]


@dataclass
class ProfitBlock:
    title: str
    num_name: str
    den_name: str
    rate_name: str
    formula_str: str
    num_vals: Triple
    den_vals: Triple
    rate_pct_vals: Triple


@dataclass
class IndicatoriProfitabilitateResult:
    marja_bruta: Triple
    marja_neta: Triple
    rent_resurse: Triple
    roa: Triple
    roe: Triple
    ros: Triple
    blocks: List[ProfitBlock] = field(default_factory=list)


def compute_sheet_19_indicatori_profitabilitate(
    bf: BilantFinanciarResult,
    cpp: CPPResult,
    annual_factors: Tuple[float, float, float] = (ND, 365 / 243, 365 / 273),
) -> IndicatoriProfitabilitateResult:
    k = annual_factors
    pb = cpp.rezultat_brut
    pn = cpp.rezultat_net
    re = cpp.rezultat_exploatare
    ca = cpp.cifra_afaceri
    ct = cpp.cheltuieli_totale
    ta = bf.total_activ
    cp = bf.capitaluri_proprii

    t = lambda fn: tuple(fn(p) for p in range(3))  # noqa: E731
    pct = lambda a, b: ratio_or_nd(a, b) * 100.0  # noqa: E731
    mb = t(lambda p: pct(pb[p], ca[p]))
    mn = t(lambda p: pct(pn[p], ca[p]))
    ros = t(lambda p: pct(re[p], ca[p]))
    rrc = t(lambda p: pct(pn[p], ct[p]))
    pn_an = t(lambda p: pn[p] * k[p])
    roa = t(lambda p: pct(pn_an[p], ta[p]))
    roe = t(lambda p: pct(pn_an[p], cp[p]) if cp[p] > 0 else ND)

    blocks = [
        ProfitBlock("Marja brută", "Profit brut", "Cifra de afaceri", "1. Marja brută din vânzări", "Rezultat brut / CA × 100", pb, ca, mb),
        ProfitBlock("Marja netă", "Profit net", "Cifra de afaceri", "2. Marja netă din vânzări", "Rezultat net / CA × 100", pn, ca, mn),
        ProfitBlock("Rentabilitatea resurselor consumate", "Profit net", "Cheltuieli totale", "3. Rentabilitatea resurselor consumate", "Rezultat net / Cheltuieli totale × 100", pn, ct, rrc),
        ProfitBlock("ROA", "Profit net anualizat", "Activ total", "4. Rentabilitatea activelor (ROA, anualizată)", "Rezultat net anualizat / Activ total × 100", pn_an, ta, roa),
        ProfitBlock("ROE", "Profit net anualizat", "Capital propriu", "5. Rentabilitatea financiară (ROE, anualizată)", "Rezultat net anualizat / Capital propriu × 100", pn_an, cp, roe),
        ProfitBlock("ROS", "Rezultat din exploatare", "Cifra de afaceri", "6. Rentabilitatea comercială operațională (ROS)", "Rezultat din exploatare / CA × 100", re, ca, ros),
    ]
    # blocks[0].den_vals = cifra de afaceri (folosit în tabloul de bord)
    return IndicatoriProfitabilitateResult(
        marja_bruta=mb, marja_neta=mn, rent_resurse=rrc, roa=roa, roe=roe, ros=ros, blocks=blocks,
    )


def render_sheet_19_indicatori_profitabilitate(doc: Document, bal: BalantaData, res: IndicatoriProfitabilitateResult) -> None:
    """Generează secțiunea Indicatorilor de Profitabilitate."""
    add_heading(doc, "19. Analiza Indicatorilor de Profitabilitate și Rentabilitate", level=1)

    headers = ["Indicator de profitabilitate", "Formulă de calcul", *bal.col_labels, "Δ (P2 - P1)"]
    rows = []
    for b in res.blocks:
        v = b.rate_pct_vals
        rows.append([b.rate_name, b.formula_str, fmt_pct(v[0]), fmt_pct(v[1]), fmt_pct(v[2]), fmt_pp((v[2] - v[1]) / 100.0)])

    make_table(doc, headers=headers, rows=rows, col_widths=[5.0, 4.4, 1.6, 1.9, 1.9, 1.8], total_rows=[3, 4, 5], font_size=7.8)

    roe_txt = (
        f"rentabilitatea financiară anualizată (ROE) este de {fmt_pct(res.roe[2])}"
        if res.roe[2] == res.roe[2]
        else "rentabilitatea financiară (ROE) nu se poate calcula, capitalurile proprii fiind negative"
    )
    add_paragraph(
        doc,
        f"Sinteză privind rentabilitatea: în perioada de raportare 2, marja netă din vânzări este de {fmt_pct(res.marja_neta[2])} "
        f"(față de {fmt_pct(res.marja_neta[1])} în perioada de raportare 1), marja operațională (ROS) de {fmt_pct(res.ros[2])}, iar la "
        f"fiecare 100 lei de cheltuieli totale societatea obține un rezultat net de {fmt_pct(res.rent_resurse[2]).rstrip('%')} lei. "
        f"Rentabilitatea anualizată a activelor (ROA) este de {fmt_pct(res.roa[2])}, iar {roe_txt}. Ratele de marjă se calculează pe "
        f"rulajele cumulate ale aceleiași perioade, iar ROA și ROE pe rezultatul anualizat, pentru a fi comparabile între perioade "
        f"de lungimi diferite și cu ratele anuale de referință.",
    )
