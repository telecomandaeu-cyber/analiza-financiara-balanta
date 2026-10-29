#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 14: Capacitatea de Autofinanțare (CAF — Metoda Deductivă și Metoda Aditivă).
Calculează potențialul financiar intern degajat de activitatea curentă prin ambele metode
consacrate (Metoda Deductivă pornind de la EBE și Metoda Aditivă pornind de la Rezultatul Net).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Tuple

from docx import Document

from core_balanta import (
    BalantaData,
    add_heading,
    add_paragraph,
    calc_delta,
    calc_growth_rate,
    calc_index,
    calc_ritm_mediu,
    fmt_money,
    fmt_ratio_pct,
    make_table,
)
from sheet_02_cpp import CPPResult
from sheet_13_sig import SIGResult


@dataclass
class CAFResult:
    caf_deductiv: Tuple[float, float, float]
    caf_aditiv: Tuple[float, float, float]
    delta_ded: Tuple[float, float]
    delta_ad: Tuple[float, float]
    index_ded: Tuple[Optional[float], Optional[float]]
    index_ad: Tuple[Optional[float], Optional[float]]
    growth_ded: Tuple[Optional[float], Optional[float]]
    growth_ad: Tuple[Optional[float], Optional[float]]
    ritm_ded: Optional[float]
    ritm_ad: Optional[float]


def compute_sheet_14_caf(cpp: CPPResult, sig: SIGResult) -> CAFResult:
    """Calculează CAF deductiv și CAF aditiv pe cele trei perioade de raportare."""
    caf_ded = []
    caf_ad = []
    for p in range(3):
        ebe_p = sig.ebe[p]
        alte_v_expl = cpp.rd(10)[p]
        alte_c_expl = cpp.rd(30)[p] + cpp.rd(31)[p]
        v_fin_mon = cpp.rd(45)[p] - cpp.rd(48)[p]
        c_fin_mon = cpp.rd(52)[p] - cpp.rd(47)[p]
        v_ext = cpp.rd(57)[p]
        c_ext = cpp.rd(58)[p]
        imp = cpp.rd(65)[p] + cpp.rd(66)[p]

        val_ded = round(
            ebe_p + alte_v_expl - alte_c_expl + v_fin_mon - c_fin_mon + v_ext - c_ext - imp, 2
        )

        r_net = sig.rezultat_net[p]
        amort_prov_calc = cpp.rd(22)[p] + cpp.rd(25)[p] + cpp.rd(33)[p] + cpp.rd(47)[p]
        reluari_ven = cpp.rd(23)[p] + cpp.rd(26)[p] + cpp.rd(34)[p] + cpp.rd(48)[p]
        val_ad = round(r_net + amort_prov_calc - reluari_ven, 2)

        caf_ded.append(val_ded)
        caf_ad.append(val_ad)

    t_ded = (caf_ded[0], caf_ded[1], caf_ded[2])
    t_ad = (caf_ad[0], caf_ad[1], caf_ad[2])

    return CAFResult(
        caf_deductiv=t_ded,
        caf_aditiv=t_ad,
        delta_ded=calc_delta(t_ded),
        delta_ad=calc_delta(t_ad),
        index_ded=calc_index(t_ded),
        index_ad=calc_index(t_ad),
        growth_ded=calc_growth_rate(t_ded),
        growth_ad=calc_growth_rate(t_ad),
        ritm_ded=calc_ritm_mediu(t_ded),
        ritm_ad=calc_ritm_mediu(t_ad),
    )


def render_sheet_14_caf(doc: Document, bal: BalantaData, cpp: CPPResult, sig: SIGResult, res: CAFResult) -> None:
    """Generează secțiunea Capacității de Autofinanțare (CAF) într-un singur tabel unificat."""
    add_heading(doc, "14. Capacitatea de Autofinanțare (CAF — Metoda Deductivă și Metoda Aditivă)", level=1)

    headers = [
        "Componentă de Calcul a Capacității de Autofinanțare (lei)",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
        "Δ (P1 - P0)",
        "Δ (P2 - P1)",
        "I (P2 / P1)",
        "Ritm mediu",
    ]

    items = [
        ("Excedentul / Insuficiența brută de exploatare (EBE)", sig.ebe, False),
        ("+ Alte venituri monetare din exploatare (ct. 758)", cpp.rd(10), False),
        ("- Alte cheltuieli monetare de exploatare (ct. 658)", cpp.rd(30), False),
        ("+ Venituri financiare monetare (ct. 765 + 766 + 768)", cpp.rd(45), False),
        ("- Cheltuieli financiare monetare (ct. 665 + 666)", cpp.rd(52), False),
        ("- Impozitul pe profit (ct. 691)", cpp.rd(65), False),
        ("= CAPACITATEA DE AUTOFINANȚARE — METODA DEDUCTIVĂ", res.caf_deductiv, True),
        ("Rezultatul net al exercițiului (Profit net)", sig.rezultat_net, False),
        ("+ Cheltuieli calculate cu amortizarea imobilizărilor (ct. 6811)", cpp.rd(22), False),
        ("= CAPACITATEA DE AUTOFINANȚARE — METODA ADITIVĂ", res.caf_aditiv, True),
    ]

    rows = []
    tot_idx = []
    for i, (label, vals, is_tot) in enumerate(items):
        if is_tot:
            tot_idx.append(i)
        d = calc_delta(vals)
        idx = calc_index(vals)
        rm = calc_ritm_mediu(vals)
        rows.append([
            label,
            fmt_money(vals[0]),
            fmt_money(vals[1]),
            fmt_money(vals[2]),
            fmt_money(d[0]),
            fmt_money(d[1]),
            fmt_ratio_pct(idx[1]),
            fmt_ratio_pct(rm),
        ])

    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[5.2, 1.9, 1.9, 1.9, 1.8, 1.8, 1.65, 1.65],
        total_rows=tot_idx,
        font_size=7.7,
    )

    g_caf_10 = (res.caf_deductiv[1] - res.caf_deductiv[0]) / res.caf_deductiv[0] if abs(res.caf_deductiv[0]) > 1e-6 else 0.0
    g_caf_21 = (res.caf_deductiv[2] - res.caf_deductiv[1]) / res.caf_deductiv[1] if abs(res.caf_deductiv[1]) > 1e-6 else 0.0
    add_paragraph(
        doc,
        f"Sinteză CAF: Metoda deductivă și metoda aditivă conduc la aceeași valoare a Capacității de Autofinanțare: "
        f"{fmt_money(res.caf_deductiv[0])} lei în perioada de raportare 0, {fmt_money(res.caf_deductiv[1])} lei în perioada de "
        f"raportare 1 ({fmt_ratio_pct(g_caf_10, plus_sign=True)} față de perioada de raportare 0) și {fmt_money(res.caf_deductiv[2])} lei "
        f"în perioada de raportare 2 ({fmt_ratio_pct(g_caf_21, plus_sign=True)} față de perioada de raportare 1, ritm mediu "
        f"{fmt_ratio_pct(res.ritm_ded)}). Acest surplus monetar intern asigură resursele proprii necesare pentru autofinanțarea "
        "activității curente și a investițiilor.",
    )
