#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 9: Analiza Solvabilității Generale.
Calculează solvabilitatea generală (Activ total / Datorii totale, prag bancar > 166%)
pe cele trei perioade de raportare (0, 1 și 2).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Tuple

from docx import Document

from core_balanta import (
    BalantaData,
    add_heading,
    add_paragraph,
    calc_delta,
    calc_index,
    describe_trajectory,
    fmt_money,
    fmt_num,
    fmt_ratio_pct,
    make_table,
    pick_variant,
    safe_div,
)
from sheet_01_bilant import BilantResult
from sheet_03_bilant_financiar import BilantFinanciarResult
from sheet_05_structura_financiara import StructuraFinanciaraResult


@dataclass
class SolvabilitateResult:
    solv_generala: Tuple[float, float, float]
    solv_efectiva: Tuple[float, float, float]
    incadrare_norma: Tuple[str, str, str]


def compute_sheet_09_solvabilitatea(
    bf: BilantFinanciarResult, sf: StructuraFinanciaraResult, bil: BilantResult
) -> SolvabilitateResult:
    ta = bf.total_activ
    dt = sf.dtot
    sg = tuple(safe_div(ta[p], dt[p]) for p in range(3))

    dt_ef = tuple(round(bil.rd(11)[p] + bil.rd(14)[p], 2) for p in range(3))
    se = tuple(safe_div(ta[p], dt_ef[p]) for p in range(3))

    norma = []
    for p, lbl in enumerate(("perioada de raportare 0", "perioada de raportare 1", "perioada de raportare 2")):
        if sg[p] >= 1.66:
            norma.append(f"În {lbl}, solvabilitatea generală ({fmt_ratio_pct(sg[p])}) depășește pragul minim de 166%.")
        else:
            norma.append(f"În {lbl}, solvabilitatea generală ({fmt_ratio_pct(sg[p])}) este sub pragul de 166% (însă peste 100%).")

    return SolvabilitateResult(
        solv_generala=sg,
        solv_efectiva=se,
        incadrare_norma=(norma[0], norma[1], norma[2]),
    )


def render_sheet_09_solvabilitatea(
    doc: Document,
    bal: BalantaData,
    bf: BilantFinanciarResult,
    sf: StructuraFinanciaraResult,
    bil: BilantResult,
    res: SolvabilitateResult,
) -> None:
    """Generează secțiunea Analizei Solvabilității în documentul Word."""
    add_heading(doc, "9. Analiza Solvabilității Generale", level=1)

    ta = bf.total_activ
    dt = sf.dtot
    sg = res.solv_generala

    d_ta, i_ta = calc_delta(ta), calc_index(ta)
    d_dt, i_dt = calc_delta(dt), calc_index(dt)
    d_sg, i_sg = calc_delta(sg), calc_index(sg)

    headers = [
        "Element / Indicator de Solvabilitate",
        "U.M. / Prag",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
        "Δ (P1 - P0)",
        "Δ (P2 - P1)",
        "I (P2 / P1)",
    ]
    rows = [
        ["1. Activ total patrimonial", "lei", fmt_money(ta[0]), fmt_money(ta[1]), fmt_money(ta[2]), fmt_money(d_ta[0]), fmt_money(d_ta[1]), fmt_ratio_pct(i_ta[1])],
        ["2. Datorii totale (DTL + DTS)", "lei", fmt_money(dt[0]), fmt_money(dt[1]), fmt_money(dt[2]), fmt_money(d_dt[0]), fmt_money(d_dt[1]), fmt_ratio_pct(i_dt[1])],
        ["3. Solvabilitatea generală (coeficient)", "> 1,66", fmt_num(sg[0], 4), fmt_num(sg[1], 4), fmt_num(sg[2], 4), fmt_num(d_sg[0], 4), fmt_num(d_sg[1], 4), fmt_ratio_pct(i_sg[1])],
        ["4. Solvabilitatea generală (în procente)", "> 166%", fmt_ratio_pct(sg[0]), fmt_ratio_pct(sg[1]), fmt_ratio_pct(sg[2]), fmt_ratio_pct(d_sg[0], plus_sign=True), fmt_ratio_pct(d_sg[1], plus_sign=True), fmt_ratio_pct(i_sg[1])],
    ]

    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[5.4, 1.8, 1.9, 1.9, 1.9, 1.7, 1.7, 1.5],
        total_rows=[2, 3],
        font_size=7.8,
    )

    traj_sg = describe_trajectory(sg, bal, "s09_sg", "rata solvabilității generale (Activ total / Datorii totale)", is_ratio=True)
    if sg[2] >= 2.0:
        solv_msg = pick_variant(
            bal,
            "s09_high",
            [
                f"La nivelul perioadei de raportare 2, activul total ({fmt_money(ta[2])} lei) depășește de {fmt_num(sg[2], 2)} ori datoriile totale ({fmt_money(dt[2])} lei), depășind cu o rezervă substanțială pragul bancar de 166% și garantând securitatea creditorilor printr-un activ net pozitiv de {fmt_money(sf.cp[2])} lei.",
                f"Coeficientul solvabilității generale din perioada de raportare 2 ({fmt_num(sg[2], 4)}, respectiv {fmt_ratio_pct(sg[2])}) confirmă o autonomie patrimonială solidă, datoriile totale ({fmt_money(dt[2])} lei) fiind acoperite de peste două ori de valoarea activelor bilanțiere ({fmt_money(ta[2])} lei).",
            ],
        )
    elif sg[2] >= 1.66:
        solv_msg = (
            f"În perioada de raportare 2, solvabilitatea generală ({fmt_ratio_pct(sg[2])}) depășește plafonul de siguranță de 166%, activul total ({fmt_money(ta[2])} lei) asigurând rambursarea integrală a datoriilor totale ({fmt_money(dt[2])} lei) și menținerea unui capital propriu de {fmt_money(sf.cp[2])} lei."
        )
    else:
        solv_msg = (
            f"În perioada de raportare 2, solvabilitatea generală se situează la {fmt_ratio_pct(sg[2])} (coeficient {fmt_num(sg[2], 4)}), activul total de {fmt_money(ta[2])} lei raportându-se la datorii totale de {fmt_money(dt[2])} lei."
        )

    add_paragraph(doc, f"Diagnostic de solvabilitate: Pe ansamblul intervalului analizat, {traj_sg}. {solv_msg}")
