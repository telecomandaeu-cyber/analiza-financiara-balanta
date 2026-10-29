#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 16: Analiza Efectului de Levier Financiar (Efectul de Îndatorare).
Compară rata rentabilității (Re) cu rata medie a dobânzii creditului bancar (Rd)
și cuantifică efectul de levier financiar pe cele trei perioade de raportare (0, 1 și 2).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Tuple

from docx import Document

from core_balanta import (
    BalantaData,
    add_heading,
    add_paragraph,
    fmt_money,
    fmt_ratio_pct,
    make_table,
    safe_div,
)
from sheet_02_cpp import CPPResult
from sheet_03_bilant_financiar import BilantFinanciarResult
from sheet_05_structura_financiara import StructuraFinanciaraResult


@dataclass
class EfectIndatorareResult:
    cp: Tuple[float, float, float]
    re_rate: Tuple[float, float, float]
    credit: Tuple[float, float, float]
    dobanzi: Tuple[float, float, float]
    rd_rate: Tuple[float, float, float]
    efect_levier: Tuple[float, float, float]
    verdict: Tuple[str, str, str]


def compute_sheet_16_efectul_indatorare(
    bf: BilantFinanciarResult, sf: StructuraFinanciaraResult, cpp: CPPResult
) -> EfectIndatorareResult:
    cp = sf.cp
    credit = sf.dtl
    dob = cpp.rd(49)
    rn = cpp.rezultat_net

    re_r = tuple(safe_div(rn[p], cp[p]) for p in range(3))
    rd_r = tuple(safe_div(dob[p], credit[p]) for p in range(3))
    lev = tuple((re_r[p] - rd_r[p]) * safe_div(credit[p], cp[p]) for p in range(3))

    verdicts = []
    for p, per_lbl in enumerate(("perioada de raportare 0", "perioada de raportare 1", "perioada de raportare 2")):
        if re_r[p] >= rd_r[p]:
            verdicts.append(f"În {per_lbl}, îndatorarea generează EFECT DE LEVIER pozitiv asupra întreprinderii (Re > Rd)")
        else:
            verdicts.append(f"În {per_lbl}, îndatorarea generează EFECT DE MĂCIUCĂ negativ asupra întreprinderii (Re < Rd)")

    return EfectIndatorareResult(
        cp=cp,
        re_rate=re_r,
        credit=credit,
        dobanzi=dob,
        rd_rate=rd_r,
        efect_levier=lev,
        verdict=(verdicts[0], verdicts[1], verdicts[2]),
    )


def render_sheet_16_efectul_indatorare(
    doc: Document, bal: BalantaData, res: EfectIndatorareResult
) -> None:
    """Generează secțiunea Efectului de Levier Financiar în documentul Word."""
    add_heading(doc, "16. Analiza Efectului de Levier Financiar (Efectul de Îndatorare)", level=1)

    headers = [
        "Nr.",
        "Elemente Efectul de Levier Financiar",
        "U.M.",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
    ]
    rows = [
        ["1", "Capitaluri proprii (CP)", "lei", fmt_money(res.cp[0]), fmt_money(res.cp[1]), fmt_money(res.cp[2])],
        ["2", "Re (Rata rentabilității capitalurilor proprii)", "%", fmt_ratio_pct(res.re_rate[0]), fmt_ratio_pct(res.re_rate[1]), fmt_ratio_pct(res.re_rate[2])],
        ["3", "Credit bancar pe termen lung în sold (ct. 162)", "lei", fmt_money(res.credit[0]), fmt_money(res.credit[1]), fmt_money(res.credit[2])],
        ["4", "Cheltuieli privind dobânzile aferente (ct. 666)", "lei", fmt_money(res.dobanzi[0]), fmt_money(res.dobanzi[1]), fmt_money(res.dobanzi[2])],
        ["5", "Rd (Rata medie a dobânzii creditului = Dobânzi / Credit)", "%", fmt_ratio_pct(res.rd_rate[0]), fmt_ratio_pct(res.rd_rate[1]), fmt_ratio_pct(res.rd_rate[2])],
        ["6", "Diferențialul de rentabilitate (Re - Rd)", "pp", fmt_ratio_pct(res.re_rate[0] - res.rd_rate[0], plus_sign=True), fmt_ratio_pct(res.re_rate[1] - res.rd_rate[1], plus_sign=True), fmt_ratio_pct(res.re_rate[2] - res.rd_rate[2], plus_sign=True)],
        ["7", "Efectul de levier financiar = (Re - Rd) × (Credit / CP)", "%", fmt_ratio_pct(res.efect_levier[0], plus_sign=True), fmt_ratio_pct(res.efect_levier[1], plus_sign=True), fmt_ratio_pct(res.efect_levier[2], plus_sign=True)],
        ["8", "Diagnostic efect de îndatorare", "verdict", res.verdict[0], res.verdict[1], res.verdict[2]],
    ]
    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[0.9, 7.2, 1.3, 2.8, 2.8, 2.8],
        total_rows=[1, 4, 6, 7],
        font_size=7.8,
    )

    add_paragraph(
        doc,
        f"Evaluarea efectului de levier: În perioada de raportare 2, rata rentabilității capitalurilor proprii este "
        f"Re = {fmt_ratio_pct(res.re_rate[2])}, iar rata medie a dobânzii aferente creditelor bancare pe termen lung este "
        f"Rd = {fmt_ratio_pct(res.rd_rate[2])}. Diferențialul de rentabilitate (Re - Rd = "
        f"{fmt_ratio_pct(res.re_rate[2] - res.rd_rate[2], plus_sign=True)}) determină un efect de levier financiar de "
        f"{fmt_ratio_pct(res.efect_levier[0], plus_sign=True)} în perioada de raportare 0, "
        f"{fmt_ratio_pct(res.efect_levier[1], plus_sign=True)} în perioada de raportare 1 și "
        f"{fmt_ratio_pct(res.efect_levier[2], plus_sign=True)} în perioada de raportare 2 ({res.verdict[2]}).",
    )
