#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 16: Analiza Efectului de Levier Financiar (Efectul de Îndatorare).

Relația fundamentală (înainte de impozit):
    Rf = Re + (Re - Rd) × Df / CP
unde
    Re = (Rezultatul brut + Cheltuielile cu dobânzile) / (CP + Df)  — rentabilitatea economică,
    Rd = Cheltuielile cu dobânzile / Df                              — costul mediu al îndatorării,
    Df = datoriile financiare purtătoare de dobândă (ct. 16x, 519, descoperit 512, 455),
    Rf = Rezultatul brut / CP                                        — rentabilitatea financiară.
Fluxurile (rezultat, dobânzi) sunt cumulate de la începutul anului și se ANUALIZEAZĂ
(× zile an / zile perioadă) pentru a fi comparabile cu ratele anuale de dobândă.
Pentru P0 (solduri de deschidere) indicatorii nu sunt disponibili.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Tuple

from docx import Document

from core_balanta import (
    ND,
    BalantaData,
    add_heading,
    add_paragraph,
    all_ok,
    fmt_money,
    fmt_pp,
    fmt_ratio_pct,
    make_table,
    safe_div,
)
from sheet_02_cpp import CPPResult
from sheet_03_bilant_financiar import BilantFinanciarResult
from sheet_05_structura_financiara import StructuraFinanciaraResult

Triple = Tuple[float, float, float]


@dataclass
class EfectIndatorareResult:
    cp: Triple
    re_rate: Triple          # rentabilitatea economică (anualizată)
    credit: Triple           # datorii financiare (Df)
    dobanzi: Triple          # cheltuieli cu dobânzile cumulate
    rd_rate: Triple          # rata medie a dobânzii (anualizată)
    efect_levier: Triple     # (Re - Rd) × Df / CP
    verdict: Tuple[str, str, str]
    rf_rate: Triple = (ND, ND, ND)       # rentabilitatea financiară înainte de impozit (anualizată)
    rezultat_brut: Triple = (ND, ND, ND)
    levier: Triple = (ND, ND, ND)        # Df / CP


def compute_sheet_16_efectul_indatorare(
    bf: BilantFinanciarResult,
    sf: StructuraFinanciaraResult,
    cpp: CPPResult,
    annual_factors: Tuple[float, float, float] = (ND, 365 / 243, 365 / 273),
) -> EfectIndatorareResult:
    cp = sf.cp
    df = bf.datorii_financiare
    dob = cpp.dobanzi
    rb = cpp.rezultat_brut

    re_r, rd_r, rf_r, lev, ef = [], [], [], [], []
    verdicts = []
    for p, lbl in enumerate(("perioada de raportare 0", "perioada de raportare 1", "perioada de raportare 2")):
        k = annual_factors[p]
        cap_inv = cp[p] + df[p]
        if not all_ok(rb[p], dob[p], k) or cp[p] <= 0 or cap_inv <= 0:
            re_r.append(ND), rd_r.append(ND), rf_r.append(ND), ef.append(ND)
            lev.append(df[p] / cp[p] if cp[p] > 0 else ND)
            verdicts.append(
                "n.d." if p == 0 else f"În {lbl}, efectul de levier nu are sens economic (capitaluri proprii ≤ 0)"
            )
            continue
        re_p = (rb[p] + dob[p]) * k / cap_inv
        rd_p = safe_div(dob[p] * k, df[p], default=0.0)
        lv = df[p] / cp[p]
        ef_p = (re_p - rd_p) * lv
        re_r.append(re_p), rd_r.append(rd_p), lev.append(lv), ef.append(ef_p), rf_r.append(re_p + ef_p)
        if df[p] <= 0.005:
            verdicts.append(f"În {lbl}, societatea nu are datorii financiare: efect de levier nul")
        elif re_p >= rd_p:
            verdicts.append("EFECT DE LEVIER pozitiv (Re > Rd)")
        else:
            verdicts.append("EFECT DE MĂCIUCĂ (Re < Rd)")

    return EfectIndatorareResult(
        cp=cp,
        re_rate=tuple(re_r),  # type: ignore[arg-type]
        credit=df,
        dobanzi=dob,
        rd_rate=tuple(rd_r),  # type: ignore[arg-type]
        efect_levier=tuple(ef),  # type: ignore[arg-type]
        verdict=(verdicts[0], verdicts[1], verdicts[2]),
        rf_rate=tuple(rf_r),  # type: ignore[arg-type]
        rezultat_brut=rb,
        levier=tuple(lev),  # type: ignore[arg-type]
    )


def render_sheet_16_efectul_indatorare(doc: Document, bal: BalantaData, res: EfectIndatorareResult) -> None:
    """Generează secțiunea Efectului de Levier Financiar în documentul Word."""
    add_heading(doc, "16. Analiza Efectului de Levier Financiar (Efectul de Îndatorare)", level=1)

    headers = ["Nr.", "Elemente ale efectului de levier financiar", "U.M.", *bal.col_labels]
    diff = tuple((res.re_rate[p] - res.rd_rate[p]) for p in range(3))
    rows = [
        ["1", "Capitaluri proprii (CP)", "lei", *[fmt_money(v) for v in res.cp]],
        ["2", "Datorii financiare purtătoare de dobândă (Df: ct. 16x, 519, 455)", "lei", *[fmt_money(v) for v in res.credit]],
        ["3", "Brațul levierului (Df / CP)", "coef.", *[fmt_money(v) for v in res.levier]],
        ["4", "Rezultatul brut cumulat", "lei", *[fmt_money(v) for v in res.rezultat_brut]],
        ["5", "Cheltuieli cu dobânzile cumulate (ct. 666)", "lei", *[fmt_money(v) for v in res.dobanzi]],
        ["6", "Re — rentabilitatea economică anualizată ((RB + Dobânzi) / (CP + Df))", "%", *[fmt_ratio_pct(v) for v in res.re_rate]],
        ["7", "Rd — rata medie anualizată a dobânzii (Dobânzi / Df)", "%", *[fmt_ratio_pct(v) for v in res.rd_rate]],
        ["8", "Diferențialul de rentabilitate (Re - Rd)", "pp", *[fmt_pp(v) for v in diff]],
        ["9", "Efectul de levier = (Re - Rd) × Df / CP", "pp", *[fmt_pp(v) for v in res.efect_levier]],
        ["10", "Rf — rentabilitatea financiară anualizată înainte de impozit (Re + efect levier)", "%", *[fmt_ratio_pct(v) for v in res.rf_rate]],
        ["11", "Diagnostic efect de îndatorare", "verdict", *res.verdict],
    ]
    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[0.8, 7.4, 1.2, 2.8, 2.8, 2.8],
        total_rows=[5, 6, 8, 9],
        font_size=7.8,
    )

    re2, rd2, ef2, df2 = res.re_rate[2], res.rd_rate[2], res.efect_levier[2], res.credit[2]
    if not all_ok(re2, rd2, ef2):
        msg = f"Evaluarea efectului de levier: {res.verdict[2]}."
    elif df2 <= 0.005:
        msg = (
            f"Evaluarea efectului de levier: societatea nu are datorii financiare purtătoare de dobândă în perioada de raportare 2; "
            f"rentabilitatea financiară coincide cu rentabilitatea economică (Re = {fmt_ratio_pct(re2)})."
        )
    elif re2 >= rd2:
        msg = (
            f"Evaluarea efectului de levier: în perioada de raportare 2, rentabilitatea economică anualizată (Re = {fmt_ratio_pct(re2)}) "
            f"depășește costul mediu al îndatorării (Rd = {fmt_ratio_pct(rd2)}); la un braț al levierului de {fmt_money(res.levier[2])}, "
            f"îndatorarea adaugă {fmt_pp(ef2)} la rentabilitatea financiară, care ajunge la {fmt_ratio_pct(res.rf_rate[2])} înainte de impozit."
        )
    else:
        msg = (
            f"Evaluarea efectului de levier: în perioada de raportare 2, costul mediu al îndatorării (Rd = {fmt_ratio_pct(rd2)}) "
            f"depășește rentabilitatea economică (Re = {fmt_ratio_pct(re2)}); la un braț al levierului de {fmt_money(res.levier[2])}, "
            f"îndatorarea reduce rentabilitatea financiară cu {fmt_pp(-ef2)} (efect de măciucă), până la {fmt_ratio_pct(res.rf_rate[2])}."
        )
    add_paragraph(
        doc,
        msg
        + " Ratele sunt anualizate (rulajele cumulate × zile an / zile perioadă) și calculate pe soldurile de la finele perioadei; "
        "împrumuturile asociaților (ct. 455) sunt considerate datorii purtătoare de dobândă.",
    )
