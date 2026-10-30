#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 8: Analiza Indicatorilor de Lichiditate (Curentă, Rapidă și Imediată).
Prezintă evoluția celor trei trepte de lichiditate (curentă, rapidă și imediată)
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
    fmt_ratio_pct,
    make_table,
    pick_variant,
    safe_div,
)
from sheet_01_bilant import BilantResult
from sheet_03_bilant_financiar import BilantFinanciarResult


@dataclass
class LichiditateResult:
    l_curenta: Tuple[float, float, float]
    l_rapida: Tuple[float, float, float]
    l_imediata: Tuple[float, float, float]
    l_curenta_exig: Tuple[float, float, float]
    l_rapida_exig: Tuple[float, float, float]
    l_imediata_exig: Tuple[float, float, float]


def compute_sheet_08_lichiditatea(bf: BilantFinanciarResult, bil: BilantResult) -> LichiditateResult:
    ac = bf.active_circulante
    dts = bf.datorii_termen_scurt
    dts_exig = bil.rd(11)

    cr_inv_disp = tuple(round(bf.creante[p] + bf.investitii_ts[p] + bf.disponibilitati[p], 2) for p in range(3))
    inv_disp = tuple(round(bf.investitii_ts[p] + bf.disponibilitati[p], 2) for p in range(3))

    lc = tuple(safe_div(ac[p], dts[p]) for p in range(3))
    lr = tuple(safe_div(cr_inv_disp[p], dts[p]) for p in range(3))
    li = tuple(safe_div(inv_disp[p], dts[p]) for p in range(3))

    lc_ex = tuple(safe_div(ac[p], dts_exig[p]) for p in range(3))
    lr_ex = tuple(safe_div(cr_inv_disp[p], dts_exig[p]) for p in range(3))
    li_ex = tuple(safe_div(inv_disp[p], dts_exig[p]) for p in range(3))

    return LichiditateResult(
        l_curenta=lc,
        l_rapida=lr,
        l_imediata=li,
        l_curenta_exig=lc_ex,
        l_rapida_exig=lr_ex,
        l_imediata_exig=li_ex,
    )


def render_sheet_08_lichiditatea(
    doc: Document, bal: BalantaData, bf: BilantFinanciarResult, bil: BilantResult, res: LichiditateResult
) -> None:
    """Generează secțiunea Analizei Indicatorilor de Lichiditate în documentul Word."""
    add_heading(doc, "8. Analiza Indicatorilor de Lichiditate (Curentă, Rapidă și Imediată)", level=1)

    ac = bf.active_circulante
    dts = bf.datorii_termen_scurt
    cr_inv_disp = tuple(round(bf.creante[p] + bf.investitii_ts[p] + bf.disponibilitati[p], 2) for p in range(3))
    inv_disp = tuple(round(bf.investitii_ts[p] + bf.disponibilitati[p], 2) for p in range(3))

    d_lc, i_lc = calc_delta(res.l_curenta), calc_index(res.l_curenta)
    d_lr, i_lr = calc_delta(res.l_rapida), calc_index(res.l_rapida)
    d_li, i_li = calc_delta(res.l_imediata), calc_index(res.l_imediata)

    headers = [
        "Element / Indicator de Lichiditate",
        "U.M. / Plafon",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
        "Δ (P1 - P0)",
        "Δ (P2 - P1)",
        "I (P2 / P1)",
    ]
    rows = [
        [
            "Active circulante totale (AC)",
            "lei",
            fmt_money(ac[0]),
            fmt_money(ac[1]),
            fmt_money(ac[2]),
            fmt_money(ac[1] - ac[0]),
            fmt_money(ac[2] - ac[1]),
            fmt_ratio_pct(calc_index(ac)[1]),
        ],
        [
            "Creanțe + Investiții TS + Disponibilități (AC - Stocuri)",
            "lei",
            fmt_money(cr_inv_disp[0]),
            fmt_money(cr_inv_disp[1]),
            fmt_money(cr_inv_disp[2]),
            fmt_money(cr_inv_disp[1] - cr_inv_disp[0]),
            fmt_money(cr_inv_disp[2] - cr_inv_disp[1]),
            fmt_ratio_pct(calc_index(cr_inv_disp)[1]),
        ],
        [
            "Disponibilități bănești + Investiții TS",
            "lei",
            fmt_money(inv_disp[0]),
            fmt_money(inv_disp[1]),
            fmt_money(inv_disp[2]),
            fmt_money(inv_disp[1] - inv_disp[0]),
            fmt_money(inv_disp[2] - inv_disp[1]),
            fmt_ratio_pct(calc_index(inv_disp)[1]),
        ],
        [
            "Datorii pe termen scurt (DTS)",
            "lei",
            fmt_money(dts[0]),
            fmt_money(dts[1]),
            fmt_money(dts[2]),
            fmt_money(dts[1] - dts[0]),
            fmt_money(dts[2] - dts[1]),
            fmt_ratio_pct(calc_index(dts)[1]),
        ],
        [
            "1. Lichiditatea curentă (Active circulante / DTS)",
            "100% – 200%",
            fmt_ratio_pct(res.l_curenta[0]),
            fmt_ratio_pct(res.l_curenta[1]),
            fmt_ratio_pct(res.l_curenta[2]),
            fmt_ratio_pct(d_lc[0], plus_sign=True),
            fmt_ratio_pct(d_lc[1], plus_sign=True),
            fmt_ratio_pct(i_lc[1]),
        ],
        [
            "2. Lichiditatea rapidă ((AC - Stocuri) / DTS)",
            "60% – 100%",
            fmt_ratio_pct(res.l_rapida[0]),
            fmt_ratio_pct(res.l_rapida[1]),
            fmt_ratio_pct(res.l_rapida[2]),
            fmt_ratio_pct(d_lr[0], plus_sign=True),
            fmt_ratio_pct(d_lr[1], plus_sign=True),
            fmt_ratio_pct(i_lr[1]),
        ],
        [
            "3. Lichiditatea imediată (Disponibilități / DTS)",
            "20% – 60%",
            fmt_ratio_pct(res.l_imediata[0]),
            fmt_ratio_pct(res.l_imediata[1]),
            fmt_ratio_pct(res.l_imediata[2]),
            fmt_ratio_pct(d_li[0], plus_sign=True),
            fmt_ratio_pct(d_li[1], plus_sign=True),
            fmt_ratio_pct(i_li[1]),
        ],
    ]

    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[5.4, 2.0, 1.9, 1.9, 1.9, 1.6, 1.6, 1.5],
        total_rows=[4, 5, 6],
        font_size=7.8,
    )

    traj_lc = describe_trajectory(res.l_curenta, bal, "s08_lc", "lichiditatea curentă (generală)", is_ratio=True)
    lc2, lr2, li2 = res.l_curenta[2], res.l_rapida[2], res.l_imediata[2]

    if lc2 > 2.0:
        liq_verdict = pick_variant(
            bal,
            "s08_high",
            [
                f"În perioada de raportare 2, toate cele trei rate depășesc limitele superioare uzuale (lichiditatea curentă = {fmt_ratio_pct(lc2)}, lichiditatea rapidă = {fmt_ratio_pct(lr2)}, lichiditatea imediată = {fmt_ratio_pct(li2)}), indicând o capacitate foarte ridicată de onorare pe loc a angajamentelor scadente și existența unui excedent de lichidități ({fmt_money(bf.disponibilitati[2])} lei).",
                f"Nivelurile consemnate în perioada de raportare 2 (lichiditate curentă de {fmt_ratio_pct(lc2)}, rapidă de {fmt_ratio_pct(lr2)} și imediată de {fmt_ratio_pct(li2)}) reflectă o acoperire supraunitară amplă a datoriilor pe termen scurt ({fmt_money(bf.datorii_termen_scurt[2])} lei) din activele curente.",
            ],
        )
    elif lc2 >= 1.2:
        liq_verdict = pick_variant(
            bal,
            "s08_opt",
            [
                f"În perioada de raportare 2, indicatorii se încadrează optim în plafoanele prudențiale bancare: lichiditatea curentă este de {fmt_ratio_pct(lc2)} (interval [120%–200%]), lichiditatea rapidă este de {fmt_ratio_pct(lr2)}, iar lichiditatea imediată atinge {fmt_ratio_pct(li2)} pe baza disponibilităților de {fmt_money(bf.disponibilitati[2])} lei.",
                f"Capacitatea de plată pe termen scurt în perioada de raportare 2 este echilibrată, cu o lichiditate generală de {fmt_ratio_pct(lc2)}, o lichiditate redusă (fără stocuri) de {fmt_ratio_pct(lr2)} și o lichiditate imediată (efectivă) de {fmt_ratio_pct(li2)}.",
            ],
        )
    else:
        liq_verdict = (
            f"În perioada de raportare 2, lichiditatea curentă se situează la {fmt_ratio_pct(lc2)}, lichiditatea rapidă la {fmt_ratio_pct(lr2)}, iar lichiditatea imediată la {fmt_ratio_pct(li2)}, semnalând necesitatea sincronizării atente a încasărilor curente cu scadențarul datoriilor pe termen scurt ({fmt_money(bf.datorii_termen_scurt[2])} lei)."
        )

    add_paragraph(doc, f"Evaluarea capacității de plată pe termen scurt: Pe parcursul celor trei perioade, {traj_lc}. {liq_verdict}")
