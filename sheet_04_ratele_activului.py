#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 4: Analiza Ratelor de Structură ale Activului.
Evaluează structura și dinamica activelor pe două paliere (ponderea în categorie și ponderea în Total Activ),
precum și corelația dinamicii activelor imobilizate și circulante cu cifra de afaceri.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List, Tuple

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
from sheet_02_cpp import CPPResult
from sheet_03_bilant_financiar import BilantFinanciarResult


@dataclass
class AssetSubRateBlock:
    title: str
    elem_name: str
    cat_name: str
    rate1_name: str
    rate2_name: str
    elem_vals: Tuple[float, float, float]
    cat_vals: Tuple[float, float, float]
    tot_vals: Tuple[float, float, float]
    rate1_vals: Tuple[float, float, float]
    rate2_vals: Tuple[float, float, float]


@dataclass
class RateActivResult:
    sub_blocks: List[AssetSubRateBlock] = field(default_factory=list)
    r_ai: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    r_ac: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    semnal_ai: Tuple[str, str] = ("", "")
    semnal_ac: Tuple[str, str] = ("", "")


def _build_sub_block(
    title: str,
    elem_name: str,
    cat_name: str,
    rate1_name: str,
    rate2_name: str,
    elem_vals: Tuple[float, float, float],
    cat_vals: Tuple[float, float, float],
    tot_vals: Tuple[float, float, float],
) -> AssetSubRateBlock:
    r1 = tuple(safe_div(elem_vals[p], cat_vals[p]) for p in range(3))
    r2 = tuple(safe_div(elem_vals[p], tot_vals[p]) for p in range(3))
    return AssetSubRateBlock(
        title=title,
        elem_name=elem_name,
        cat_name=cat_name,
        rate1_name=rate1_name,
        rate2_name=rate2_name,
        elem_vals=elem_vals,
        cat_vals=cat_vals,
        tot_vals=tot_vals,
        rate1_vals=r1,
        rate2_vals=r2,
    )


def compute_sheet_04_ratele_activului(bf: BilantFinanciarResult, cpp: CPPResult) -> RateActivResult:
    res = RateActivResult()
    ai = bf.active_imobilizate
    ac = bf.active_circulante
    ta = bf.total_activ

    inv_disp = tuple(round(bf.investitii_ts[p] + bf.disponibilitati[p], 2) for p in range(3))

    res.sub_blocks = [
        _build_sub_block(
            "Rata imobilizărilor necorporale",
            "Imobilizări necorporale",
            "Active imobilizate",
            "Ponderea în Active imobilizate (Rata 1)",
            "Ponderea în Total Activ (Rata 2)",
            bf.imobilizari_necorporale,
            ai,
            ta,
        ),
        _build_sub_block(
            "Rata imobilizărilor corporale",
            "Imobilizări corporale",
            "Active imobilizate",
            "Ponderea în Active imobilizate (Rata 1)",
            "Ponderea în Total Activ (Rata 2)",
            bf.imobilizari_corporale,
            ai,
            ta,
        ),
        _build_sub_block(
            "Rata imobilizărilor financiare",
            "Imobilizări financiare",
            "Active imobilizate",
            "Ponderea în Active imobilizate (Rata 1)",
            "Ponderea în Total Activ (Rata 2)",
            bf.imobilizari_financiare,
            ai,
            ta,
        ),
        _build_sub_block(
            "Rata stocurilor",
            "Stocuri",
            "Active circulante",
            "Ponderea în Active circulante (Rata 1)",
            "Ponderea în Total Activ (Rata 2)",
            bf.stocuri,
            ac,
            ta,
        ),
        _build_sub_block(
            "Rata creanțelor",
            "Creanțe",
            "Active circulante",
            "Ponderea în Active circulante (Rata 1)",
            "Ponderea în Total Activ (Rata 2)",
            bf.creante,
            ac,
            ta,
        ),
        _build_sub_block(
            "Rata investițiilor TS și disponibilităților",
            "Investiții TS + Disponibilități",
            "Active circulante",
            "Ponderea în Active circulante (Rata 1)",
            "Ponderea în Total Activ (Rata 2)",
            inv_disp,
            ac,
            ta,
        ),
    ]

    res.r_ai = tuple(safe_div(ai[p], ta[p]) for p in range(3))
    res.r_ac = tuple(safe_div(ac[p], ta[p]) for p in range(3))

    i_ai = calc_index(ai)
    i_ac = calc_index(ac)
    i_ca = calc_index(cpp.cifra_afaceri)

    s_ai = []
    for idx in (0, 1):
        if i_ca[idx] is not None and i_ai[idx] is not None and i_ca[idx] >= i_ai[idx]:
            s_ai.append("Semnal favorabil (I_CA >= I_AI)")
        else:
            s_ai.append("Semnal nefavorabil (I_CA < I_AI)")
    res.semnal_ai = (s_ai[0], s_ai[1])

    s_ac = []
    for idx in (0, 1):
        if i_ca[idx] is not None and i_ac[idx] is not None and i_ca[idx] >= i_ac[idx]:
            s_ac.append("Semnal favorabil (I_CA >= I_AC)")
        else:
            s_ac.append("Semnal nefavorabil (I_CA < I_AC)")
    res.semnal_ac = (s_ac[0], s_ac[1])

    return res


def render_sheet_04_ratele_activului(
    doc: Document, bal: BalantaData, bf: BilantFinanciarResult, cpp: CPPResult, res: RateActivResult
) -> None:
    """Generează secțiunea Analizei Ratelor de Structură ale Activului."""
    add_heading(doc, "4. Analiza Ratelor de Structură ale Activului și Corelația cu Cifra de Afaceri", level=1)

    headers = [
        "Rată de Structură a Activului (Ponderea în Categorie / în Total Activ)",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
        "Δ (P1 - P0)",
        "Δ (P2 - P1)",
        "I (P2 / P1)",
    ]

    d_rai = calc_delta(res.r_ai)
    i_rai = calc_index(res.r_ai)
    d_rac = calc_delta(res.r_ac)
    i_rac = calc_index(res.r_ac)

    rows = [
        [
            "A. Rata Activelor Imobilizate (R_AI = Active imobilizate / Total Activ)",
            fmt_ratio_pct(res.r_ai[0]),
            fmt_ratio_pct(res.r_ai[1]),
            fmt_ratio_pct(res.r_ai[2]),
            fmt_ratio_pct(d_rai[0], plus_sign=True),
            fmt_ratio_pct(d_rai[1], plus_sign=True),
            fmt_ratio_pct(i_rai[1]),
        ],
    ]
    tot_idx = [0]

    for b in res.sub_blocks[:2]:
        d1, i1 = calc_delta(b.rate1_vals), calc_index(b.rate1_vals)
        d2, i2 = calc_delta(b.rate2_vals), calc_index(b.rate2_vals)
        rows.append([
            f"  • {b.elem_name} — pondere în Active imobilizate / în Total Activ",
            f"{fmt_ratio_pct(b.rate1_vals[0])} / {fmt_ratio_pct(b.rate2_vals[0])}",
            f"{fmt_ratio_pct(b.rate1_vals[1])} / {fmt_ratio_pct(b.rate2_vals[1])}",
            f"{fmt_ratio_pct(b.rate1_vals[2])} / {fmt_ratio_pct(b.rate2_vals[2])}",
            f"{fmt_ratio_pct(d1[0], plus_sign=True)} / {fmt_ratio_pct(d2[0], plus_sign=True)}",
            f"{fmt_ratio_pct(d1[1], plus_sign=True)} / {fmt_ratio_pct(d2[1], plus_sign=True)}",
            f"{fmt_ratio_pct(i1[1])} / {fmt_ratio_pct(i2[1])}",
        ])

    tot_idx.append(len(rows))
    rows.append([
        "B. Rata Activelor Circulante (R_AC = Active circulante / Total Activ)",
        fmt_ratio_pct(res.r_ac[0]),
        fmt_ratio_pct(res.r_ac[1]),
        fmt_ratio_pct(res.r_ac[2]),
        fmt_ratio_pct(d_rac[0], plus_sign=True),
        fmt_ratio_pct(d_rac[1], plus_sign=True),
        fmt_ratio_pct(i_rac[1]),
    ])

    for b in res.sub_blocks[3:]:
        d1, i1 = calc_delta(b.rate1_vals), calc_index(b.rate1_vals)
        d2, i2 = calc_delta(b.rate2_vals), calc_index(b.rate2_vals)
        rows.append([
            f"  • {b.elem_name} — pondere în Active circulante / în Total Activ",
            f"{fmt_ratio_pct(b.rate1_vals[0])} / {fmt_ratio_pct(b.rate2_vals[0])}",
            f"{fmt_ratio_pct(b.rate1_vals[1])} / {fmt_ratio_pct(b.rate2_vals[1])}",
            f"{fmt_ratio_pct(b.rate1_vals[2])} / {fmt_ratio_pct(b.rate2_vals[2])}",
            f"{fmt_ratio_pct(d1[0], plus_sign=True)} / {fmt_ratio_pct(d2[0], plus_sign=True)}",
            f"{fmt_ratio_pct(d1[1], plus_sign=True)} / {fmt_ratio_pct(d2[1], plus_sign=True)}",
            f"{fmt_ratio_pct(i1[1])} / {fmt_ratio_pct(i2[1])}",
        ])

    i_ca = calc_index(cpp.cifra_afaceri)
    i_ai = calc_index(bf.active_imobilizate)
    i_ac = calc_index(bf.active_circulante)

    tot_idx.extend([len(rows), len(rows) + 1])
    rows.append([
        "Corelație Cifra de afaceri vs. Active imobilizate (I_CA >= I_AI)",
        "—",
        f"I_CA={fmt_ratio_pct(i_ca[0])} vs I_AI={fmt_ratio_pct(i_ai[0])}",
        f"I_CA={fmt_ratio_pct(i_ca[1])} vs I_AI={fmt_ratio_pct(i_ai[1])}",
        res.semnal_ai[0],
        res.semnal_ai[1],
        res.semnal_ai[1],
    ])
    rows.append([
        "Corelație Cifra de afaceri vs. Active circulante (I_CA >= I_AC)",
        "—",
        f"I_CA={fmt_ratio_pct(i_ca[0])} vs I_AC={fmt_ratio_pct(i_ac[0])}",
        f"I_CA={fmt_ratio_pct(i_ca[1])} vs I_AC={fmt_ratio_pct(i_ac[1])}",
        res.semnal_ac[0],
        res.semnal_ac[1],
        res.semnal_ac[1],
    ])

    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[5.8, 2.0, 2.0, 2.0, 2.0, 2.0, 2.0],
        total_rows=tot_idx,
        font_size=7.7,
    )

    traj_rai = describe_trajectory(res.r_ai, bal, "s04_rai", "rata activelor imobilizate (R_AI)", is_ratio=True)
    r_stoc2 = res.sub_blocks[3].rate1_vals[2]
    r_cre2 = res.sub_blocks[4].rate1_vals[2]
    r_disp2 = res.sub_blocks[5].rate1_vals[2]

    if "favorabil" in res.semnal_ai[1] and "favorabil" in res.semnal_ac[1]:
        cor_msg = pick_variant(
            bal,
            "s04_cor_both_ok",
            [
                f"Corelația dinamică în perioada de raportare 2 este dublu favorabilă: indicele cifrei de afaceri (I_CA = {fmt_ratio_pct(i_ca[1])}) depășește atât indicele activelor imobilizate (I_AI = {fmt_ratio_pct(i_ai[1])}), cât și pe cel al activelor circulante (I_AC = {fmt_ratio_pct(i_ac[1])}), atestând o accelerare a vitezei de rotație și o exploatare intensivă a patrimoniului.",
                f"Compararea ritmurilor de creștere între perioada de raportare 1 și perioada de raportare 2 confirmă devansarea dinamicii activelor (I_AI = {fmt_ratio_pct(i_ai[1])}, I_AC = {fmt_ratio_pct(i_ac[1])}) de către cifra de afaceri (I_CA = {fmt_ratio_pct(i_ca[1])}), generând semnal favorabil pe ambele paliere.",
            ],
        )
    elif "favorabil" not in res.semnal_ai[1] and "favorabil" not in res.semnal_ac[1]:
        cor_msg = pick_variant(
            bal,
            "s04_cor_both_bad",
            [
                f"În intervalul P2/P1, indicele cifrei de afaceri (I_CA = {fmt_ratio_pct(i_ca[1])}) se situează sub indicii de creștere ai activelor imobilizate ({fmt_ratio_pct(i_ai[1])}) și circulante ({fmt_ratio_pct(i_ac[1])}), semnalând o acumulare de active al căror potențial comercial urmează a fi valorificat în exercițiile viitoare.",
                f"Dinamica dintre perioada de raportare 1 și perioada de raportare 2 indică o creștere a bazei de active (I_AI = {fmt_ratio_pct(i_ai[1])}, I_AC = {fmt_ratio_pct(i_ac[1])}) mai rapidă decât cea a cifrei de afaceri (I_CA = {fmt_ratio_pct(i_ca[1])}), ceea ce temperează temporar randamentul mijloacelor economice.",
            ],
        )
    else:
        cor_msg = (
            f"Raportul dintre dinamica cifrei de afaceri (I_CA = {fmt_ratio_pct(i_ca[1])}) și cea a activelor în perioada de raportare 2 este mixt: {res.semnal_ai[1]} pentru activele imobilizate (I_AI = {fmt_ratio_pct(i_ai[1])}) și {res.semnal_ac[1]} pentru activele circulante (I_AC = {fmt_ratio_pct(i_ac[1])})."
        )

    add_paragraph(
        doc,
        f"Analiza structurii activului: Pe ansamblul intervalului, {traj_rai}, în timp ce rata activelor circulante (R_AC) se situează la "
        f"{fmt_ratio_pct(res.r_ac[2])} în perioada de raportare 2 ({fmt_money(bf.active_circulante[2])} lei). În structura internă a activelor "
        f"circulante din perioada de raportare 2, stocurile dețin {fmt_ratio_pct(r_stoc2)}, creanțele reprezintă {fmt_ratio_pct(r_cre2)}, "
        f"iar disponibilitățile bănești și investițiile pe termen scurt însumează {fmt_ratio_pct(r_disp2)}. {cor_msg}",
    )
