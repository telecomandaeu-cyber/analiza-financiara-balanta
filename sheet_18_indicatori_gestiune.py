#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 18: Analiza Indicatorilor de Gestiune (Viteza de Rotație a Elementelor Patrimoniale).
Cuantifică viteza de rotație (în număr de rotații) și durata medie a unei rotații (în zile)
pentru stocuri, creanțe și datorii pe cele trei perioade de raportare (0, 1 și 2).
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
    fmt_num,
    fmt_ratio_pct,
    make_table,
    pick_variant,
    safe_div,
)
from sheet_02_cpp import CPPResult
from sheet_03_bilant_financiar import BilantFinanciarResult


@dataclass
class IndicatoriGestiuneResult:
    vit_stoc: Tuple[float, float, float]
    dur_stoc: Tuple[float, float, float]
    vit_creante: Tuple[float, float, float]
    dur_creante: Tuple[float, float, float]
    vit_dat: Tuple[float, float, float]
    dur_dat: Tuple[float, float, float]
    dur_clienti_411: Tuple[float, float, float]
    dur_furnizori_401: Tuple[float, float, float]


def compute_sheet_18_indicatori_gestiune(
    bal: BalantaData, bf: BilantFinanciarResult, cpp: CPPResult
) -> IndicatoriGestiuneResult:
    stoc = bf.stocuri
    cr = bf.creante
    dat = bf.datorii_termen_scurt
    ca = cpp.cifra_afaceri

    vit_s = tuple(safe_div(ca[p], stoc[p]) for p in range(3))
    dur_s = tuple(safe_div(stoc[p], ca[p]) * 360.0 for p in range(3))

    vit_c = tuple(safe_div(ca[p], cr[p]) for p in range(3))
    dur_c = tuple(safe_div(cr[p], ca[p]) * 365.0 for p in range(3))

    vit_d = tuple(safe_div(ca[p], dat[p]) for p in range(3))
    dur_d = tuple(safe_div(dat[p], ca[p]) * 365.0 for p in range(3))

    dur_411 = tuple(safe_div(bal.pos_d("411", p), ca[p]) * 365.0 for p in range(3))
    dur_401 = tuple(safe_div(bal.pos_c("401", p), ca[p]) * 365.0 for p in range(3))

    return IndicatoriGestiuneResult(
        vit_stoc=vit_s,
        dur_stoc=dur_s,
        vit_creante=vit_c,
        dur_creante=dur_c,
        vit_dat=vit_d,
        dur_dat=dur_d,
        dur_clienti_411=dur_411,
        dur_furnizori_401=dur_401,
    )


def render_sheet_18_indicatori_gestiune(
    doc: Document,
    bal: BalantaData,
    bf: BilantFinanciarResult,
    cpp: CPPResult,
    res: IndicatoriGestiuneResult,
) -> None:
    """Generează secțiunea Indicatorilor de Gestiune într-un tabel unificat."""
    add_heading(doc, "18. Analiza Indicatorilor de Gestiune (Viteza de Rotație a Stocurilor, Creanțelor și Datoriilor)", level=1)

    headers = [
        "Indicator de Gestiune (Viteza în rotații / Durata în zile)",
        "U.M.",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
        "Δ (P1 - P0)",
        "Δ (P2 - P1)",
        "I (P2 / P1)",
    ]

    items = [
        ("1. Viteza de rotație a stocurilor (CA / Stocuri)", "rotații", res.vit_stoc, 2, False),
        ("   • Durata medie de rotație a stocurilor (360 zile)", "zile", res.dur_stoc, 2, True),
        ("2. Viteza de rotație a creanțelor totale (CA / Creanțe)", "rotații", res.vit_creante, 4, False),
        ("   • Durata medie de rotație a creanțelor totale (365 zile)", "zile", res.dur_creante, 2, True),
        ("   • Durata medie de încasare a clienților comerciali (ct. 411)", "zile", res.dur_clienti_411, 2, True),
        ("3. Viteza de rotație a datoriilor pe termen scurt (CA / DTS)", "rotații", res.vit_dat, 4, False),
        ("   • Durata medie de rotație a datoriilor pe termen scurt (365 zile)", "zile", res.dur_dat, 2, True),
        ("   • Durata medie de plată a furnizorilor comerciali (ct. 401)", "zile", res.dur_furnizori_401, 2, True),
    ]

    rows = []
    tot_idx = []
    for idx_r, (lbl, um, vals, dec, is_t) in enumerate(items):
        if is_t:
            tot_idx.append(idx_r)
        d = calc_delta(vals)
        i = calc_index(vals)
        rows.append([
            lbl,
            um,
            fmt_num(vals[0], dec),
            fmt_num(vals[1], dec),
            fmt_num(vals[2], dec),
            fmt_num(d[0], dec),
            fmt_num(d[1], dec),
            fmt_ratio_pct(i[1]),
        ])

    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[5.8, 1.4, 1.9, 1.9, 1.9, 1.7, 1.7, 1.5],
        total_rows=tot_idx,
        font_size=7.7,
    )

    ds0, ds2 = res.dur_stoc[0], res.dur_stoc[2]
    dc0, dc2 = res.dur_creante[0], res.dur_creante[2]
    dd0, dd2 = res.dur_dat[0], res.dur_dat[2]
    d_cli2 = res.dur_clienti_411[2]
    d_furn2 = res.dur_furnizori_401[2]

    if d_cli2 <= d_furn2:
        decalaj_com = f"clienții comerciali (ct. 411) sunt încasați în medie la {fmt_num(d_cli2, 1)} zile, mai rapid decât termenul mediu de plată a furnizorilor (ct. 401 = {fmt_num(d_furn2, 1)} zile), generând un decalaj comercial favorabil trezoreriei"
    else:
        decalaj_com = f"termenul mediu de încasare a clienților comerciali (ct. 411 = {fmt_num(d_cli2, 1)} zile) depășește durata medie de achitare a furnizorilor (ct. 401 = {fmt_num(d_furn2, 1)} zile), ceea ce implică finanțarea creditului-client din resurse proprii"

    s18_msg = pick_variant(
        bal,
        "s18_gest",
        [
            f"Eficiența gestiunii ciclului operațional: În perioada de raportare 2, stocurile se reînnoiesc în medie la {fmt_num(ds2, 2)} zile ({fmt_num(res.vit_stoc[2], 2)} rotații, față de {fmt_num(ds0, 2)} zile în perioada de raportare 0), creanțele totale au o durată medie de rotație de {fmt_num(dc2, 2)} zile ({fmt_num(res.vit_creante[2], 2)} rotații, față de {fmt_num(dc0, 2)} zile în perioada 0), iar datoriile pe termen scurt se rotează la {fmt_num(dd2, 2)} zile ({fmt_num(res.vit_dat[2], 2)} rotații). La nivel strict comercial, în perioada de raportare 2 {decalaj_com}.",
            f"Analiza vitezelor de rotație: Comparativ cu perioada de raportare 0, în perioada de raportare 2 durata medie de stocare se situează la {fmt_num(ds2, 2)} zile ({fmt_num(res.vit_stoc[2], 2)} rotații/perioadă), durata de recuperare a creanțelor totale la {fmt_num(dc2, 2)} zile ({fmt_num(res.vit_creante[2], 2)} rotații), iar durata de exigibilitate a datoriilor curente la {fmt_num(dd2, 2)} zile (față de {fmt_num(dd0, 2)} zile în perioada 0). Pe relația comercială directă din perioada de raportare 2, {decalaj_com}.",
            f"Diagnostic operațional al rotației capitalurilor: La finele perioadei de raportare 2, ciclul de conversie a activelor circulante este caracterizat de o rotație a stocurilor la {fmt_num(ds2, 2)} zile, o rotație a creanțelor totale la {fmt_num(dc2, 2)} zile și o rotație a datoriilor pe termen scurt la {fmt_num(dd2, 2)} zile; în planul creditului comercial din perioada de raportare 2, {decalaj_com}.",
        ],
    )
    add_paragraph(doc, s18_msg)
