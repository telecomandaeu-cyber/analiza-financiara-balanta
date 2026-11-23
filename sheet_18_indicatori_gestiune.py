#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 18: Analiza Indicatorilor de Gestiune (Viteza de Rotație și Durata în Zile).

Rulajele sunt cumulate de la începutul anului, așa că duratele se calculează pe numărul
REAL de zile acoperite de fiecare perioadă (ex. 243 de zile pentru ianuarie–august):
    Durata (zile)          = Sold la finele perioadei / Rulaj cumulat × Zile perioadă
    Viteza (rotații / an)  = Rulaj cumulat anualizat / Sold la finele perioadei
Bazele de raportare: cifra de afaceri pentru stocuri, creanțe și clienți; cumpărările
de la terți (consumuri externe + mărfuri) pentru furnizori. Clienții și furnizorii includ TVA,
în timp ce rulajele sunt fără TVA, deci duratele comerciale sunt ușor supraevaluate.
Pentru P0 (solduri de deschidere, fără rulaje) indicatorii nu sunt disponibili.
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
    fmt_num,
    make_table,
    pick_variant,
)
from sheet_02_cpp import CPPResult
from sheet_03_bilant_financiar import BilantFinanciarResult

Triple = Tuple[float, float, float]


@dataclass
class IndicatoriGestiuneResult:
    vit_stoc: Triple
    dur_stoc: Triple
    vit_creante: Triple
    dur_creante: Triple
    vit_dat: Triple
    dur_dat: Triple
    dur_clienti_411: Triple
    dur_furnizori_401: Triple
    ciclu_conversie: Triple = (ND, ND, ND)


def _dur(sold: float, flux: float, zile: int) -> float:
    if not all_ok(sold, flux) or zile <= 0 or flux <= 0:
        return ND
    return sold / flux * zile


def _vit(sold: float, flux: float, zile: int, zile_an: int) -> float:
    if not all_ok(sold, flux) or zile <= 0 or sold <= 0:
        return ND
    return flux * zile_an / zile / sold


def compute_sheet_18_indicatori_gestiune(
    bal: BalantaData, bf: BilantFinanciarResult, cpp: CPPResult
) -> IndicatoriGestiuneResult:
    zile = bal.period.days
    za = bal.period.days_in_year
    stoc = bf.stocuri
    cr = bf.creante
    dat = bf.datorii_termen_scurt
    ca = cpp.cifra_afaceri
    g = lambda k, p: cpp.it(k)[p]  # noqa: E731
    cumparari = tuple(
        g("mat_prime", p) + g("alte_mat", p) + g("energie", p) + g("prest_ext", p) + g("marfuri", p) - g("reduceri_primite", p)
        for p in range(3)
    )
    clienti = tuple(bal.sold_d("411", p) + bal.sold_d("418", p) for p in range(3))
    furnizori = tuple(bal.sold_c("401", p) + bal.sold_c("404", p) + bal.sold_c("408", p) for p in range(3))

    t = lambda fn: tuple(fn(p) for p in range(3))  # noqa: E731
    dur_s = t(lambda p: _dur(stoc[p], ca[p], zile[p]))
    dur_cli = t(lambda p: _dur(clienti[p], ca[p], zile[p]))
    dur_furn = t(lambda p: _dur(furnizori[p], cumparari[p], zile[p]))
    return IndicatoriGestiuneResult(
        vit_stoc=t(lambda p: _vit(stoc[p], ca[p], zile[p], za)),
        dur_stoc=dur_s,
        vit_creante=t(lambda p: _vit(cr[p], ca[p], zile[p], za)),
        dur_creante=t(lambda p: _dur(cr[p], ca[p], zile[p])),
        vit_dat=t(lambda p: _vit(dat[p], ca[p], zile[p], za)),
        dur_dat=t(lambda p: _dur(dat[p], ca[p], zile[p])),
        dur_clienti_411=dur_cli,
        dur_furnizori_401=dur_furn,
        ciclu_conversie=t(lambda p: dur_s[p] + dur_cli[p] - dur_furn[p]),
    )


def render_sheet_18_indicatori_gestiune(
    doc: Document,
    bal: BalantaData,
    bf: BilantFinanciarResult,
    cpp: CPPResult,
    res: IndicatoriGestiuneResult,
) -> None:
    """Generează secțiunea Indicatorilor de Gestiune."""
    add_heading(doc, "18. Analiza Indicatorilor de Gestiune (Viteza de Rotație a Stocurilor, Creanțelor și Datoriilor)", level=1)

    z = bal.period.days
    headers = ["Indicator de gestiune", "U.M.", *bal.col_labels, "Δ (P2 - P1)"]
    items = [
        ("1. Viteza de rotație a stocurilor (CA anualizată / Stocuri)", "rotații/an", res.vit_stoc, 2, False),
        ("   • Durata de rotație a stocurilor (Stocuri / CA × zile)", "zile", res.dur_stoc, 1, True),
        ("2. Viteza de rotație a creanțelor (CA anualizată / Creanțe)", "rotații/an", res.vit_creante, 2, False),
        ("   • Durata de rotație a creanțelor totale", "zile", res.dur_creante, 1, True),
        ("   • Durata de încasare a clienților (ct. 411 + 418)", "zile", res.dur_clienti_411, 1, True),
        ("3. Viteza de rotație a datoriilor pe termen scurt (CA anualizată / DTS)", "rotații/an", res.vit_dat, 2, False),
        ("   • Durata de rotație a datoriilor pe termen scurt", "zile", res.dur_dat, 1, True),
        ("   • Durata de plată a furnizorilor (ct. 401 + 404 + 408 / cumpărări)", "zile", res.dur_furnizori_401, 1, True),
        ("4. Ciclul de conversie a numerarului (stocuri + clienți - furnizori)", "zile", res.ciclu_conversie, 1, True),
    ]
    rows = []
    tot_idx = []
    for i, (lbl, um, vals, dec, is_t) in enumerate(items):
        if is_t:
            tot_idx.append(i)
        rows.append([lbl, um, fmt_num(vals[0], dec), fmt_num(vals[1], dec), fmt_num(vals[2], dec), fmt_num(vals[2] - vals[1], dec)])

    make_table(doc, headers=headers, rows=rows, col_widths=[6.6, 1.5, 1.9, 1.9, 1.9, 1.7], total_rows=tot_idx, font_size=7.7)

    ds2, dc2, dd2 = res.dur_stoc[2], res.dur_creante[2], res.dur_dat[2]
    d_cli2, d_furn2 = res.dur_clienti_411[2], res.dur_furnizori_401[2]
    if not all_ok(d_cli2, d_furn2):
        decalaj_com = "decalajul dintre încasarea clienților și plata furnizorilor nu poate fi determinat"
    elif d_cli2 <= d_furn2:
        decalaj_com = (
            f"clienții sunt încasați în medie în {fmt_num(d_cli2, 1)} zile, mai repede decât sunt plătiți furnizorii "
            f"({fmt_num(d_furn2, 1)} zile), ceea ce degajă un decalaj comercial favorabil trezoreriei"
        )
    else:
        decalaj_com = (
            f"durata de încasare a clienților ({fmt_num(d_cli2, 1)} zile) depășește durata de plată a furnizorilor "
            f"({fmt_num(d_furn2, 1)} zile), creditul acordat clienților fiind finanțat din resursele proprii ale societății"
        )
    s18 = pick_variant(
        bal,
        "s18_gest",
        [
            f"Eficiența gestiunii ciclului operațional: raportat la cele {z[2]} de zile ale perioadei de raportare 2, stocurile se rotesc în {fmt_num(ds2, 1)} zile ({fmt_num(res.vit_stoc[2], 2)} rotații/an), creanțele totale în {fmt_num(dc2, 1)} zile, iar datoriile pe termen scurt în {fmt_num(dd2, 1)} zile. Pe relația comercială directă, {decalaj_com}; ciclul de conversie a numerarului este de {fmt_num(res.ciclu_conversie[2], 1)} zile.",
            f"Analiza vitezelor de rotație: în perioada de raportare 2 ({z[2]} de zile de activitate cumulată), durata de stocare este de {fmt_num(ds2, 1)} zile, durata de recuperare a creanțelor totale de {fmt_num(dc2, 1)} zile, iar durata datoriilor curente de {fmt_num(dd2, 1)} zile (față de {fmt_num(res.dur_dat[1], 1)} zile în perioada de raportare 1). La nivel comercial, {decalaj_com}.",
        ],
    )
    add_paragraph(
        doc,
        s18 + " Duratele folosesc soldurile de la finele perioadei; soldurile clienților și furnizorilor includ TVA, iar rulajele nu, ceea ce supraevaluează ușor duratele comerciale.",
    )
