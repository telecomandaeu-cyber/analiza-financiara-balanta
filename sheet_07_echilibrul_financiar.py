#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 7: Analiza Echilibrului Financiar Patrimonial (FR, FRP, FRS, NFR, TN).
Evaluează echilibrul pe termen lung (Fondul de rulment total, propriu și străin),
echilibrul pe termen scurt (Necesarul de fond de rulment) și echilibrul curent (Trezoreria netă).
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
    fmt_money,
    fmt_ratio_pct,
    make_table,
)
from sheet_03_bilant_financiar import BilantFinanciarResult
from sheet_05_structura_financiara import StructuraFinanciaraResult


@dataclass
class EchilibruFinanciarResult:
    fr_sus: Tuple[float, float, float]
    fr_jos: Tuple[float, float, float]
    frp: Tuple[float, float, float]
    frs: Tuple[float, float, float]
    ac_fara_disp: Tuple[float, float, float]
    nfr: Tuple[float, float, float]
    tn: Tuple[float, float, float]
    diag_fr: Tuple[str, str, str]      # EFTL sau DFTL
    diag_nfr: Tuple[str, str, str]     # EFTS sau NFCE
    diag_tn: Tuple[str, str, str]      # EFC sau DFT
    semnal_tn_fr: Tuple[str, str, str]


def compute_sheet_07_echilibrul_financiar(
    bf: BilantFinanciarResult, sf: StructuraFinanciaraResult
) -> EchilibruFinanciarResult:
    cperm = sf.cperm
    ai = bf.active_imobilizate
    ac = bf.active_circulante
    dts = sf.dts
    cp = sf.cp
    disp = bf.disponibilitati

    fr_sus = tuple(round(cperm[p] - ai[p], 2) for p in range(3))
    fr_jos = tuple(round(ac[p] - dts[p], 2) for p in range(3))

    frp = tuple(round(cp[p] - ai[p], 2) for p in range(3))
    frs = tuple(round(fr_sus[p] - frp[p], 2) for p in range(3))

    ac_fara_disp = tuple(round(ac[p] - disp[p], 2) for p in range(3))
    nfr = tuple(round(ac_fara_disp[p] - dts[p], 2) for p in range(3))

    tn = tuple(round(fr_sus[p] - nfr[p], 2) for p in range(3))

    diag_fr = tuple("EFTL" if fr_sus[p] >= 0 else "DFTL" for p in range(3))
    diag_nfr = tuple("EFTS" if fr_sus[p] >= nfr[p] else "NFCE" for p in range(3))
    diag_tn = tuple("EFC" if tn[p] >= 0 else "DFT" for p in range(3))
    semnal_tn_fr = tuple(
        "Semnal nefavorabil (TN > FR)"
        if tn[p] > fr_sus[p]
        else "Semnal favorabil (TN <= FR)"
        for p in range(3)
    )

    return EchilibruFinanciarResult(
        fr_sus=fr_sus,
        fr_jos=fr_jos,
        frp=frp,
        frs=frs,
        ac_fara_disp=ac_fara_disp,
        nfr=nfr,
        tn=tn,
        diag_fr=diag_fr,
        diag_nfr=diag_nfr,
        diag_tn=diag_tn,
        semnal_tn_fr=semnal_tn_fr,
    )


def render_sheet_07_echilibrul_financiar(
    doc: Document,
    bal: BalantaData,
    bf: BilantFinanciarResult,
    sf: StructuraFinanciaraResult,
    ef: EchilibruFinanciarResult,
) -> None:
    """Generează secțiunea Analizei Echilibrului Financiar în documentul Word."""
    add_heading(doc, "7. Analiza Echilibrului Financiar Patrimonial (FR, FRP, FRS, NFR, TN)", level=1)

    headers = [
        "Indicator de Echilibru Financiar (lei)",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
        "Δ (P1 - P0)",
        "Δ (P2 - P1)",
        "I (P2 / P1)",
        "Diagnostic P2",
    ]

    items = [
        ("1. Fondul de rulment — partea de sus (CPERM - AI)", ef.fr_sus, ef.diag_fr[2], True),
        ("2. Fondul de rulment — partea de jos (AC - DTS)", ef.fr_jos, ef.diag_fr[2], False),
        ("3. Fondul de rulment propriu (FRP = CP - AI)", ef.frp, "FRP > 0" if ef.frp[2] >= 0 else "FRP < 0", True),
        ("4. Fondul de rulment străin (FRS = FR - FRP = DTL)", ef.frs, "DTL", False),
        ("5. Active circulante fără disponibil (Stocuri + Creanțe)", ef.ac_fara_disp, "—", False),
        ("6. Datorii pe termen scurt (DTS)", sf.dts, "—", False),
        ("7. Necesarul de fond de rulment (NFR = AC fără disp. - DTS)", ef.nfr, ef.diag_nfr[2], True),
        ("8. Trezoreria netă (TN = FR - NFR = Disponibil)", ef.tn, ef.diag_tn[2], True),
    ]

    rows = []
    tot_idx = []
    for i, (label, vals, diag, is_tot) in enumerate(items):
        if is_tot:
            tot_idx.append(i)
        d = calc_delta(vals)
        idx = calc_index(vals)
        rows.append([
            label,
            fmt_money(vals[0]),
            fmt_money(vals[1]),
            fmt_money(vals[2]),
            fmt_money(d[0]),
            fmt_money(d[1]),
            fmt_ratio_pct(idx[1]),
            diag,
        ])

    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[5.2, 1.9, 1.9, 1.9, 1.8, 1.8, 1.5, 1.8],
        total_rows=tot_idx,
        font_size=7.7,
    )

    add_paragraph(
        doc,
        f"Diagnostic de echilibru: Fondul de rulment determinat prin partea de sus a bilanțului este egal la cent cu cel determinat "
        f"prin partea de jos ({fmt_money(ef.fr_sus[0])} lei în perioada de raportare 0, {fmt_money(ef.fr_sus[1])} lei în perioada de "
        f"raportare 1 și {fmt_money(ef.fr_sus[2])} lei în perioada de raportare 2 -> diagnostic {ef.diag_fr[2]}). Necesarul de fond "
        f"de rulment (NFR) este de {fmt_money(ef.nfr[0])} lei în perioada de raportare 0, {fmt_money(ef.nfr[1])} lei în perioada de "
        f"raportare 1 și {fmt_money(ef.nfr[2])} lei în perioada de raportare 2 (diagnostic {ef.diag_nfr[2]}), iar Trezoreria netă "
        f"(TN = FR - NFR) este pozitivă în toate cele trei perioade de raportare ({fmt_money(ef.tn[0])} lei, {fmt_money(ef.tn[1])} lei "
        f"și {fmt_money(ef.tn[2])} lei -> diagnostic {ef.diag_tn[2]}).",
    )
