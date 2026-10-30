#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 11: Diagnosticul Echilibrului Financiar Static.
Centralizează indicatorii FR, NFR și TN pe cele trei perioade de raportare (0, 1 și 2),
împreună cu diagnosticul calitativ (EFTL/DFTL, EFTS/NFCE, EFC/DFT).
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
    make_table,
    pick_variant,
)
from sheet_07_echilibrul_financiar import EchilibruFinanciarResult


@dataclass
class EchilibruStaticResult:
    fr: Tuple[float, float, float]
    nfr: Tuple[float, float, float]
    tn: Tuple[float, float, float]
    interp_fr: Tuple[str, str, str]
    interp_nfr: Tuple[str, str, str]
    interp_tn: Tuple[str, str, str]


def compute_sheet_11_echilibru_static(ef: EchilibruFinanciarResult) -> EchilibruStaticResult:
    return EchilibruStaticResult(
        fr=ef.fr_sus,
        nfr=ef.nfr,
        tn=ef.tn,
        interp_fr=ef.diag_fr,
        interp_nfr=ef.diag_nfr,
        interp_tn=ef.diag_tn,
    )


def render_sheet_11_echilibru_static(
    doc: Document, bal: BalantaData, res: EchilibruStaticResult
) -> None:
    """Generează secțiunea Echilibrului Financiar Static în documentul Word."""
    add_heading(doc, "11. Diagnosticul Echilibrului Financiar Static (FR, NFR, TN)", level=1)

    headers = [
        "Indicator de Echilibru Static",
        "Perioada de raportare 0 (lei)",
        "Perioada de raportare 1 (lei)",
        "Perioada de raportare 2 (lei)",
        "Diagnostic P0",
        "Diagnostic P1",
        "Diagnostic P2",
    ]
    rows = [
        [
            "FR (Fondul de rulment = CPERM - AI)",
            fmt_money(res.fr[0]),
            fmt_money(res.fr[1]),
            fmt_money(res.fr[2]),
            res.interp_fr[0],
            res.interp_fr[1],
            res.interp_fr[2],
        ],
        [
            "NFR (Necesarul de fond de rulment = AC fără disp. - DTS)",
            fmt_money(res.nfr[0]),
            fmt_money(res.nfr[1]),
            fmt_money(res.nfr[2]),
            res.interp_nfr[0],
            res.interp_nfr[1],
            res.interp_nfr[2],
        ],
        [
            "TN (Trezoreria netă = FR - NFR)",
            fmt_money(res.tn[0]),
            fmt_money(res.tn[1]),
            fmt_money(res.tn[2]),
            res.interp_tn[0],
            res.interp_tn[1],
            res.interp_tn[2],
        ],
    ]
    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[5.0, 2.3, 2.3, 2.3, 1.9, 1.9, 1.9],
        total_rows=[2],
        font_size=7.8,
    )

    all_eq_p2 = res.interp_fr[2] == "EFTL" and res.interp_nfr[2] == "EFTS" and res.interp_tn[2] == "EFC"
    if all_eq_p2:
        s11_msg = pick_variant(
            bal,
            "s11_all_ok",
            [
                f"Concluzie privind echilibrul static: În perioada de raportare 2, societatea respectă simultan toate cele trei condiții de echilibru structural — EFTL (FR = {fmt_money(res.fr[2])} lei >= 0), EFTS (FR >= NFR = {fmt_money(res.nfr[2])} lei) și EFC (TN = {fmt_money(res.tn[2])} lei >= 0) — demonstrând că marja de siguranță pe termen lung acoperă integral necesarul ciclului de exploatare și consolidează trezoreria.",
                f"Evaluarea statică a poziției financiare în perioada de raportare 2 confirmă triplul echilibru patrimonial (EFTL + EFTS + EFC): Fondul de rulment ({fmt_money(res.fr[2])} lei) devansează Necesarul de fond de rulment ({fmt_money(res.nfr[2])} lei), rezultând o Trezorerie netă excedentară de {fmt_money(res.tn[2])} lei.",
            ],
        )
    else:
        s11_msg = (
            f"Concluzie privind echilibrul static: În perioada de raportare 2, tabloul de încadrare indică {res.interp_fr[2]} pe termen lung (FR = {fmt_money(res.fr[2])} lei), {res.interp_nfr[2]} pe termen scurt (NFR = {fmt_money(res.nfr[2])} lei) și {res.interp_tn[2]} la nivelul trezoreriei curente (TN = {fmt_money(res.tn[2])} lei)."
        )
    add_paragraph(doc, s11_msg)
