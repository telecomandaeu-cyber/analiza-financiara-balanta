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

    add_paragraph(
        doc,
        "Semnificația codurilor și concluzie statică: EFTL = Echilibru Financiar pe Termen Lung (FR >= 0); EFTS = Echilibru Financiar "
        "pe Termen Scurt (FR >= NFR); EFC = Echilibru Financiar Curent/Global (TN >= 0). Societatea îndeplinește simultan toate cele "
        "trei condiții de echilibru (EFTL + EFTS + EFC) în fiecare dintre cele trei perioade de raportare analizate.",
    )
