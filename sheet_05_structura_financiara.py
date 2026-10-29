#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 5: Analiza Structurii Financiare a Pasivului.
Analizează arhitectura surselor de finanțare prin optica apartenenței fondurilor
(Total Pasiv 1 = Capitaluri proprii + Datorii totale) și optica stabilității în timp
(Total Pasiv 2 = Capitaluri permanente + Datorii pe termen scurt).
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
    fmt_pct,
    fmt_ratio_pct,
    make_table,
    safe_div,
)
from sheet_03_bilant_financiar import BilantFinanciarResult


@dataclass
class StructuraFinanciaraResult:
    cp: Tuple[float, float, float]
    dtl: Tuple[float, float, float]
    dts: Tuple[float, float, float]
    cperm: Tuple[float, float, float]
    dtot: Tuple[float, float, float]
    pasiv1: Tuple[float, float, float]
    pasiv2: Tuple[float, float, float]


def compute_sheet_05_structura_financiara(bf: BilantFinanciarResult) -> StructuraFinanciaraResult:
    cp = bf.capitaluri_proprii
    dtl = bf.datorii_termen_lung
    dts = bf.datorii_termen_scurt
    cperm = bf.capitaluri_permanente
    dtot = tuple(round(dtl[p] + dts[p], 2) for p in range(3))
    pasiv1 = tuple(round(cp[p] + dtot[p], 2) for p in range(3))
    pasiv2 = tuple(round(cperm[p] + dts[p], 2) for p in range(3))
    return StructuraFinanciaraResult(
        cp=cp,
        dtl=dtl,
        dts=dts,
        cperm=cperm,
        dtot=dtot,
        pasiv1=pasiv1,
        pasiv2=pasiv2,
    )


def render_sheet_05_structura_financiara(
    doc: Document, bal: BalantaData, sf: StructuraFinanciaraResult
) -> None:
    """Generează secțiunea Structurii Financiare a Pasivului în documentul Word."""
    add_heading(doc, "5. Analiza Structurii Financiare a Pasivului", level=1)

    items = [
        ("1", "Capitaluri proprii (CP)", sf.cp, False),
        ("2", "Datorii pe termen lung (DTL)", sf.dtl, False),
        ("3", "Datorii pe termen scurt (DTS)", sf.dts, False),
        ("4", "Capitaluri permanente (CPERM = CP + DTL)", sf.cperm, True),
        ("5", "Datorii totale (DT = DTL + DTS)", sf.dtot, True),
        ("6", "Total Pasiv 1 (Capitaluri proprii + Datorii totale)", sf.pasiv1, True),
        ("7", "Total Pasiv 2 (Capitaluri permanente + Datorii pe termen scurt)", sf.pasiv2, True),
    ]

    headers = [
        "Nr.",
        "Agregat de Structură Financiară (lei)",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
        "Δ (P1 - P0)",
        "Δ (P2 - P1)",
        "I (P1 / P0)",
        "I (P2 / P1)",
    ]
    rows = []
    tot_idx = []
    for i, (nr, name, vals, is_tot) in enumerate(items):
        if is_tot:
            tot_idx.append(i)
        d = calc_delta(vals)
        idx = calc_index(vals)
        rows.append([
            nr,
            name,
            fmt_money(vals[0]),
            fmt_money(vals[1]),
            fmt_money(vals[2]),
            fmt_money(d[0]),
            fmt_money(d[1]),
            fmt_ratio_pct(idx[0]),
            fmt_ratio_pct(idx[1]),
        ])

    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[0.7, 5.1, 1.9, 1.9, 1.9, 1.7, 1.7, 1.45, 1.45],
        total_rows=tot_idx,
        font_size=7.7,
    )

    add_paragraph(
        doc,
        f"Sinteză financiară: Între perioada de raportare 0 și perioada de raportare 2, datoriile pe termen lung înregistrează o "
        f"variație de {fmt_money(sf.dtl[2] - sf.dtl[0])} lei ({fmt_pct(safe_div(sf.dtl[2] - sf.dtl[0], sf.dtl[0])*100, plus_sign=True)}), "
        f"datoriile pe termen scurt variază cu {fmt_money(sf.dts[2] - sf.dts[0])} lei "
        f"({fmt_pct(safe_div(sf.dts[2] - sf.dts[0], sf.dts[0])*100, plus_sign=True)}), iar capitalurile proprii evoluează cu "
        f"{fmt_money(sf.cp[2] - sf.cp[0])} lei ({fmt_pct(safe_div(sf.cp[2] - sf.cp[0], sf.cp[0])*100, plus_sign=True)}, de la "
        f"{fmt_money(sf.cp[0])} lei la {fmt_money(sf.cp[2])} lei). Ponderea capitalurilor proprii în pasivul total este de "
        f"{fmt_ratio_pct(safe_div(sf.cp[0], sf.pasiv1[0]))} în perioada de raportare 0, {fmt_ratio_pct(safe_div(sf.cp[1], sf.pasiv1[1]))} "
        f"în perioada de raportare 1 și {fmt_ratio_pct(safe_div(sf.cp[2], sf.pasiv1[2]))} în perioada de raportare 2.",
    )
