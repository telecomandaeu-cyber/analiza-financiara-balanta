#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 6: Analiza Ratelor de Structură ale Pasivului (Autonomie și Îndatorare).
Evaluează cele cinci rate fundamentale de structură a pasivului: RSF, RIG, RAFG, RIT și RAFT.
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
    fmt_ratio_pct,
    make_table,
    pick_variant,
    safe_div,
)
from sheet_05_structura_financiara import StructuraFinanciaraResult


@dataclass
class LiabilityRateBlock:
    title: str
    num_name: str
    den_name: str
    rate_name: str
    norm_label: str
    num_vals: Tuple[float, float, float]
    den_vals: Tuple[float, float, float]
    rate_vals: Tuple[float, float, float]


@dataclass
class RatePasivResult:
    rsf: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    rig: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    rafg: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    rit: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    raft: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    blocks: List[LiabilityRateBlock] = field(default_factory=list)


def compute_sheet_06_ratele_pasivului(sf: StructuraFinanciaraResult) -> RatePasivResult:
    res = RatePasivResult()
    tp = sf.pasiv1
    res.rsf = tuple(safe_div(sf.cperm[p], tp[p]) for p in range(3))
    res.rig = tuple(safe_div(sf.dtot[p], tp[p]) for p in range(3))
    res.rafg = tuple(safe_div(sf.cp[p], tp[p]) for p in range(3))
    res.rit = tuple(safe_div(sf.dtl[p], sf.dtot[p]) for p in range(3))
    res.raft = tuple(safe_div(sf.cp[p], sf.cperm[p]) for p in range(3))

    res.blocks = [
        LiabilityRateBlock(
            "Rata stabilității financiare (RSF)",
            "Capitaluri permanente",
            "Total Pasiv",
            "1. Rata stabilității financiare (RSF = CPERM / Total Pasiv)",
            "> 50% (optim > 66%)",
            sf.cperm,
            tp,
            res.rsf,
        ),
        LiabilityRateBlock(
            "Rata îndatorării globale (RIG)",
            "Datorii totale",
            "Total Pasiv",
            "2. Rata îndatorării globale (RIG = DT / Total Pasiv)",
            "<= 66% (maxim 2/3)",
            sf.dtot,
            tp,
            res.rig,
        ),
        LiabilityRateBlock(
            "Rata autonomiei financiare globale (RAFG)",
            "Capitaluri proprii",
            "Total Pasiv",
            "3. Rata autonomiei financiare globale (RAFG = CP / Total Pasiv)",
            ">= 33% (optim >= 50%)",
            sf.cp,
            tp,
            res.rafg,
        ),
        LiabilityRateBlock(
            "Rata îndatorării la termen (RIT)",
            "Datorii pe termen lung",
            "Datorii totale",
            "4. Rata îndatorării la termen (RIT = DTL / Datorii totale)",
            "<= 50%",
            sf.dtl,
            sf.dtot,
            res.rit,
        ),
        LiabilityRateBlock(
            "Rata autonomiei financiare la termen (RAFT)",
            "Capitaluri proprii",
            "Capitaluri permanente",
            "5. Rata autonomiei financiare la termen (RAFT = CP / CPERM)",
            ">= 50%",
            sf.cp,
            sf.cperm,
            res.raft,
        ),
    ]
    return res


def render_sheet_06_ratele_pasivului(doc: Document, bal: BalantaData, res: RatePasivResult) -> None:
    """Generează secțiunea Ratelor de Structură ale Pasivului în documentul Word."""
    add_heading(doc, "6. Analiza Ratelor de Structură ale Pasivului (Autonomie și Îndatorare)", level=1)

    headers = [
        "Rată de Structură a Pasivului",
        "Plafon Bancar",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
        "Δ (P1 - P0)",
        "Δ (P2 - P1)",
        "I (P2 / P1)",
    ]
    rows = []
    for b in res.blocks:
        d = calc_delta(b.rate_vals)
        i = calc_index(b.rate_vals)
        rows.append([
            b.rate_name,
            b.norm_label,
            fmt_ratio_pct(b.rate_vals[0]),
            fmt_ratio_pct(b.rate_vals[1]),
            fmt_ratio_pct(b.rate_vals[2]),
            fmt_ratio_pct(d[0], plus_sign=True),
            fmt_ratio_pct(d[1], plus_sign=True),
            fmt_ratio_pct(i[1]),
        ])

    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[5.4, 2.2, 1.9, 1.9, 1.9, 1.5, 1.5, 1.5],
        total_rows=[0, 2, 4],
        font_size=7.8,
    )

    traj_rsf = describe_trajectory(res.rsf, bal, "s06_rsf", "rata stabilității financiare (RSF)", is_ratio=True)
    if res.rsf[2] >= 0.66 and res.rafg[2] >= 0.50:
        verdict_pasiv = pick_variant(
            bal,
            "s06_v_strong",
            [
                f"În perioada de raportare 2, atât rata autonomiei financiare globale (RAFG = {fmt_ratio_pct(res.rafg[2])}), cât și rata autonomiei la termen (RAFT = {fmt_ratio_pct(res.raft[2])}) depășesc pragul bancar de 50%, în timp ce rata îndatorării globale (RIG = {fmt_ratio_pct(res.rig[2])}) și la termen (RIT = {fmt_ratio_pct(res.rit[2])}) confirmă un risc financiar redus.",
                f"Structura pasivului din perioada de raportare 2 indică o autonomie financiară ridicată: RAFG atinge {fmt_ratio_pct(res.rafg[2])}, RAFT se situează la {fmt_ratio_pct(res.raft[2])}, iar îndatorarea globală (RIG = {fmt_ratio_pct(res.rig[2])}) se menține confortabil sub plafonul critic de 66%.",
            ],
        )
    elif res.rsf[2] >= 0.50:
        verdict_pasiv = pick_variant(
            bal,
            "s06_v_med",
            [
                f"În perioada de raportare 2, resursele stabile depășesc jumătate din pasiv, rata autonomiei financiare globale (RAFG) fiind de {fmt_ratio_pct(res.rafg[2])}, rata autonomiei la termen (RAFT) de {fmt_ratio_pct(res.raft[2])}, iar rata îndatorării globale (RIG) de {fmt_ratio_pct(res.rig[2])}.",
                f"La finele perioadei de raportare 2, echilibrul dintre sursele proprii și cele împrumutate este reflectat de RAFG = {fmt_ratio_pct(res.rafg[2])}, RAFT = {fmt_ratio_pct(res.raft[2])}, RIT = {fmt_ratio_pct(res.rit[2])} și RIG = {fmt_ratio_pct(res.rig[2])}.",
            ],
        )
    else:
        verdict_pasiv = (
            f"În perioada de raportare 2, ponderea ridicată a datoriilor pe termen scurt determină o rată a îndatorării globale (RIG) de {fmt_ratio_pct(res.rig[2])} și o rată a autonomiei financiare globale (RAFG) de {fmt_ratio_pct(res.rafg[2])} (RAFT = {fmt_ratio_pct(res.raft[2])}, RIT = {fmt_ratio_pct(res.rit[2])})."
        )

    add_paragraph(doc, f"Evaluarea ratelor de structură ale pasivului: Pe intervalul analizat, {traj_rsf}. {verdict_pasiv}")
