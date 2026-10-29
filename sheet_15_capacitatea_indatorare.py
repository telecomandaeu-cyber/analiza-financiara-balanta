#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 15: Analiza Capacității de Îndatorare Bancară.
Evaluează cei patru coeficienți bancari uzuali (CIG 1 <= 0,66; CIG 2 <= 2,00; CIT 1 <= 0,50; CIT 2 <= 1,00),
atât în varianta Bilanțului Financiar (cu veniturile în avans ct. 472+475 incluse), cât și pe datoriile efective.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Tuple

from docx import Document

from core_balanta import (
    BalantaData,
    add_heading,
    add_paragraph,
    fmt_num,
    fmt_ratio_pct,
    make_table,
    safe_div,
)
from sheet_01_bilant import BilantResult
from sheet_03_bilant_financiar import BilantFinanciarResult
from sheet_05_structura_financiara import StructuraFinanciaraResult


@dataclass
class CapacitateIndatorareResult:
    cig1_fin: Tuple[float, float, float]
    cig2_fin: Tuple[float, float, float]
    cit1_fin: Tuple[float, float, float]
    cit2_fin: Tuple[float, float, float]
    verdict_fin: Tuple[str, str, str]

    cig1_ef: Tuple[float, float, float]
    cig2_ef: Tuple[float, float, float]
    cit1_ef: Tuple[float, float, float]
    cit2_ef: Tuple[float, float, float]
    verdict_ef: Tuple[str, str, str]


def compute_sheet_15_capacitatea_indatorare(
    bil: BilantResult, bf: BilantFinanciarResult, sf: StructuraFinanciaraResult
) -> CapacitateIndatorareResult:
    cig1_f = tuple(safe_div(sf.dtot[p], sf.pasiv1[p]) for p in range(3))
    cig2_f = tuple(safe_div(sf.dtot[p], sf.cp[p]) for p in range(3))
    cit1_f = tuple(safe_div(sf.dtl[p], sf.cperm[p]) for p in range(3))
    cit2_f = tuple(safe_div(sf.dtl[p], sf.cp[p]) for p in range(3))

    v_fin = []
    for p, per_lbl in enumerate(("perioada de raportare 0", "perioada de raportare 1", "perioada de raportare 2")):
        ok = (cig1_f[p] <= 0.66) and (cig2_f[p] <= 2.0) and (cit1_f[p] <= 0.50) and (cit2_f[p] <= 1.0)
        v_fin.append(
            f"În {per_lbl}, întreprinderea {'ARE' if ok else 'NU ARE'} capacitate de îndatorare în optica brută (cu venituri în avans)"
        )

    dt_ef = tuple(round(bil.rd(11)[p] + bil.rd(14)[p], 2) for p in range(3))
    cig1_e = tuple(safe_div(dt_ef[p], sf.pasiv1[p]) for p in range(3))
    cig2_e = tuple(safe_div(dt_ef[p], sf.cp[p]) for p in range(3))
    cit1_e = cit1_f
    cit2_e = cit2_f

    v_ef = []
    for p, per_lbl in enumerate(("perioada de raportare 0", "perioada de raportare 1", "perioada de raportare 2")):
        ok = (cig1_e[p] <= 0.66) and (cig2_e[p] <= 2.0) and (cit1_e[p] <= 0.50) and (cit2_e[p] <= 1.0)
        v_ef.append(
            f"În {per_lbl}, întreprinderea {'ARE' if ok else 'NU ARE'} capacitate deplină de îndatorare pe datoriile efective"
        )

    return CapacitateIndatorareResult(
        cig1_fin=cig1_f,
        cig2_fin=cig2_f,
        cit1_fin=cit1_f,
        cit2_fin=cit2_f,
        verdict_fin=(v_fin[0], v_fin[1], v_fin[2]),
        cig1_ef=cig1_e,
        cig2_ef=cig2_e,
        cit1_ef=cit1_e,
        cit2_ef=cit2_e,
        verdict_ef=(v_ef[0], v_ef[1], v_ef[2]),
    )


def render_sheet_15_capacitatea_indatorare(
    doc: Document, bal: BalantaData, res: CapacitateIndatorareResult
) -> None:
    """Generează secțiunea Capacității de Îndatorare Bancară într-un tabel unificat."""
    add_heading(doc, "15. Analiza Capacității de Îndatorare Bancară (CIG 1, CIG 2, CIT 1, CIT 2)", level=1)

    headers = [
        "Coeficient și Plafon Bancar",
        "Relația de Calcul",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
        "Normă P0",
        "Normă P1",
        "Normă P2",
    ]
    rows = [
        [
            "CIG 1 <= 0,66 (Bilanț Financiar)",
            "Datorii totale / Pasiv total",
            fmt_num(res.cig1_fin[0], 4),
            fmt_num(res.cig1_fin[1], 4),
            fmt_num(res.cig1_fin[2], 4),
            "Da" if res.cig1_fin[0] <= 0.66 else "Nu",
            "Da" if res.cig1_fin[1] <= 0.66 else "Nu",
            "Da" if res.cig1_fin[2] <= 0.66 else "Nu",
        ],
        [
            "CIG 1 efectiv <= 0,66 (fără ct. 472+475)",
            "Datorii efective / Pasiv total",
            fmt_num(res.cig1_ef[0], 4),
            fmt_num(res.cig1_ef[1], 4),
            fmt_num(res.cig1_ef[2], 4),
            "Da" if res.cig1_ef[0] <= 0.66 else "Nu",
            "Da" if res.cig1_ef[1] <= 0.66 else "Nu",
            "Da" if res.cig1_ef[2] <= 0.66 else "Nu",
        ],
        [
            "CIG 2 <= 2,00 (Bilanț Financiar)",
            "Datorii totale / Capitaluri proprii",
            fmt_num(res.cig2_fin[0], 4),
            fmt_num(res.cig2_fin[1], 4),
            fmt_num(res.cig2_fin[2], 4),
            "Da" if res.cig2_fin[0] <= 2.0 else "Nu",
            "Da" if res.cig2_fin[1] <= 2.0 else "Nu",
            "Da" if res.cig2_fin[2] <= 2.0 else "Nu",
        ],
        [
            "CIG 2 efectiv <= 2,00 (fără ct. 472+475)",
            "Datorii efective / Capitaluri proprii",
            fmt_num(res.cig2_ef[0], 4),
            fmt_num(res.cig2_ef[1], 4),
            fmt_num(res.cig2_ef[2], 4),
            "Da" if res.cig2_ef[0] <= 2.0 else "Nu",
            "Da" if res.cig2_ef[1] <= 2.0 else "Nu",
            "Da" if res.cig2_ef[2] <= 2.0 else "Nu",
        ],
        [
            "CIT 1 <= 0,50 (Îndatorare la termen)",
            "Datorii pe termen lung / Capitaluri permanente",
            fmt_num(res.cit1_fin[0], 4),
            fmt_num(res.cit1_fin[1], 4),
            fmt_num(res.cit1_fin[2], 4),
            "Da" if res.cit1_fin[0] <= 0.50 else "Nu",
            "Da" if res.cit1_fin[1] <= 0.50 else "Nu",
            "Da" if res.cit1_fin[2] <= 0.50 else "Nu",
        ],
        [
            "CIT 2 <= 1,00 (Îndatorare la termen)",
            "Datorii pe termen lung / Capitaluri proprii",
            fmt_num(res.cit2_fin[0], 4),
            fmt_num(res.cit2_fin[1], 4),
            fmt_num(res.cit2_fin[2], 4),
            "Da" if res.cit2_fin[0] <= 1.0 else "Nu",
            "Da" if res.cit2_fin[1] <= 1.0 else "Nu",
            "Da" if res.cit2_fin[2] <= 1.0 else "Nu",
        ],
    ]

    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[4.2, 4.2, 1.9, 1.9, 1.9, 1.2, 1.2, 1.2],
        total_rows=[1, 3, 4, 5],
        font_size=7.7,
    )

    add_paragraph(
        doc,
        f"Verdict bancar: În perioada de raportare 2, coeficienții de îndatorare la termen înregistrează nivelurile CIT 1 = "
        f"{fmt_num(res.cit1_fin[2], 4)} (plafon maxim 0,50) și CIT 2 = {fmt_num(res.cit2_fin[2], 4)} (plafon maxim 1,00), iar "
        f"coeficienții îndatorării globale efective sunt CIG 1 = {fmt_num(res.cig1_ef[2], 4)} ({fmt_ratio_pct(res.cig1_ef[2])}, față "
        f"de plafonul maxim de 66%) și CIG 2 = {fmt_num(res.cig2_ef[2], 4)} (față de plafonul maxim de 2,00), reflectând "
        f"proporția în care datoriile totale sunt acoperite de pasivul total și de capitalurile proprii.",
    )
