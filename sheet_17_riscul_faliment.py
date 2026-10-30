#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 17: Evaluarea Riscului de Insolvență (Modelele Altman și Conan-Holder).
Determină funcțiile-scor Altman Z și Conan-Holder Z pe cele trei perioade de raportare (0, 1 și 2),
atât în varianta standard, cât și în varianta ajustată economic (fără veniturile în avans ct. 472+475
în datorii și cu veniturile din proiecte ct. 758 incluse în EBE și VA).
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
    make_table,
    pick_variant,
    safe_div,
)
from sheet_01_bilant import BilantResult
from sheet_02_cpp import CPPResult
from sheet_03_bilant_financiar import BilantFinanciarResult
from sheet_05_structura_financiara import StructuraFinanciaraResult
from sheet_07_echilibrul_financiar import EchilibruFinanciarResult
from sheet_13_sig import SIGResult


@dataclass
class RiscFalimentResult:
    alt_v1: Tuple[float, float, float]
    alt_v2: Tuple[float, float, float]
    alt_v3: Tuple[float, float, float]
    alt_v4: Tuple[float, float, float]
    alt_v5: Tuple[float, float, float]
    z_altman: Tuple[float, float, float]
    interp_altman: Tuple[str, str, str]
    z_altman_ajustat: Tuple[float, float, float]

    ch_v1: Tuple[float, float, float]
    ch_v2: Tuple[float, float, float]
    ch_v3: Tuple[float, float, float]
    ch_v4: Tuple[float, float, float]
    ch_v5: Tuple[float, float, float]
    z_canon_holder: Tuple[float, float, float]
    interp_canon_holder: Tuple[str, str, str]
    z_ch_ajustat: Tuple[float, float, float]


def _interp_altman(z: float) -> str:
    if z < 1.80:
        return "Faliment iminent (Z < 1,80)"
    if z <= 3.00:
        return "Dificultate financiară temporară (1,80 <= Z <= 3,00)"
    return "Societate solvabilă / Fără risc (Z > 3,00)"


def _interp_ch(z: float) -> str:
    if z < 4.0:
        return "Faliment iminent / Risc ridicat (Z < 4)"
    if z <= 9.0:
        return "Zonă de incertitudine / prudență (4 <= Z <= 9)"
    return "Situație financiară bună / Risc redus (Z > 9)"


def compute_sheet_17_riscul_faliment(
    bil: BilantResult,
    cpp: CPPResult,
    bf: BilantFinanciarResult,
    sf: StructuraFinanciaraResult,
    ef: EchilibruFinanciarResult,
    sig: SIGResult,
) -> RiscFalimentResult:
    ta = bf.total_activ
    pb = cpp.rezultat_brut
    ca = cpp.cifra_afaceri
    cp = sf.cp
    dt = sf.dtot
    fr = ef.fr_sus

    v1_a = tuple(safe_div(pb[p], ta[p]) for p in range(3))
    v2_a = tuple(safe_div(ca[p], ta[p]) for p in range(3))
    v3_a = tuple(safe_div(cp[p], dt[p]) for p in range(3))
    v4_a = tuple(safe_div(fr[p], ta[p]) for p in range(3))
    v5_a = v1_a
    z_a = tuple(
        3.3 * v1_a[p] + 1.0 * v2_a[p] + 0.6 * v3_a[p] + 1.2 * v4_a[p] + 1.4 * v5_a[p]
        for p in range(3)
    )
    i_a = tuple(_interp_altman(z_a[p]) for p in range(3))

    dt_ef = tuple(round(bil.rd(11)[p] + bil.rd(14)[p], 2) for p in range(3))
    v3_a_ef = tuple(safe_div(cp[p], dt_ef[p]) for p in range(3))
    z_a_ef = tuple(
        3.3 * v1_a[p] + 1.0 * v2_a[p] + 0.6 * v3_a_ef[p] + 1.2 * v4_a[p] + 1.4 * v5_a[p]
        for p in range(3)
    )

    cr_inv_disp = tuple(round(bf.creante[p] + bf.investitii_ts[p] + bf.disponibilitati[p], 2) for p in range(3))
    dts = sf.dts
    cperm = sf.cperm
    tp = sf.pasiv1
    ch_fin = cpp.cheltuieli_financiare
    ch_pers = cpp.rd(18)
    va = sig.valoarea_adaugata
    ebe = sig.ebe

    v1_c = tuple(safe_div(cr_inv_disp[p], dts[p]) for p in range(3))
    v2_c = tuple(safe_div(cperm[p], tp[p]) for p in range(3))
    v3_c = tuple(safe_div(ch_fin[p], ca[p]) for p in range(3))
    v4_c = tuple(safe_div(ch_pers[p], va[p]) for p in range(3))
    v5_c = tuple(safe_div(ebe[p], va[p]) for p in range(3))
    z_c = tuple(
        16.0 * v1_c[p] + 22.0 * v2_c[p] - 87.0 * v3_c[p] - 10.0 * v4_c[p] + 24.0 * v5_c[p]
        for p in range(3)
    )
    i_c = tuple(_interp_ch(z_c[p]) for p in range(3))

    z_c_aj = []
    for p in range(3):
        va_aj = va[p] + cpp.rd(10)[p]
        ebe_aj = ebe[p] + cpp.rd(10)[p] - cpp.rd(30)[p]
        v1_aj = safe_div(cr_inv_disp[p], bil.rd(11)[p])
        v2_aj = safe_div(cperm[p], cperm[p] + bil.rd(11)[p])
        v3_aj = safe_div(ch_fin[p], ca[p] + cpp.rd(10)[p])
        v4_aj = safe_div(ch_pers[p], va_aj)
        v5_aj = safe_div(ebe_aj, va_aj)
        z_c_aj.append(16.0 * v1_aj + 22.0 * v2_aj - 87.0 * v3_aj - 10.0 * v4_aj + 24.0 * v5_aj)

    return RiscFalimentResult(
        alt_v1=v1_a,
        alt_v2=v2_a,
        alt_v3=v3_a,
        alt_v4=v4_a,
        alt_v5=v5_a,
        z_altman=z_a,
        interp_altman=i_a,
        z_altman_ajustat=(z_a_ef[0], z_a_ef[1], z_a_ef[2]),
        ch_v1=v1_c,
        ch_v2=v2_c,
        ch_v3=v3_c,
        ch_v4=v4_c,
        ch_v5=v5_c,
        z_canon_holder=z_c,
        interp_canon_holder=i_c,
        z_ch_ajustat=(z_c_aj[0], z_c_aj[1], z_c_aj[2]),
    )


def render_sheet_17_riscul_faliment(
    doc: Document, bal: BalantaData, res: RiscFalimentResult
) -> None:
    """Generează secțiunea Evaluării Riscului de Insolvență într-un tabel sintetic unificat."""
    add_heading(doc, "17. Evaluarea Riscului de Insolvență (Modelele de Scoring Altman și Conan-Holder)", level=1)

    headers = [
        "Model / Variabilă a Funcției-Scor",
        "Perioada de raportare 0",
        "Perioada de raportare 1",
        "Perioada de raportare 2",
        "Încadrare Perioada de raportare 2",
    ]
    rows = [
        ["I. MODELUL ALTMAN: V1 (Profit brut / Active totale) & V5", fmt_num(res.alt_v1[0], 4), fmt_num(res.alt_v1[1], 4), fmt_num(res.alt_v1[2], 4), "Pondere 3,3 + 1,4"],
        ["   V2 (Cifra de afaceri / Active totale)  |  V4 (FR / Active totale)", f"{fmt_num(res.alt_v2[0], 4)} | {fmt_num(res.alt_v4[0], 4)}", f"{fmt_num(res.alt_v2[1], 4)} | {fmt_num(res.alt_v4[1], 4)}", f"{fmt_num(res.alt_v2[2], 4)} | {fmt_num(res.alt_v4[2], 4)}", "Pondere 1,0 | 1,2"],
        ["   V3 (Capitaluri proprii / Datorii totale Bilanț Financiar)", fmt_num(res.alt_v3[0], 4), fmt_num(res.alt_v3[1], 4), fmt_num(res.alt_v3[2], 4), "Pondere 0,6"],
        ["➔ Scorul Z Altman (Standard Bilanț Financiar)", fmt_num(res.z_altman[0], 4), fmt_num(res.z_altman[1], 4), fmt_num(res.z_altman[2], 4), res.interp_altman[2]],
        ["➔ Scorul Z Altman Ajustat (fără veniturile în avans ct. 472+475)", fmt_num(res.z_altman_ajustat[0], 4), fmt_num(res.z_altman_ajustat[1], 4), fmt_num(res.z_altman_ajustat[2], 4), _interp_altman(res.z_altman_ajustat[2])],
        ["II. MODELUL CONAN-HOLDER: V1 (Lichiditate rapidă) | V2 (RSF)", f"{fmt_num(res.ch_v1[0], 4)} | {fmt_num(res.ch_v2[0], 4)}", f"{fmt_num(res.ch_v1[1], 4)} | {fmt_num(res.ch_v2[1], 4)}", f"{fmt_num(res.ch_v1[2], 4)} | {fmt_num(res.ch_v2[2], 4)}", "Pondere +16 | +22"],
        ["   V3 (Ch. fin./CA) | V4 (Ch. pers./VA) | V5 (EBE/VA)", f"{fmt_num(res.ch_v3[0], 3)} | {fmt_num(res.ch_v4[0], 2)} | {fmt_num(res.ch_v5[0], 2)}", f"{fmt_num(res.ch_v3[1], 3)} | {fmt_num(res.ch_v4[1], 2)} | {fmt_num(res.ch_v5[1], 2)}", f"{fmt_num(res.ch_v3[2], 3)} | {fmt_num(res.ch_v4[2], 2)} | {fmt_num(res.ch_v5[2], 2)}", "Pondere -87 | -10 | +24"],
        ["➔ Scorul Z Conan-Holder (Standard, doar pe CA din ct. 704)", fmt_num(res.z_canon_holder[0], 4), fmt_num(res.z_canon_holder[1], 4), fmt_num(res.z_canon_holder[2], 4), res.interp_canon_holder[2]],
        ["➔ Scorul Z Conan-Holder Ajustat (cu veniturile din proiecte ct. 758)", fmt_num(res.z_ch_ajustat[0], 4), fmt_num(res.z_ch_ajustat[1], 4), fmt_num(res.z_ch_ajustat[2], 4), _interp_ch(res.z_ch_ajustat[2])],
    ]

    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[6.4, 2.3, 2.3, 2.3, 4.5],
        total_rows=[3, 4, 7, 8],
        font_size=7.7,
    )

    z_a2 = res.z_altman_ajustat[2]
    z_ch2 = res.z_ch_ajustat[2]
    if z_a2 > 3.0:
        s17_msg = pick_variant(
            bal,
            "s17_solv",
            [
                f"Evaluarea riscului de insolvență: Scorul Z Altman evoluează de la {fmt_num(res.z_altman[0], 4)} în perioada de raportare 0 la {fmt_num(res.z_altman[1], 4)} în perioada de raportare 1 și {fmt_num(res.z_altman[2], 4)} în perioada de raportare 2 (respectiv {fmt_num(z_a2, 4)} în varianta ajustată -> {_interp_altman(z_a2)}), plasând societatea în zona de siguranță financiară (> 3,00). Concomitent, funcția-scor Conan-Holder înregistrează {fmt_num(res.z_canon_holder[2], 4)} în varianta standard și {fmt_num(z_ch2, 4)} în varianta ajustată ({_interp_ch(z_ch2)}).",
                f"Diagnostic prin metoda scoringului: La finele perioadei de raportare 2, modelul Altman indică un scor standard de {fmt_num(res.z_altman[2], 4)} și un scor ajustat de {fmt_num(z_a2, 4)} (> 3,00 — {_interp_altman(z_a2)}), confirmat și de modelul Conan-Holder (scor standard {fmt_num(res.z_canon_holder[2], 4)}, scor ajustat {fmt_num(z_ch2, 4)} — {_interp_ch(z_ch2)}), ceea ce atestă un risc de faliment neglijabil.",
            ],
        )
    elif z_a2 >= 1.8:
        s17_msg = pick_variant(
            bal,
            "s17_med",
            [
                f"Evaluarea riscului de insolvență: În perioada de raportare 2, Scorul Z Altman se situează în intervalul [1,80 – 3,00] ({fmt_num(res.z_altman[2], 4)} în varianta standard și {fmt_num(z_a2, 4)} în varianta ajustată -> {_interp_altman(z_a2)}), în timp ce funcția Conan-Holder indică {fmt_num(res.z_canon_holder[2], 4)} (ajustat {fmt_num(z_ch2, 4)} -> {_interp_ch(z_ch2)}).",
                f"Diagnostic de risc: Pe parcursul celor trei perioade de raportare, Scorul Z Altman evoluează de la {fmt_num(res.z_altman[0], 4)} la {fmt_num(res.z_altman[1], 4)} și {fmt_num(res.z_altman[2], 4)} (ajustat {fmt_num(z_a2, 4)} — {_interp_altman(z_a2)}), iar scorul Conan-Holder în perioada de raportare 2 este de {fmt_num(res.z_canon_holder[2], 4)} (ajustat {fmt_num(z_ch2, 4)} — {_interp_ch(z_ch2)}).",
            ],
        )
    else:
        s17_msg = (
            f"Evaluarea riscului de insolvență: În perioada de raportare 2, Scorul Z Altman este de {fmt_num(res.z_altman[2], 4)} (ajustat {fmt_num(z_a2, 4)} -> {_interp_altman(z_a2)}), iar Scorul Conan-Holder este de {fmt_num(res.z_canon_holder[2], 4)} (ajustat {fmt_num(z_ch2, 4)} -> {_interp_ch(z_ch2)}), impunând măsuri de consolidare a capitalurilor proprii și a vitezei de rotație."
        )
    add_paragraph(doc, s17_msg)
