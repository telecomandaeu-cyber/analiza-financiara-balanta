#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 17: Evaluarea Riscului de Insolvență (Modelele Altman Z' și Conan-Holder).

1. Altman Z' (1983) — varianta pentru societăți necotate (capitalul propriu la valoare contabilă):
     Z' = 0,717·X1 + 0,847·X2 + 3,107·X3 + 0,420·X4 + 0,998·X5
     X1 = Fond de rulment net (AC - DTS) / Activ total
     X2 = Rezerve și rezultate reinvestite / Activ total
     X3 = EBIT (rezultat brut + dobânzi), anualizat / Activ total
     X4 = Capitaluri proprii / Datorii totale
     X5 = Cifra de afaceri anualizată / Activ total
     Zone: Z' > 2,90 sigur; 1,23 – 2,90 zonă gri; Z' < 1,23 risc ridicat.

2. Conan-Holder (1979):
     Z = 0,24·R1 + 0,22·R2 + 0,16·R3 - 0,87·R4 - 0,10·R5
     R1 = EBE anualizat / Datorii totale
     R2 = Capitaluri permanente / Total pasiv
     R3 = (Creanțe + Investiții TS + Disponibilități) / Total activ
     R4 = Cheltuieli financiare / Cifra de afaceri
     R5 = Cheltuieli cu personalul / Valoarea adăugată
     Zone: Z > 0,16 foarte bună; 0,10 – 0,16 bună; 0,04 – 0,10 vigilență; -0,05 – 0,04 pericol; Z < -0,05 eșec.

Fluxurile cumulate de la începutul anului se anualizează când sunt raportate la stocuri
(active, datorii). Pentru P0 (solduri de deschidere, fără rulaje) scorurile sunt n.d.
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
    fmt_num,
    is_nd,
    make_table,
    ratio_or_nd,
    safe_div,
)
from sheet_01_bilant import BilantResult
from sheet_02_cpp import CPPResult
from sheet_03_bilant_financiar import BilantFinanciarResult
from sheet_05_structura_financiara import StructuraFinanciaraResult
from sheet_07_echilibrul_financiar import EchilibruFinanciarResult
from sheet_13_sig import SIGResult

Triple = Tuple[float, float, float]


@dataclass
class RiscFalimentResult:
    alt_v1: Triple
    alt_v2: Triple
    alt_v3: Triple
    alt_v4: Triple
    alt_v5: Triple
    z_altman: Triple
    interp_altman: Tuple[str, str, str]
    z_altman_ajustat: Triple

    ch_v1: Triple
    ch_v2: Triple
    ch_v3: Triple
    ch_v4: Triple
    ch_v5: Triple
    z_canon_holder: Triple
    interp_canon_holder: Tuple[str, str, str]
    z_ch_ajustat: Triple


def _interp_altman(z: float) -> str:
    if is_nd(z):
        return "n.d."
    if z < 1.23:
        return "Faliment iminent / risc ridicat (Z' < 1,23)"
    if z <= 2.90:
        return "Zonă de incertitudine (1,23 ≤ Z' ≤ 2,90)"
    return "Societate solvabilă / risc redus (Z' > 2,90)"


def _interp_ch(z: float) -> str:
    if is_nd(z):
        return "n.d."
    if z > 0.16:
        return "Situație financiară bună / foarte bună (Z > 0,16)"
    if z > 0.10:
        return "Situație financiară bună (0,10 < Z ≤ 0,16)"
    if z > 0.04:
        return "Zonă de incertitudine / vigilență (0,04 < Z ≤ 0,10)"
    if z > -0.05:
        return "Dificultate financiară / pericol (-0,05 < Z ≤ 0,04)"
    return "Faliment iminent / eșec (Z ≤ -0,05)"


def _t(fn) -> Triple:
    return tuple(fn(p) for p in range(3))  # type: ignore[return-value]


def compute_sheet_17_riscul_faliment(
    bil: BilantResult,
    cpp: CPPResult,
    bf: BilantFinanciarResult,
    sf: StructuraFinanciaraResult,
    ef: EchilibruFinanciarResult,
    sig: SIGResult,
    annual_factors: Tuple[float, float, float] = (ND, 365 / 243, 365 / 273),
) -> RiscFalimentResult:
    k = annual_factors
    ta = bf.total_activ
    ca = cpp.cifra_afaceri
    cp = sf.cp
    dt = sf.dtot
    ebit = _t(lambda p: cpp.rezultat_brut[p] + cpp.dobanzi[p])

    # ---------------- Altman Z' (societăți necotate)
    x1 = _t(lambda p: safe_div(ef.fr_jos[p], ta[p]))
    x2 = _t(lambda p: safe_div(bf.rezerve_si_rezultate[p], ta[p]))
    x3 = _t(lambda p: safe_div(ebit[p] * k[p], ta[p]))
    x4 = _t(lambda p: ratio_or_nd(cp[p], dt[p]))
    x5 = _t(lambda p: safe_div(ca[p] * k[p], ta[p]))

    def _z(p: int, x4v: float) -> float:
        return 0.717 * x1[p] + 0.847 * x2[p] + 3.107 * x3[p] + 0.420 * x4v + 0.998 * x5[p]

    z_a = _t(lambda p: _z(p, x4[p]))
    i_a = _t(lambda p: _interp_altman(z_a[p]))
    # Varianta ajustată: datoriile fără veniturile în avans (ct. 472, 475, 478), care nu sunt exigibile.
    dt_ef = _t(lambda p: bil.rd(11)[p] + bil.rd(14)[p])
    x4_ef = _t(lambda p: ratio_or_nd(cp[p], dt_ef[p]))
    z_a_ef = _t(lambda p: _z(p, x4_ef[p]))

    # ---------------- Conan-Holder
    cr_inv_disp = _t(lambda p: bf.creante[p] + bf.investitii_ts[p] + bf.disponibilitati[p])
    ch_fin = cpp.cheltuieli_financiare
    ch_pers = cpp.rd(18)
    va = sig.valoarea_adaugata
    ebe = sig.ebe

    r1 = _t(lambda p: ratio_or_nd(ebe[p] * k[p], dt[p]))
    r2 = _t(lambda p: safe_div(sf.cperm[p], sf.pasiv1[p]))
    r3 = _t(lambda p: safe_div(cr_inv_disp[p], ta[p]))
    r4 = _t(lambda p: ratio_or_nd(ch_fin[p], ca[p]))
    r5 = _t(lambda p: ratio_or_nd(ch_pers[p], va[p]) if not is_nd(va[p]) and va[p] > 0 else ND)

    def _zc(p: int, r1v: float) -> float:
        return 0.24 * r1v + 0.22 * r2[p] + 0.16 * r3[p] - 0.87 * r4[p] - 0.10 * r5[p]

    z_c = _t(lambda p: _zc(p, r1[p]))
    i_c = _t(lambda p: _interp_ch(z_c[p]))
    r1_ef = _t(lambda p: ratio_or_nd(ebe[p] * k[p], dt_ef[p]))
    z_c_aj = _t(lambda p: _zc(p, r1_ef[p]))

    return RiscFalimentResult(
        alt_v1=x1, alt_v2=x2, alt_v3=x3, alt_v4=x4, alt_v5=x5,
        z_altman=z_a, interp_altman=i_a, z_altman_ajustat=z_a_ef,
        ch_v1=r1, ch_v2=r2, ch_v3=r3, ch_v4=r4, ch_v5=r5,
        z_canon_holder=z_c, interp_canon_holder=i_c, z_ch_ajustat=z_c_aj,
    )


def render_sheet_17_riscul_faliment(doc: Document, bal: BalantaData, res: RiscFalimentResult) -> None:
    """Generează secțiunea Evaluării Riscului de Insolvență."""
    add_heading(doc, "17. Evaluarea Riscului de Insolvență (Modelele de Scoring Altman Z' și Conan-Holder)", level=1)

    headers = ["Model / Variabilă a funcției-scor", "Pondere", *bal.col_labels]

    def _row(lbl: str, pond: str, vals: Triple, dec: int = 4):
        return [lbl, pond, *[fmt_num(v, dec) for v in vals]]

    rows = [
        _row("Altman X1 = Fond de rulment net / Activ total", "0,717", res.alt_v1),
        _row("Altman X2 = Rezerve și rezultate reinvestite / Activ total", "0,847", res.alt_v2),
        _row("Altman X3 = EBIT anualizat / Activ total", "3,107", res.alt_v3),
        _row("Altman X4 = Capitaluri proprii / Datorii totale", "0,420", res.alt_v4),
        _row("Altman X5 = Cifra de afaceri anualizată / Activ total", "0,998", res.alt_v5),
        _row("➔ Scorul Altman Z' (societăți necotate)", "—", res.z_altman),
        _row("➔ Scorul Altman Z' ajustat (datorii fără venituri în avans)", "—", res.z_altman_ajustat),
        _row("Conan-Holder R1 = EBE anualizat / Datorii totale", "+0,24", res.ch_v1),
        _row("Conan-Holder R2 = Capitaluri permanente / Total pasiv", "+0,22", res.ch_v2),
        _row("Conan-Holder R3 = (Creanțe + Inv. TS + Disponibil) / Activ total", "+0,16", res.ch_v3),
        _row("Conan-Holder R4 = Cheltuieli financiare / Cifra de afaceri", "-0,87", res.ch_v4),
        _row("Conan-Holder R5 = Cheltuieli cu personalul / Valoarea adăugată", "-0,10", res.ch_v5),
        _row("➔ Scorul Conan-Holder Z", "—", res.z_canon_holder),
        _row("➔ Scorul Conan-Holder Z ajustat (datorii fără venituri în avans)", "—", res.z_ch_ajustat),
        ["Încadrare Altman Z'", "—", "n.d.", _interp_altman(res.z_altman[1]), _interp_altman(res.z_altman[2])],
        ["Încadrare Conan-Holder Z", "—", "n.d.", _interp_ch(res.z_canon_holder[1]), _interp_ch(res.z_canon_holder[2])],
    ]
    make_table(
        doc,
        headers=headers,
        rows=rows,
        col_widths=[7.2, 1.3, 2.4, 3.3, 3.3],
        total_rows=[5, 6, 12, 13],
        font_size=7.6,
    )

    za1, za2 = res.z_altman[1], res.z_altman[2]
    zc2 = res.z_canon_holder[2]
    add_paragraph(
        doc,
        f"Evaluarea riscului de insolvență: scorul Altman Z' evoluează de la {fmt_num(za1, 4)} în perioada de raportare 1 la "
        f"{fmt_num(za2, 4)} în perioada de raportare 2 ({_interp_altman(za2)}; varianta ajustată {fmt_num(res.z_altman_ajustat[2], 4)}), "
        f"iar scorul Conan-Holder în perioada de raportare 2 este {fmt_num(zc2, 4)} ({_interp_ch(zc2)}; ajustat "
        f"{fmt_num(res.z_ch_ajustat[2], 4)}). Variabilele de flux (EBIT, EBE, cifra de afaceri) sunt anualizate; pentru perioada de "
        f"raportare 0 (solduri de deschidere, fără rulaje) scorurile nu pot fi calculate. Modelele de scoring sunt calibrate pe alte "
        f"economii și perioade, de aceea rezultatul are valoare orientativă și trebuie coroborat cu indicatorii de lichiditate și solvabilitate.",
    )
