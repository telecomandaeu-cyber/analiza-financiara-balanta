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
    describe_trajectory,
    fmt_money,
    fmt_ratio_pct,
    make_table,
    pick_variant,
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
    dts_nefin: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    trez_activa: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    trez_pasiva: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    check: Tuple[float, float, float] = (0.0, 0.0, 0.0)


def compute_sheet_07_echilibrul_financiar(
    bf: BilantFinanciarResult, sf: StructuraFinanciaraResult
) -> EchilibruFinanciarResult:
    cperm = sf.cperm
    ai = bf.active_imobilizate
    ac = bf.active_circulante
    dts = sf.dts
    cp = sf.cp
    # Trezoreria activă = disponibilități + investiții financiare pe termen scurt;
    # trezoreria pasivă = credite bancare pe termen scurt (ct. 519, descoperit de cont 512).
    trez_activa = tuple(round(bf.disponibilitati[p] + bf.investitii_ts[p], 2) for p in range(3))
    trez_pasiva = bf.credite_trezorerie

    fr_sus = tuple(round(cperm[p] - ai[p], 2) for p in range(3))
    fr_jos = tuple(round(ac[p] - dts[p], 2) for p in range(3))

    frp = tuple(round(cp[p] - ai[p], 2) for p in range(3))
    frs = tuple(round(fr_sus[p] - frp[p], 2) for p in range(3))

    # NFR = (stocuri + creanțe) - (datorii pe termen scurt fără creditele de trezorerie)
    ac_fara_disp = tuple(round(ac[p] - trez_activa[p], 2) for p in range(3))
    dts_nefin = tuple(round(dts[p] - trez_pasiva[p], 2) for p in range(3))
    nfr = tuple(round(ac_fara_disp[p] - dts_nefin[p], 2) for p in range(3))

    # TN calculat direct (trezorerie activă - pasivă); identitatea TN = FR - NFR se verifică.
    tn = tuple(round(trez_activa[p] - trez_pasiva[p], 2) for p in range(3))
    check = tuple(round(fr_sus[p] - nfr[p] - tn[p], 2) for p in range(3))

    diag_fr = tuple("EFTL" if fr_sus[p] >= 0 else "DFTL" for p in range(3))
    diag_nfr = tuple("EFTS" if fr_sus[p] >= nfr[p] else "NFCE" for p in range(3))
    diag_tn = tuple("EFC" if tn[p] >= 0 else "DFT" for p in range(3))
    # TN > FR apare când NFR < 0: trezoreria este alimentată din resurse temporare de exploatare.
    semnal_tn_fr = tuple(
        "TN > FR (trezorerie finanțată din NFR negativ)" if tn[p] > fr_sus[p] else "TN ≤ FR"
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
        dts_nefin=dts_nefin,
        trez_activa=trez_activa,
        trez_pasiva=trez_pasiva,
        check=check,
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
        ("6. Datorii pe termen scurt fără creditele de trezorerie", ef.dts_nefin, "—", False),
        ("7. Necesarul de fond de rulment (NFR = rd. 5 - rd. 6)", ef.nfr, ef.diag_nfr[2], True),
        ("8. Trezoreria netă (TN = Disponibil + Inv. TS - Credite TS = FR - NFR)", ef.tn, ef.diag_tn[2], True),
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

    traj_fr = describe_trajectory(ef.fr_sus, bal, "s07_fr", "Fondul de rulment patrimonial (FR)")
    traj_tn = describe_trajectory(ef.tn, bal, "s07_tn", "Trezoreria netă (TN = FR - NFR)")

    if ef.fr_sus[2] >= 0 and ef.nfr[2] >= 0 and ef.tn[2] >= 0:
        eq_msg = pick_variant(
            bal,
            "s07_eq_1",
            [
                f"În perioada de raportare 2, excedentul de resurse permanente (FR = {fmt_money(ef.fr_sus[2])} lei, din care Fond de rulment propriu FRP = {fmt_money(ef.frp[2])} lei) finanțează integral Necesarul de fond de rulment (NFR = {fmt_money(ef.nfr[2])} lei) și degajă o trezorerie netă pozitivă ({fmt_money(ef.tn[2])} lei -> {ef.diag_tn[2]}).",
                f"Configurația din perioada de raportare 2 (FR = {fmt_money(ef.fr_sus[2])} lei > NFR = {fmt_money(ef.nfr[2])} lei > 0) atestă că întregul decalaj de finanțare al stocurilor și creanțelor este acoperit din fondul de rulment, lăsând disponibilități bănești nete excedentare de {fmt_money(ef.tn[2])} lei.",
            ],
        )
    elif ef.fr_sus[2] >= 0 and ef.nfr[2] < 0 and ef.tn[2] >= 0:
        eq_msg = pick_variant(
            bal,
            "s07_eq_2",
            [
                f"În perioada de raportare 2, Fondul de rulment pozitiv ({fmt_money(ef.fr_sus[2])} lei, FRP = {fmt_money(ef.frp[2])} lei) se cumulează cu o resursă netă atrasă din exploatare (NFR negativ = {fmt_money(ef.nfr[2])} lei), amplificând trezoreria netă la {fmt_money(ef.tn[2])} lei.",
                f"Valoarea negativă a NFR în perioada de raportare 2 ({fmt_money(ef.nfr[2])} lei) arată că datoriile de exploatare depășesc activele circulante non-monetare, degajând resurse suplimentare care se adaugă Fondului de rulment ({fmt_money(ef.fr_sus[2])} lei) în formarea trezoreriei nete ({fmt_money(ef.tn[2])} lei).",
            ],
        )
    else:
        eq_msg = (
            f"În perioada de raportare 2, raportul dintre Fondul de rulment (FR = {fmt_money(ef.fr_sus[2])} lei, FRP = {fmt_money(ef.frp[2])} lei -> {ef.diag_fr[2]}) și Necesarul de fond de rulment (NFR = {fmt_money(ef.nfr[2])} lei -> {ef.diag_nfr[2]}) determină o Trezorerie netă de {fmt_money(ef.tn[2])} lei ({ef.diag_tn[2]})."
        )

    ok_sus_jos = all(abs(ef.fr_sus[p] - ef.fr_jos[p]) <= 0.05 for p in range(3))
    ok_tn = all(abs(c) <= 0.05 for c in ef.check)
    if ok_sus_jos and ok_tn:
        verif = "Fondul de rulment calculat „de sus” (CPERM - AI) este egal cu cel calculat „de jos” (AC - DTS), iar identitatea TN = FR - NFR se verifică în toate perioadele"
    else:
        verif = "ATENȚIE: identitățile FR (sus) = FR (jos) și TN = FR - NFR nu se verifică în toate perioadele, ceea ce indică un bilanț neechilibrat"
    asoc2 = bf.imprumuturi_asociati[2]
    asoc_txt = (
        f" Datoriile pe termen scurt includ împrumuturile asociaților (ct. 455) de {fmt_money(asoc2)} lei, care, fiind tratate ca datorii de exploatare, reduc NFR; dacă acestea au caracter permanent, ar trebui privite ca resurse stabile, ceea ce ar majora corespunzător FR."
        if asoc2 > 0.005
        else ""
    )
    add_paragraph(
        doc,
        f"Diagnostic de echilibru financiar: {verif}; {traj_fr}, în timp ce {traj_tn}. {eq_msg}{asoc_txt}",
    )
