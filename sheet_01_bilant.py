#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 1: Bilanțul Contabil Sintetic (OMFP nr. 1802/2014).
Preia soldurile din balanța de verificare pentru cele trei perioade de raportare:
  - Perioada de raportare 0 (01.01.2026)
  - Perioada de raportare 1 (31.08.2026)
  - Perioada de raportare 2 (30.09.2026)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from docx import Document

from core_balanta import (
    BalantaData,
    add_heading,
    add_paragraph,
    describe_trajectory,
    fmt_money,
    fmt_pct,
    fmt_ratio_pct,
    make_table,
    pick_variant,
    safe_div,
)


@dataclass
class BilantRow:
    grupa: str
    denumire: str
    formula_conturi: str
    rand: str
    v0: float
    v1: float
    v2: float
    is_total: bool = False

    @property
    def vals(self) -> Tuple[float, float, float]:
        return (self.v0, self.v1, self.v2)


@dataclass
class BilantResult:
    rows_by_code: Dict[int, Tuple[float, float, float]] = field(default_factory=dict)
    table_rows: List[BilantRow] = field(default_factory=list)

    def rd(self, code: int) -> Tuple[float, float, float]:
        return self.rows_by_code.get(code, (0.0, 0.0, 0.0))


def compute_sheet_01_bilant(bal: BalantaData) -> BilantResult:
    """Calculează rândurile Bilanțului Contabil conform OMFP nr. 1802/2014."""
    res = BilantResult()
    r: Dict[int, List[float]] = {k: [0.0, 0.0, 0.0] for k in range(1, 42)}

    for p in (0, 1, 2):
        # A. ACTIVE IMOBILIZATE
        imob_necorp_brut = sum(bal.sold_d(c, p) for c in ("201", "203", "205", "2071", "208", "233", "234"))
        amort_necorp = sum(bal.sold_c(c, p) for c in ("280", "290", "2933"))
        r[1][p] = round(imob_necorp_brut - amort_necorp, 2)

        imob_corp_brut = sum(bal.sold_d(c, p) for c in ("211", "212", "213", "214", "223", "224", "231", "232"))
        amort_corp = sum(bal.sold_c(c, p) for c in ("281", "291", "2931"))
        r[2][p] = round(imob_corp_brut - amort_corp, 2)

        imob_fin_brut = sum(bal.sold_d(c, p) for c in ("261", "263", "265", "267"))
        depr_fin = bal.sold_c("296", p)
        r[3][p] = round(imob_fin_brut - depr_fin, 2)

        r[4][p] = round(r[1][p] + r[2][p] + r[3][p], 2)

        # B. ACTIVE CIRCULANTE
        stoc_brut = sum(
            bal.sold_d(c, p)
            for c in ("301", "302", "303", "321", "322", "323", "331", "332", "341", "345", "346", "351", "361", "371", "381")
        )
        dif_pret = sum(bal.net_d(c, p) for c in ("308", "348", "368", "378", "388"))
        depr_stoc = sum(bal.sold_c(c, p) for c in ("391", "392", "393", "394", "395", "396", "397", "398"))
        furn_deb_stoc = bal.pos_d("409", p)
        r[5][p] = round(stoc_brut + dif_pret - depr_stoc + furn_deb_stoc, 2)

        creante_fixe = (
            bal.pos_d("411", p)
            + bal.pos_d("413", p)
            + bal.pos_d("418", p)
            + bal.pos_d("425", p)
            + bal.pos_d("428", p)
            + bal.pos_d("431", p)
            + bal.pos_d("436", p)
            + bal.pos_d("437", p)
            + bal.pos_d("438", p)
            + bal.pos_d("441", p)
            + bal.pos_d("4424", p)
            + bal.pos_d("4426", p)
            + bal.pos_d("444", p)
            + bal.pos_d("445", p)
            + bal.pos_d("446", p)
            + bal.pos_d("447", p)
            + bal.pos_d("448", p)
            + bal.pos_d("451", p)
            + bal.pos_d("453", p)
            + bal.pos_d("456", p)
            + bal.pos_d("458", p)
            + bal.pos_d("461", p)
            + bal.pos_d("463", p)
            + bal.pos_d("473", p)
            + bal.pos_d("518", p)
        )
        ajust_creante = sum(bal.sold_c(c, p) for c in ("491", "495", "496"))
        r[6][p] = round(creante_fixe - ajust_creante, 2)

        inv_ts = sum(bal.pos_d(c, p) for c in ("501", "505", "506", "508")) - sum(
            bal.sold_c(c, p) for c in ("591", "595", "596", "598")
        )
        r[7][p] = round(inv_ts, 2)

        disp = sum(bal.pos_d(c, p) for c in ("512", "531", "532", "541", "542", "581"))
        r[8][p] = round(disp, 2)

        r[9][p] = round(r[5][p] + r[6][p] + r[7][p] + r[8][p], 2)

        # C. CHELTUIELI ÎN AVANS
        r[10][p] = round(bal.pos_d("471", p), 2)

        # D. DATORII SUB UN AN
        dat_scurte = (
            bal.pos_c("269", p)
            + bal.pos_c("401", p)
            + bal.pos_c("403", p)
            + bal.pos_c("404", p)
            + bal.pos_c("405", p)
            + bal.pos_c("408", p)
            + bal.pos_c("419", p)
            + bal.pos_c("421", p)
            + bal.pos_c("423", p)
            + bal.pos_c("424", p)
            + bal.pos_c("426", p)
            + bal.pos_c("427", p)
            + bal.pos_c("428", p)
            + bal.pos_c("431", p)
            + bal.pos_c("436", p)
            + bal.pos_c("437", p)
            + bal.pos_c("438", p)
            + bal.pos_c("441", p)
            + bal.pos_c("4423", p)
            + bal.pos_c("4428", p)
            + bal.pos_c("444", p)
            + bal.pos_c("446", p)
            + bal.pos_c("447", p)
            + bal.pos_c("448", p)
            + bal.pos_c("451", p)
            + bal.pos_c("453", p)
            + bal.pos_c("455", p)
            + bal.pos_c("456", p)
            + bal.pos_c("457", p)
            + bal.pos_c("458", p)
            + bal.pos_c("462", p)
            + bal.pos_c("473", p)
            + bal.pos_c("509", p)
            + bal.pos_c("518", p)
            + bal.pos_c("519", p)
        )
        r[11][p] = round(dat_scurte, 2)

        # I. VENITURI ÎN AVANS
        r[17][p] = round(bal.pos_c("475", p), 2)
        r[18][p] = round(bal.pos_c("472", p), 2)
        r[19][p] = r[18][p]
        r[20][p] = 0.0
        r[21][p] = round(bal.pos_c("478", p), 2)
        r[22][p] = round(bal.pos_c("2075", p), 2)
        r[16][p] = round(r[17][p] + r[18][p] + r[21][p] + r[22][p], 2)

        # E. ACTIVE CIRCULANTE NETE / DATORII CURENTE NETE
        r[12][p] = round(r[9][p] + r[10][p] - r[11][p] - r[19][p], 2)

        # F. TOTAL ACTIVE MINUS DATORII CURENTE
        r[13][p] = round(r[4][p] + r[12][p], 2)

        # G. DATORII PESTE UN AN
        dat_lungi = (
            bal.pos_c("161", p)
            + bal.pos_c("162", p)
            + bal.pos_c("166", p)
            + bal.pos_c("167", p)
            + bal.pos_c("168", p)
            - bal.pos_d("169", p)
        )
        r[14][p] = round(dat_lungi, 2)

        # H. PROVIZIOANE
        r[15][p] = round(bal.pos_c("151", p), 2)

        # J. CAPITAL ȘI REZERVE
        r[24][p] = round(bal.pos_c("1012", p), 2)
        r[25][p] = round(bal.pos_c("1011", p), 2)
        r[26][p] = round(bal.pos_c("1015", p), 2)
        r[27][p] = round(bal.pos_c("1018", p), 2)
        r[23][p] = round(r[24][p] + r[25][p] + r[26][p] + r[27][p], 2)
        r[28][p] = round(bal.pos_c("104", p), 2)
        r[29][p] = round(bal.pos_c("105", p), 2)
        r[30][p] = round(bal.pos_c("106", p), 2)
        r[31][p] = round(bal.pos_d("109", p), 2)
        r[32][p] = round(bal.pos_c("141", p), 2)
        r[33][p] = round(bal.pos_d("149", p), 2)

        r[34][p] = round(bal.pos_c("117", p), 2)
        r[35][p] = round(bal.pos_d("117", p), 2)

        r[36][p] = round(bal.pos_c("121", p), 2)
        r[37][p] = round(bal.pos_d("121", p), 2)

        r[38][p] = round(bal.pos_d("129", p), 2)

        r[39][p] = round(
            r[23][p]
            + r[28][p]
            + r[29][p]
            + r[30][p]
            - r[31][p]
            + r[32][p]
            - r[33][p]
            + r[34][p]
            - r[35][p]
            + r[36][p]
            - r[37][p]
            - r[38][p],
            2,
        )
        r[40][p] = round(bal.pos_c("1016", p), 2)
        r[41][p] = round(r[39][p] + r[40][p], 2)

    for code, vals in r.items():
        res.rows_by_code[code] = (vals[0], vals[1], vals[2])

    definitions = [
        ("A.", "I. Imobilizări necorporale", "(ct. 201..208 - 280)", 1, False),
        ("", "II. Imobilizări corporale", "(ct. 211..231 - 281)", 2, False),
        ("", "ACTIVE IMOBILIZATE — TOTAL", "(rd. 01 + 02 + 03)", 4, True),
        ("B.", "I. Stocuri (materii prime, produse, mărfuri, avansuri)", "(ct. 301..381 + 4091)", 5, False),
        ("", "II. Creanțe comerciale, fiscale și diverse", "(ct. 411+425+438+4424+445+461+463)", 6, False),
        ("", "IV. Casa și conturi la bănci", "(ct. 512+531+542)", 8, False),
        ("", "ACTIVE CIRCULANTE — TOTAL", "(rd. 05 + 06 + 07 + 08)", 9, True),
        ("C.", "Cheltuieli în avans", "(ct. 471)", 10, False),
        ("D.", "DATORII CE TREBUIE PLĂTITE ÎNTR-O PERIOADĂ DE PÂNĂ LA UN AN", "(ct. 401+419+421+431+436+441+4423+444+455+457+462)", 11, True),
        ("E.", "Active circulante nete / Datorii curente nete", "(rd. 09 + 10 - 11 - 19)", 12, False),
        ("F.", "TOTAL ACTIVE MINUS DATORII CURENTE", "(rd. 04 + 12)", 13, True),
        ("G.", "DATORII CE TREBUIE PLĂTITE ÎNTR-O PERIOADĂ MAI MARE DE UN AN", "(ct. 161..168)", 14, True),
        ("I.", "VENITURI ÎN AVANS — TOTAL", "(rd. 17 + 18)", 16, True),
        ("J.", "I. Capital subscris vărsat", "(ct. 1012)", 24, False),
        ("", "IV. Rezerve", "(ct. 106)", 30, False),
        ("", "V. Profitul sau pierderea reportat(ă) — Sold C", "(ct. 117)", 34, False),
        ("", "VI. Profitul exercițiului financiar — Sold C", "(ct. 121)", 36, False),
        ("", "Repartizarea profitului", "(ct. 129)", 38, False),
        ("", "CAPITALURI PROPRII — TOTAL", "(rd. 23 + 30 + 34 + 36 - 38)", 39, True),
    ]

    for grp, den, form_str, code, is_tot in definitions:
        v0, v1, v2 = res.rd(code)
        if not is_tot and abs(v0) < 1e-6 and abs(v1) < 1e-6 and abs(v2) < 1e-6:
            continue
        res.table_rows.append(
            BilantRow(
                grupa=grp,
                denumire=f"{den} {form_str}",
                formula_conturi=form_str,
                rand=f"{code:02d}",
                v0=v0,
                v1=v1,
                v2=v2,
                is_total=is_tot,
            )
        )

    return res


def render_sheet_01_bilant(doc: Document, bal: BalantaData, res: BilantResult) -> None:
    """Generează secțiunea Bilanțului Contabil Sintetic în documentul Word."""
    add_heading(doc, "1. Bilanțul Contabil Sintetic (OMFP nr. 1802/2014)", level=1)

    headers = [
        "Grupa",
        "Element patrimonial și corespondență contabilă",
        "Rd.",
        "Perioada de raportare 0 (lei)",
        "Perioada de raportare 1 (lei)",
        "Perioada de raportare 2 (lei)",
        "Δ (P2 - P0) (lei)",
    ]
    rows_data = []
    total_indices = []
    for idx, r in enumerate(res.table_rows):
        if r.is_total:
            total_indices.append(idx)
        rows_data.append([
            r.grupa,
            r.denumire,
            r.rand,
            fmt_money(r.v0),
            fmt_money(r.v1),
            fmt_money(r.v2),
            fmt_money(r.v2 - r.v0),
        ])

    make_table(
        doc,
        headers=headers,
        rows=rows_data,
        col_widths=[0.9, 6.8, 0.8, 2.3, 2.3, 2.3, 2.4],
        total_rows=total_indices,
        font_size=7.8,
    )

    ai0, ai1, ai2 = res.rd(4)
    ac0, ac1, ac2 = res.rd(9)
    ca0, ca1, ca2 = res.rd(10)
    ta0, ta1, ta2 = ai0 + ac0 + ca0, ai1 + ac1 + ca1, ai2 + ac2 + ca2
    dts0, dts1, dts2 = res.rd(11)
    dtl0, dtl1, dtl2 = res.rd(14)
    cp0, cp1, cp2 = res.rd(39)
    disp2 = res.rd(8)[2]
    stoc2 = res.rd(5)[2]
    creante2 = res.rd(6)[2]

    traj_ta = describe_trajectory((ta0, ta1, ta2), bal, "s01_ta", "volumul total al activului (egal cu pasivul total)")
    traj_cp = describe_trajectory((cp0, cp1, cp2), bal, "s01_cp", "capitalul propriu (activul net contabil)")

    pond_ai2 = safe_div(ai2, ta2)
    if pond_ai2 >= 0.50:
        profil_activ = pick_variant(
            bal,
            "s01_prof_ai",
            [
                f"Structura patrimonială din perioada de raportare 2 este preponderent imobilizată (activele fixe nete de {fmt_money(ai2)} lei dețin {fmt_ratio_pct(pond_ai2)} din activ), reflectând o bază tehnică și investițională consistentă, completată de active circulante de {fmt_money(ac2)} lei.",
                f"Arhitectura activului în perioada de raportare 2 evidențiază dominanța imobilizărilor nete ({fmt_money(ai2)} lei, respectiv {fmt_ratio_pct(pond_ai2)} din patrimoniu), în timp ce activele curente însumează {fmt_money(ac2)} lei (din care trezorerie {fmt_money(disp2)} lei).",
                f"Alocarea plasamentelor patrimoniale la finele perioadei de raportare 2 relevă o intensitate ridicată a capitalului fix ({fmt_money(ai2)} lei, {fmt_ratio_pct(pond_ai2)} din total), secondată de active circulante de {fmt_money(ac2)} lei.",
            ],
        )
    else:
        dom_ac = "disponibilitățile bănești" if disp2 >= max(stoc2, creante2) else ("creanțele comerciale și diverse" if creante2 >= stoc2 else "stocurile operaționale")
        profil_activ = pick_variant(
            bal,
            "s01_prof_ac",
            [
                f"Configurația activului în perioada de raportare 2 este orientată spre ciclul operațional curent: activele circulante ({fmt_money(ac2)} lei, pilonul principal fiind {dom_ac}) dețin {fmt_ratio_pct(1.0 - pond_ai2)} din patrimoniu, iar imobilizările nete însumează {fmt_money(ai2)} lei ({fmt_ratio_pct(pond_ai2)}).",
                f"În perioada de raportare 2, mobilitatea patrimonială este ridicată, activele circulante ({fmt_money(ac2)} lei, susținute prioritar de {dom_ac}) devansând activele imobilizate nete ({fmt_money(ai2)} lei, respectiv {fmt_ratio_pct(pond_ai2)} din activ).",
                f"Ponderea majoritară în activul perioadei de raportare 2 revine elementelor circulante ({fmt_money(ac2)} lei, dominate de {dom_ac}), în timp ce baza imobilizată netă se cifrează la {fmt_money(ai2)} lei ({fmt_ratio_pct(pond_ai2)} din total).",
            ],
        )

    if cp2 < 0:
        pasiv_msg = f"Pe latura resurselor, deși {traj_cp}, valoarea negativă a capitalurilor proprii indică un deficit de acoperire patrimonială, finanțarea fiind asigurată din datorii curente ({fmt_money(dts2)} lei) și datorii pe termen lung ({fmt_money(dtl2)} lei)."
    elif dtl2 > 0:
        pasiv_msg = pick_variant(
            bal,
            "s01_pas_dtl",
            [
                f"În planul finanțării, {traj_cp}, iar obligațiile față de terți în perioada de raportare 2 se împart între datorii curente exigibile sub un an ({fmt_money(dts2)} lei) și împrumuturi pe termen lung ({fmt_money(dtl2)} lei).",
                f"Sub aspectul surselor de finanțare, {traj_cp}; complementar, datoriile pe termen scurt se situează la {fmt_money(dts2)} lei în perioada de raportare 2, iar datoriile peste un an la {fmt_money(dtl2)} lei.",
                f"Pe partea de pasiv, {traj_cp}, resursele împrumutate însumând în perioada de raportare 2 datorii curente de {fmt_money(dts2)} lei și datorii pe termen lung de {fmt_money(dtl2)} lei.",
            ],
        )
    else:
        pasiv_msg = pick_variant(
            bal,
            "s01_pas_nodtl",
            [
                f"În planul resurselor, {traj_cp}; societatea nu înregistrează datorii pe termen lung în perioada de raportare 2 (0,00 lei), obligațiile față de terți fiind integral pe termen scurt ({fmt_money(dts2)} lei).",
                f"Sub aspectul structurii pasivului, {traj_cp}, în timp ce datoriile pe termen lung sunt nule în perioada de raportare 2, iar angajamentele curente sub un an totalizează {fmt_money(dts2)} lei.",
            ],
        )

    add_paragraph(doc, f"Diagnostic bilanțier: În conformitate cu principiul dublei reprezentări, {traj_ta}. {profil_activ} {pasiv_msg}")
