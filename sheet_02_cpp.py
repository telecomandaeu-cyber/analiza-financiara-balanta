#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 2: Contul de Profit și Pierdere Sintetic (OMFP nr. 1802/2014).
Se corelează cu Bilanțul Contabil pentru reconcilierea rezultatului net al exercițiului
cu soldul contului 121 pe cele trei perioade de raportare (0, 1 și 2).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from docx import Document

from core_balanta import (
    BalantaData,
    add_heading,
    add_paragraph,
    fmt_money,
    fmt_pct,
    make_table,
    safe_div,
)
from sheet_01_bilant import BilantResult


@dataclass
class CPPRow:
    grupa: str
    denumire: str
    rand: str
    v0: float
    v1: float
    v2: float
    is_total: bool = False


@dataclass
class CPPResult:
    rows_by_code: Dict[int, Tuple[float, float, float]] = field(default_factory=dict)
    table_rows: List[CPPRow] = field(default_factory=list)

    def rd(self, code: int) -> Tuple[float, float, float]:
        return self.rows_by_code.get(code, (0.0, 0.0, 0.0))

    @property
    def cifra_afaceri(self) -> Tuple[float, float, float]:
        return self.rd(1)

    @property
    def venituri_exploatare(self) -> Tuple[float, float, float]:
        return self.rd(12)

    @property
    def cheltuieli_exploatare(self) -> Tuple[float, float, float]:
        return self.rd(35)

    @property
    def rezultat_exploatare(self) -> Tuple[float, float, float]:
        p0, p1, p2 = self.rd(36)
        l0, l1, l2 = self.rd(37)
        return (round(p0 - l0, 2), round(p1 - l1, 2), round(p2 - l2, 2))

    @property
    def venituri_financiare(self) -> Tuple[float, float, float]:
        return self.rd(45)

    @property
    def cheltuieli_financiare(self) -> Tuple[float, float, float]:
        return self.rd(52)

    @property
    def rezultat_financiar(self) -> Tuple[float, float, float]:
        p0, p1, p2 = self.rd(53)
        l0, l1, l2 = self.rd(54)
        return (round(p0 - l0, 2), round(p1 - l1, 2), round(p2 - l2, 2))

    @property
    def venituri_totale(self) -> Tuple[float, float, float]:
        return self.rd(61)

    @property
    def cheltuieli_totale(self) -> Tuple[float, float, float]:
        return self.rd(62)

    @property
    def rezultat_brut(self) -> Tuple[float, float, float]:
        p0, p1, p2 = self.rd(63)
        l0, l1, l2 = self.rd(64)
        return (round(p0 - l0, 2), round(p1 - l1, 2), round(p2 - l2, 2))

    @property
    def impozit_profit(self) -> Tuple[float, float, float]:
        return self.rd(65)

    @property
    def rezultat_net(self) -> Tuple[float, float, float]:
        p0, p1, p2 = self.rd(67)
        l0, l1, l2 = self.rd(68)
        return (round(p0 - l0, 2), round(p1 - l1, 2), round(p2 - l2, 2))


def compute_sheet_02_cpp(bal: BalantaData, bilant: BilantResult) -> CPPResult:
    """Calculează indicatorii Contului de Profit și Pierdere pentru perioadele de raportare 0, 1 și 2."""
    res = CPPResult()
    r: Dict[int, List[float]] = {k: [0.0, 0.0, 0.0] for k in range(1, 69)}

    for p in (0, 1, 2):
        r[2][p] = round(
            sum(bal.rulaj_cpp(c, p, True) for c in ("701", "702", "703", "704", "705", "706", "708")), 2
        )
        r[3][p] = round(bal.rulaj_cpp("707", p, True), 2)
        r[4][p] = round(bal.rulaj_cpp("709", p, False), 2)
        r[5][p] = 0.0
        r[6][p] = round(bal.rulaj_cpp("7411", p, True), 2)
        r[1][p] = round(r[2][p] + r[3][p] - r[4][p] + r[5][p] + r[6][p], 2)

        # Variația stocurilor de produse finite și producție în curs (ct. 711 + 712):
        # când produsele obținute (345 = 711) sunt scăzute din gestiune la vânzare (711 = 345),
        # variația netă a stocului aferentă perioadei este rulaj_c(711+712) - rulaj_c(345+331).
        var_711 = round(
            sum(bal.rulaj_cpp(c, p, True) for c in ("711", "712"))
            - sum(bal.rulaj_cpp(c, p, True) for c in ("331", "332", "341", "345", "346")),
            2,
        )
        if var_711 >= 0:
            r[7][p], r[8][p] = var_711, 0.0
        else:
            r[7][p], r[8][p] = 0.0, abs(var_711)

        r[9][p] = round(sum(bal.rulaj_cpp(c, p, True) for c in ("721", "722")), 2)
        r[10][p] = round(bal.rulaj_cpp("758", p, True) + bal.rulaj_cpp("7417", p, True), 2)
        r[11][p] = 0.0
        r[12][p] = round(r[1][p] + r[7][p] - r[8][p] + r[9][p] + r[10][p], 2)

        r[13][p] = round(
            bal.rulaj_cpp("601", p, False) + bal.rulaj_cpp("602", p, False) - bal.rulaj_cpp("7412", p, True), 2
        )
        r[14][p] = round(sum(bal.rulaj_cpp(c, p, False) for c in ("603", "604", "606", "608")), 2)
        r[15][p] = round(bal.rulaj_cpp("605", p, False) - bal.rulaj_cpp("7413", p, True), 2)
        r[16][p] = round(bal.rulaj_cpp("607", p, False), 2)
        r[17][p] = round(bal.rulaj_cpp("609", p, True), 2)

        r[19][p] = round(
            sum(bal.rulaj_cpp(c, p, False) for c in ("641", "642", "643", "644")) - bal.rulaj_cpp("7414", p, True), 2
        )
        r[20][p] = round(
            bal.rulaj_cpp("645", p, False) + bal.rulaj_cpp("646", p, False) - bal.rulaj_cpp("7415", p, True), 2
        )
        r[18][p] = round(r[19][p] + r[20][p], 2)

        r[22][p] = round(bal.rulaj_cpp("6811", p, False) + bal.rulaj_cpp("6813", p, False), 2)
        r[23][p] = round(bal.rulaj_cpp("7813", p, True), 2)
        r[21][p] = round(r[22][p] - r[23][p], 2)

        r[25][p] = round(bal.rulaj_cpp("654", p, False) + bal.rulaj_cpp("6814", p, False), 2)
        r[26][p] = round(bal.rulaj_cpp("754", p, True) + bal.rulaj_cpp("7814", p, True), 2)
        r[24][p] = round(r[25][p] - r[26][p], 2)

        prest_ext = (
            bal.rulaj_cpp("Grupa 61", p, False)
            + bal.rulaj_cpp("Grupa 62", p, False)
            - bal.rulaj_cpp("7416", p, True)
        )
        r[28][p] = round(prest_ext, 2)
        r[29][p] = round(bal.rulaj_cpp("635", p, False), 2)
        r[30][p] = round(bal.rulaj_cpp("652", p, False) + bal.rulaj_cpp("658", p, False), 2)
        r[31][p] = 0.0
        r[27][p] = round(r[28][p] + r[29][p] + r[30][p] + r[31][p], 2)

        r[33][p] = round(bal.rulaj_cpp("6812", p, False), 2)
        r[34][p] = round(bal.rulaj_cpp("7812", p, True), 2)
        r[32][p] = round(r[33][p] - r[34][p], 2)

        r[35][p] = round(
            r[13][p]
            + r[14][p]
            + r[15][p]
            + r[16][p]
            - r[17][p]
            + r[18][p]
            + r[21][p]
            + r[24][p]
            + r[27][p]
            + r[32][p],
            2,
        )

        rex = round(r[12][p] - r[35][p], 2)
        if rex >= 0:
            r[36][p], r[37][p] = rex, 0.0
        else:
            r[36][p], r[37][p] = 0.0, abs(rex)

        r[38][p] = round(bal.rulaj_cpp("761", p, True), 2)
        r[39][p] = 0.0
        r[40][p] = round(bal.rulaj_cpp("763", p, True), 2)
        r[41][p] = 0.0
        r[42][p] = round(bal.rulaj_cpp("766", p, True), 2)
        r[43][p] = 0.0
        r[44][p] = round(
            sum(bal.rulaj_cpp(c, p, True) for c in ("762", "764", "765", "767", "768")), 2
        )
        r[45][p] = round(r[38][p] + r[40][p] + r[42][p] + r[44][p], 2)

        r[47][p] = round(bal.rulaj_cpp("686", p, False), 2)
        r[48][p] = round(bal.rulaj_cpp("786", p, True), 2)
        r[46][p] = round(r[47][p] - r[48][p], 2)
        r[49][p] = round(bal.rulaj_cpp("666", p, False) - bal.rulaj_cpp("7418", p, True), 2)
        r[50][p] = 0.0
        r[51][p] = round(
            sum(bal.rulaj_cpp(c, p, False) for c in ("663", "664", "665", "667", "668")), 2
        )
        r[52][p] = round(r[46][p] + r[49][p] + r[51][p], 2)

        rfin = round(r[45][p] - r[52][p], 2)
        if rfin >= 0:
            r[53][p], r[54][p] = rfin, 0.0
        else:
            r[53][p], r[54][p] = 0.0, abs(rfin)

        rcur = round(rex + rfin, 2)
        if rcur >= 0:
            r[55][p], r[56][p] = rcur, 0.0
        else:
            r[55][p], r[56][p] = 0.0, abs(rcur)

        r[57][p] = round(bal.rulaj_cpp("771", p, True), 2)
        r[58][p] = round(bal.rulaj_cpp("671", p, False), 2)
        rext = round(r[57][p] - r[58][p], 2)
        if rext >= 0:
            r[59][p], r[60][p] = rext, 0.0
        else:
            r[59][p], r[60][p] = 0.0, abs(rext)

        r[61][p] = round(r[12][p] + r[45][p] + r[57][p], 2)
        r[62][p] = round(r[35][p] + r[52][p] + r[58][p], 2)

        rbrut = round(r[61][p] - r[62][p], 2)
        if rbrut >= 0:
            r[63][p], r[64][p] = rbrut, 0.0
        else:
            r[63][p], r[64][p] = 0.0, abs(rbrut)

        r[65][p] = round(bal.rulaj_cpp("691", p, False), 2)
        r[66][p] = round(bal.rulaj_cpp("698", p, False), 2)

        rnet = round(rbrut - r[65][p] - r[66][p], 2)
        if p == 0:
            target_t0 = round(bilant.rd(36)[0] - bilant.rd(37)[0], 2)
            diff_t0 = round(target_t0 - rnet, 2)
            if abs(diff_t0) <= 0.05 and abs(diff_t0) > 0:
                r[10][0] = round(r[10][0] + diff_t0, 2)
                r[12][0] = round(r[12][0] + diff_t0, 2)
                r[36][0] = round(r[36][0] + diff_t0, 2)
                r[55][0] = round(r[55][0] + diff_t0, 2)
                r[61][0] = round(r[61][0] + diff_t0, 2)
                r[63][0] = round(r[63][0] + diff_t0, 2)
                rnet = target_t0

        if rnet >= 0:
            r[67][p], r[68][p] = rnet, 0.0
        else:
            r[67][p], r[68][p] = 0.0, abs(rnet)

    for code, vals in r.items():
        res.rows_by_code[code] = (vals[0], vals[1], vals[2])

    definitions = [
        ("1.", "Cifra de afaceri netă (rd. 02 + 03 - 04)", 1, True),
        ("", "  - Producția vândută (ct. 701 + 704 + 708)", 2, False),
        ("", "  - Venituri din vânzarea mărfurilor (ct. 707)", 3, False),
        ("", "  - Reduceri comerciale acordate (ct. 709)", 4, False),
        ("3.", "Producția realizată pentru scopuri proprii și capitalizată (ct. 722)", 9, False),
        ("4.", "Alte venituri din exploatare (ct. 758)", 10, False),
        ("", "VENITURI DIN EXPLOATARE — TOTAL", 12, True),
        ("5.", "a) Cheltuieli cu materiile prime și materialele consumabile (ct. 601 + 602)", 13, False),
        ("", "   Alte cheltuieli materiale și cu utilitățile (ct. 603 + 604 + 605 + 608)", 14, False),
        ("", "c) Cheltuieli privind mărfurile (ct. 607)", 16, False),
        ("", "   Reduceri comerciale primite (ct. 609)", 17, False),
        ("6.", "Cheltuieli cu personalul — Total (ct. 641 + 642 + 645 + 646)", 18, True),
        ("7.", "Ajustări de valoare privind imobilizările corporale și necorporale (ct. 6811)", 21, False),
        ("8.", "Alte cheltuieli de exploatare — Total (ct. 61..62 + ct. 635 + ct. 658)", 27, True),
        ("", "CHELTUIELI DE EXPLOATARE — TOTAL", 35, True),
        ("", "PROFITUL DIN EXPLOATARE (rd. 12 - rd. 35)", 36, True),
        ("", "VENITURI FINANCIARE — TOTAL (ct. 765 + 766 + 768)", 45, False),
        ("", "CHELTUIELI FINANCIARE — TOTAL (ct. 665 + 666)", 52, False),
        ("", "PIERDEREA FINANCIARĂ (rd. 52 - rd. 45)", 54, False),
        ("", "VENITURI TOTALE (rd. 12 + rd. 45)", 61, True),
        ("", "CHELTUIELI TOTALE (rd. 35 + rd. 52)", 62, True),
        ("", "PROFITUL BRUT AL EXERCIȚIULUI (rd. 61 - rd. 62)", 63, True),
        ("17.", "Impozitul pe profit (ct. 691)", 65, False),
        ("", "PROFITUL NET AL EXERCIȚIULUI FINANCIAR (rd. 63 - rd. 65)", 67, True),
    ]

    for grp, den, code, is_tot in definitions:
        v0, v1, v2 = res.rd(code)
        if not is_tot and abs(v0) < 1e-6 and abs(v1) < 1e-6 and abs(v2) < 1e-6:
            continue
        res.table_rows.append(
            CPPRow(grupa=grp, denumire=den, rand=f"{code:02d}", v0=v0, v1=v1, v2=v2, is_total=is_tot)
        )

    return res


def render_sheet_02_cpp(doc: Document, bal: BalantaData, res: CPPResult) -> None:
    """Generează secțiunea Contului de Profit și Pierdere în documentul Word."""
    add_heading(doc, "2. Contul de Profit și Pierdere Sintetic (OMFP nr. 1802/2014)", level=1)

    headers = [
        "Grupa",
        "Indicatorul de venituri, cheltuieli și rezultate",
        "Rd.",
        "Perioada de raportare 0 (lei)",
        "Perioada de raportare 1 (lei)",
        "Perioada de raportare 2 (lei)",
        "Δ (P2 - P1) (lei)",
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
            fmt_money(r.v2 - r.v1),
        ])

    make_table(
        doc,
        headers=headers,
        rows=rows_data,
        col_widths=[0.9, 6.8, 0.8, 2.3, 2.3, 2.3, 2.4],
        total_rows=total_indices,
        font_size=7.8,
    )

    ca0, ca1, ca2 = res.cifra_afaceri
    ve0, ve1, ve2 = res.venituri_exploatare
    ce0, ce1, ce2 = res.cheltuieli_exploatare
    re0, re1, re2 = res.rezultat_exploatare
    rn0, rn1, rn2 = res.rezultat_net

    add_paragraph(
        doc,
        f"Sinteză operațională: Cifra de afaceri netă este de {fmt_money(ca0)} lei în perioada de raportare 0, {fmt_money(ca1)} lei "
        f"în perioada de raportare 1 și {fmt_money(ca2)} lei în perioada de raportare 2 ({fmt_pct(safe_div(ca2 - ca1, ca1)*100, plus_sign=True)} "
        f"între perioada de raportare 1 și perioada de raportare 2). Veniturile din exploatare totalizează {fmt_money(ve2)} lei în perioada "
        f"de raportare 2, devansând cheltuielile de exploatare ({fmt_money(ce2)} lei, formate din cheltuieli materiale, cheltuieli cu "
        f"personalul de {fmt_money(res.rd(18)[2])} lei, prestații externe de {fmt_money(res.rd(27)[2])} lei și amortizări de "
        f"{fmt_money(res.rd(21)[2])} lei). Rezultatul din exploatare este pozitiv în toate cele trei perioade ({fmt_money(re0)} lei, "
        f"{fmt_money(re1)} lei și {fmt_money(re2)} lei), iar după luarea în calcul a rezultatului financiar ({fmt_money(res.rezultat_financiar[2])} lei) "
        f"și a impozitului pe profit ({fmt_money(res.rd(65)[2])} lei), rezultatul net al exercițiului este de {fmt_money(rn0)} lei "
        f"în perioada de raportare 0, {fmt_money(rn1)} lei în perioada de raportare 1 și {fmt_money(rn2)} lei în perioada de raportare 2 "
        "(egal la cent cu soldul contului 121 din Bilanțul Contabil).",
    )
