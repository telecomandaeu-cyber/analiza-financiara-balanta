#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 1: Bilanțul Contabil Sintetic (structura OMFP nr. 1802/2014).

Soldurile sunt preluate din balanța de verificare pentru cele trei momente de analiză:
  - Perioada de raportare 0: soldurile inițiale ale exercițiului (01.01);
  - Perioada de raportare 1: soldurile la finele lunii precedente (TS - rulajul lunii);
  - Perioada de raportare 2: soldurile finale ale lunii de raportare.

Fiecare cont din balanță este repartizat pe un rând de bilanț după natura lui
(clasă, grupă, cont sintetic), iar conturile bifuncționale sunt împărțite pe analitice
(ex. 4424 TVA de recuperat → creanțe, 4428 TVA neexigibilă cu sold creditor → datorii).
Astfel nu se pierde niciun sold, iar egalitatea Activ = Pasiv este verificată explicit.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from docx import Document

from core_balanta import (
    BalantaData,
    add_callout,
    add_heading,
    add_paragraph,
    describe_trajectory,
    fmt_money,
    fmt_ratio_pct,
    make_table,
    pick_variant,
    safe_div,
)

TOL = 0.05  # toleranța de rotunjire (lei)


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
    # Agregate suplimentare folosite de capitolele următoare
    credite_trezorerie: Tuple[float, float, float] = (0.0, 0.0, 0.0)  # 519 + descoperit 512
    imprumuturi_asociati: Tuple[float, float, float] = (0.0, 0.0, 0.0)  # 455 (sold creditor)
    rezerve_si_rezultate: Tuple[float, float, float] = (0.0, 0.0, 0.0)  # rezerve + rezultat reportat + rezultatul exercițiului - repartizare
    total_activ: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    total_pasiv: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    diferente: Tuple[float, float, float] = (0.0, 0.0, 0.0)
    note: List[str] = field(default_factory=list)
    conturi_nemapate: List[str] = field(default_factory=list)

    def rd(self, code: int) -> Tuple[float, float, float]:
        return self.rows_by_code.get(code, (0.0, 0.0, 0.0))

    @property
    def echilibrat(self) -> bool:
        return all(abs(d) <= TOL for d in self.diferente)


# ------------------------------------------------------------------------------
# Repartizarea conturilor pe rândurile bilanțului
# ------------------------------------------------------------------------------
ASSET_LINES = {1, 2, 3, 5, 6, 7, 8, 10}
LIAB_LINES = {11, 14, 15, 17, 18, 21, 22}

_NECORP = ("201", "203", "205", "206", "207", "208", "233", "234", "280", "290")
_FIN = ("261", "262", "263", "264", "265", "266", "267", "296")
_INV_TS = ("501", "502", "503", "505", "506", "507", "508", "590", "591", "592", "593", "595", "596", "598")
_DISP = ("511", "512", "531", "532", "541", "542", "581")


def _classify(code: str, nd: float) -> Tuple[str, str]:
    """
    Returnează (destinație, observație) pentru soldul net debitor `nd` al contului `code`.
    Destinația este un rând de bilanț ("L<n>"), o componentă de capital ("EQ:<cheie>"),
    "UNCLOSED" (conturi de venituri/cheltuieli neînchise) sau "IGNORE" (clasele 8–9).
    """
    c = code
    c3 = c[:3]
    cls = c[0]
    if cls in "089":
        return "IGNORE", ""
    if cls in "67":
        return "UNCLOSED", ""

    if cls == "1":
        if c3 == "101":
            return {
                "1011": "EQ:cap_nevarsat",
                "1015": "EQ:cap_regie",
                "1016": "EQ:patrimoniu_public",
                "1017": "EQ:cap_alte",
                "1018": "EQ:cap_alte",
            }.get(c, "EQ:cap_varsat"), ""
        if c3 == "104":
            return "EQ:prime", ""
        if c3 == "105":
            return "EQ:rezerve_reev", ""
        if c3 in ("106", "107", "108"):
            return "EQ:rezerve", ""
        if c3 == "109":
            return "EQ:actiuni_proprii", ""
        if c3 == "117":
            return "EQ:reportat", ""
        if c3 == "121":
            return "EQ:rezultat", ""
        if c3 == "129":
            return "EQ:repartizare", ""
        if c3 == "141":
            return "EQ:castiguri", ""
        if c3 == "149":
            return "EQ:pierderi", ""
        if c3 == "151":
            return "L15", ""
        if c[:2] == "16":
            return "L14", "dtl"
        return "EQ:alte", f"contul {c} (clasa 1) a fost inclus în capitalurile proprii"

    if cls == "2":
        if c.startswith("2075"):
            return "L22", ""
        if c.startswith("2933") or c.startswith("2903"):
            return "L1", ""
        if c.startswith("4094"):
            return "L1", ""
        if c3 in _NECORP:
            return "L1", ""
        if c3 in _FIN:
            return "L3", ""
        if c3 == "269":
            return ("L11", "") if nd <= 0 else ("L3", "")
        return "L2", ""

    if cls == "3":
        return "L5", ""

    if cls == "4":
        if c.startswith("4093"):
            return "L2", ""
        if c.startswith("4094"):
            return "L1", ""
        if c.startswith("4092"):
            return ("L6", "") if nd >= 0 else ("L11", "")
        if c3 == "409":  # 4091 sau 409 fără analitice: avansuri pentru stocuri
            return ("L5", "") if nd >= 0 else ("L11", "")
        if c3 == "471":
            return ("L10", "") if nd >= 0 else ("L18", "")
        if c3 == "472":
            return ("L18", "") if nd <= 0 else ("L6", "")
        if c3 == "475":
            return "L17", ""
        if c3 == "478":
            return "L21", ""
        if c3 in ("490", "491", "495", "496"):
            return "L6", ""
        if nd >= 0:
            return "L6", ""
        return "L11", ("asociati" if c3 == "455" else "")

    if cls == "5":
        if c3 in _INV_TS:
            return "L7", ""
        if c3 == "509":
            return ("L11", "") if nd <= 0 else ("L7", "")
        if c3 == "519":
            return ("L11", "trez") if nd <= 0 else ("L6", "")
        if c3 == "518":
            return ("L6", "") if nd >= 0 else ("L11", "")
        if c3 in _DISP:
            if nd >= 0:
                return "L8", ""
            return "L11", ("trez" if c3 == "512" else "")
        return ("L8", "") if nd >= 0 else ("L11", "")

    return "IGNORE", f"contul {c} nu a putut fi încadrat"


def _current_year_result(bal: BalantaData) -> Tuple[float, float, float]:
    """Rezultatul net al exercițiului curent din Contul de profit și pierdere (import întârziat)."""
    from sheet_02_cpp import compute_net_result  # evită importul circular

    return compute_net_result(bal)


def compute_sheet_01_bilant(bal: BalantaData) -> BilantResult:
    """Calculează rândurile Bilanțului Contabil (OMFP nr. 1802/2014) pentru P0, P1 și P2."""
    res = BilantResult()
    r: Dict[int, List[float]] = {k: [0.0, 0.0, 0.0] for k in range(1, 42)}
    eq_keys = (
        "cap_varsat", "cap_nevarsat", "cap_regie", "cap_alte", "patrimoniu_public", "prime", "rezerve_reev",
        "rezerve", "actiuni_proprii", "reportat", "rezultat", "repartizare", "castiguri", "pierderi", "alte",
    )
    eq: Dict[str, List[float]] = {k: [0.0, 0.0, 0.0] for k in eq_keys}
    unclosed = [0.0, 0.0, 0.0]
    trez = [0.0, 0.0, 0.0]
    asoc = [0.0, 0.0, 0.0]
    notes: List[str] = []

    for code in bal.leaf_accounts():
        for p in (0, 1, 2):
            nd = bal.net_d(code, p)
            if abs(nd) < 0.005:
                continue
            dest, obs = _classify(code, nd)
            if dest == "IGNORE":
                if obs and obs not in notes:
                    notes.append(obs)
                    res.conturi_nemapate.append(code)
                continue
            if dest == "UNCLOSED":
                unclosed[p] += -nd
                continue
            if dest.startswith("EQ:"):
                eq[dest[3:]][p] += -nd
                if obs and obs not in notes:
                    notes.append(obs)
                    res.conturi_nemapate.append(code)
                continue
            line = int(dest[1:])
            amount = nd if line in ASSET_LINES else -nd
            r[line][p] += amount
            if obs == "trez":
                trez[p] += -nd
            elif obs == "asociati":
                asoc[p] += -nd

    # Rezultatul exercițiului: soldul 121 + conturile de venituri/cheltuieli neînchise.
    # Pentru P1 și P2, rezultatul curent se ia din CPP; eventuala diferență față de soldul
    # contului 121 (de regulă rezultatul exercițiului precedent încă nerepartizat) se
    # prezintă la rezultatul reportat.
    rn_cpp = _current_year_result(bal)
    for p in (0, 1, 2):
        total_rez = eq["rezultat"][p] + unclosed[p]
        if p > 0 and abs(unclosed[p]) > TOL:
            notes.append(
                f"În perioada de raportare {p}, conturile de venituri și cheltuieli nu erau închise "
                f"(rezultat neînchis de {fmt_money(unclosed[p])} lei) și au fost incluse în rezultatul exercițiului."
            )
        if p > 0 and abs(total_rez - rn_cpp[p]) > TOL:
            dif = total_rez - rn_cpp[p]
            eq["reportat"][p] += dif
            total_rez = rn_cpp[p]
            notes.append(
                f"În perioada de raportare {p}, soldul contului 121 ({fmt_money(eq['rezultat'][p] + unclosed[p])} lei) diferă de "
                f"rezultatul net din CPP ({fmt_money(rn_cpp[p])} lei); diferența de {fmt_money(dif)} lei "
                f"(rezultat al exercițiului precedent nerepartizat) a fost prezentată la rezultatul reportat."
            )
        eq["rezultat"][p] = total_rez

    for p in (0, 1, 2):
        r[4][p] = r[1][p] + r[2][p] + r[3][p]
        r[9][p] = r[5][p] + r[6][p] + r[7][p] + r[8][p]
        r[19][p] = r[18][p]  # veniturile în avans (472) se reiau, de regulă, în mai puțin de un an
        r[20][p] = r[17][p] + r[21][p] + r[22][p]
        r[16][p] = r[17][p] + r[18][p] + r[21][p] + r[22][p]
        r[12][p] = r[9][p] + r[10][p] - r[11][p] - r[19][p]
        r[13][p] = r[4][p] + r[12][p]

        r[24][p] = eq["cap_varsat"][p]
        r[25][p] = eq["cap_nevarsat"][p]
        r[26][p] = eq["cap_regie"][p]
        r[27][p] = eq["cap_alte"][p] + eq["alte"][p]
        r[23][p] = r[24][p] + r[25][p] + r[26][p] + r[27][p]
        r[28][p] = eq["prime"][p]
        r[29][p] = eq["rezerve_reev"][p]
        r[30][p] = eq["rezerve"][p]
        r[31][p] = -eq["actiuni_proprii"][p]
        r[32][p] = eq["castiguri"][p]
        r[33][p] = -eq["pierderi"][p]
        rep = eq["reportat"][p]
        r[34][p], r[35][p] = (rep, 0.0) if rep >= 0 else (0.0, -rep)
        rez = eq["rezultat"][p]
        r[36][p], r[37][p] = (rez, 0.0) if rez >= 0 else (0.0, -rez)
        r[38][p] = -eq["repartizare"][p]
        r[39][p] = (
            r[23][p] + r[28][p] + r[29][p] + r[30][p] - r[31][p] + r[32][p] - r[33][p]
            + r[34][p] - r[35][p] + r[36][p] - r[37][p] - r[38][p]
        )
        r[40][p] = eq["patrimoniu_public"][p]
        r[41][p] = r[39][p] + r[40][p]

    for code in r:
        r[code] = [round(v, 2) for v in r[code]]
        res.rows_by_code[code] = (r[code][0], r[code][1], r[code][2])

    ta = tuple(round(r[4][p] + r[9][p] + r[10][p], 2) for p in range(3))
    tp = tuple(round(r[11][p] + r[14][p] + r[15][p] + r[16][p] + r[41][p], 2) for p in range(3))
    res.total_activ = ta
    res.total_pasiv = tp
    res.diferente = tuple(round(ta[p] - tp[p], 2) for p in range(3))
    res.credite_trezorerie = tuple(round(v, 2) for v in trez)
    res.imprumuturi_asociati = tuple(round(v, 2) for v in asoc)
    res.rezerve_si_rezultate = tuple(
        round(r[30][p] + r[34][p] - r[35][p] + r[36][p] - r[37][p] - r[38][p], 2) for p in range(3)
    )
    for p in range(3):
        if abs(res.diferente[p]) > TOL:
            notes.append(
                f"ATENȚIE: în perioada de raportare {p} totalul activului diferă de totalul pasivului cu "
                f"{fmt_money(res.diferente[p])} lei — verificați balanța și conturile nemapate."
            )
    res.note = notes + list(bal.warnings)

    definitions = [
        ("A.", "I. Imobilizări necorporale", "(ct. 201..208, 233, 234 - 280 - 290 - 2933)", 1, False),
        ("", "II. Imobilizări corporale", "(ct. 211..235 - 281 - 291 - 2931)", 2, False),
        ("", "III. Imobilizări financiare", "(ct. 261..267 - 296)", 3, False),
        ("", "ACTIVE IMOBILIZATE — TOTAL", "(rd. 01 + 02 + 03)", 4, True),
        ("B.", "I. Stocuri", "(ct. 301..381 ± 308..388 - 39x + 4091)", 5, False),
        ("", "II. Creanțe", "(solduri debitoare clasa 4, ct. 5187 - 49x)", 6, False),
        ("", "III. Investiții pe termen scurt", "(ct. 501..508 - 59x)", 7, False),
        ("", "IV. Casa și conturi la bănci", "(ct. 511, 512, 531, 532, 541, 542, 581)", 8, False),
        ("", "ACTIVE CIRCULANTE — TOTAL", "(rd. 05 + 06 + 07 + 08)", 9, True),
        ("C.", "Cheltuieli în avans", "(ct. 471)", 10, False),
        ("D.", "Datorii ce trebuie plătite într-o perioadă de până la un an", "(solduri creditoare clasa 4, ct. 269, 509, 519, 5186)", 11, True),
        ("E.", "Active circulante nete / Datorii curente nete", "(rd. 09 + 10 - 11 - 19)", 12, False),
        ("F.", "TOTAL ACTIVE MINUS DATORII CURENTE", "(rd. 04 + 12)", 13, True),
        ("G.", "Datorii ce trebuie plătite într-o perioadă mai mare de un an", "(ct. 161..168 - 169)", 14, True),
        ("H.", "Provizioane", "(ct. 151)", 15, False),
        ("I.", "Venituri în avans — total", "(ct. 472 + 475 + 478 + 2075)", 16, True),
        ("J.", "I. Capital subscris vărsat", "(ct. 1012)", 24, False),
        ("", "   Capital subscris nevărsat", "(ct. 1011)", 25, False),
        ("", "II. Prime de capital", "(ct. 104)", 28, False),
        ("", "III. Rezerve din reevaluare", "(ct. 105)", 29, False),
        ("", "IV. Rezerve", "(ct. 106)", 30, False),
        ("", "   Acțiuni proprii", "(ct. 109)", 31, False),
        ("", "V. Profitul reportat — sold C", "(ct. 117)", 34, False),
        ("", "   Pierderea reportată — sold D", "(ct. 117)", 35, False),
        ("", "VI. Profitul perioadei — sold C", "(ct. 121)", 36, False),
        ("", "   Pierderea perioadei — sold D", "(ct. 121)", 37, False),
        ("", "   Repartizarea profitului", "(ct. 129)", 38, False),
        ("", "CAPITALURI PROPRII — TOTAL", "(rd. 23 + 28 + 29 + 30 - 31 + 32 - 33 ± 34/35 ± 36/37 - 38)", 39, True),
        ("", "Patrimoniul public", "(ct. 1016)", 40, False),
        ("", "CAPITALURI — TOTAL", "(rd. 39 + 40)", 41, True),
    ]

    for grp, den, form_str, code, is_tot in definitions:
        v0, v1, v2 = res.rd(code)
        if not is_tot and abs(v0) < 1e-6 and abs(v1) < 1e-6 and abs(v2) < 1e-6:
            continue
        if code == 41 and all(abs(v) < 1e-6 for v in res.rd(40)):
            continue  # identic cu rd. 39
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
    for lbl, vals in (("TOTAL ACTIV (rd. 04 + 09 + 10)", ta), ("TOTAL PASIV (rd. 11 + 14 + 15 + 16 + 41)", tp)):
        res.table_rows.append(BilantRow("", lbl, "", "—", vals[0], vals[1], vals[2], True))

    return res


def render_sheet_01_bilant(doc: Document, bal: BalantaData, res: BilantResult) -> None:
    """Generează secțiunea Bilanțului Contabil Sintetic în documentul Word."""
    add_heading(doc, "1. Bilanțul Contabil Sintetic (OMFP nr. 1802/2014)", level=1)

    c0, c1, c2 = bal.col_labels
    headers = ["Grupa", "Element patrimonial și corespondență contabilă", "Rd.", c0, c1, c2, "Δ (P2 - P0)"]
    rows_data = []
    total_indices = []
    for idx, r in enumerate(res.table_rows):
        if r.is_total:
            total_indices.append(idx)
        rows_data.append([
            r.grupa, r.denumire, r.rand, fmt_money(r.v0), fmt_money(r.v1), fmt_money(r.v2), fmt_money(r.v2 - r.v0),
        ])

    make_table(
        doc,
        headers=headers,
        rows=rows_data,
        col_widths=[0.9, 6.8, 0.8, 2.3, 2.3, 2.3, 2.4],
        total_rows=total_indices,
        font_size=7.8,
    )

    if res.echilibrat:
        verif = (
            "Activul total este egal cu pasivul total în toate cele trei perioade "
            f"({fmt_money(res.total_activ[0])} lei, {fmt_money(res.total_activ[1])} lei și {fmt_money(res.total_activ[2])} lei); "
            "toate soldurile din balanță au fost repartizate pe rândurile bilanțului."
        )
    else:
        verif = "Bilanțul NU se închide în toate perioadele: " + "; ".join(
            f"P{p}: diferență {fmt_money(res.diferente[p])} lei" for p in range(3)
        ) + "."
    extra = (" " + " ".join(res.note)) if res.note else ""
    add_callout(doc, verif + extra, title="Verificarea bilanțului")

    ai0, ai1, ai2 = res.rd(4)
    ac2 = res.rd(9)[2]
    ta0, ta1, ta2 = res.total_activ
    dts2 = res.rd(11)[2]
    dtl2 = res.rd(14)[2]
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
                f"Structura patrimonială din perioada de raportare 2 este preponderent imobilizată (activele fixe nete de {fmt_money(ai2)} lei dețin {fmt_ratio_pct(pond_ai2)} din activ), completată de active circulante de {fmt_money(ac2)} lei.",
                f"Arhitectura activului în perioada de raportare 2 evidențiază dominanța imobilizărilor nete ({fmt_money(ai2)} lei, respectiv {fmt_ratio_pct(pond_ai2)} din patrimoniu), în timp ce activele curente însumează {fmt_money(ac2)} lei (din care trezorerie {fmt_money(disp2)} lei).",
            ],
        )
    else:
        dom_ac = "disponibilitățile bănești" if disp2 >= max(stoc2, creante2) else ("creanțele" if creante2 >= stoc2 else "stocurile")
        profil_activ = (
            f"În perioada de raportare 2, activele circulante ({fmt_money(ac2)} lei, componenta principală fiind {dom_ac}) "
            f"dețin {fmt_ratio_pct(safe_div(ac2, ta2))} din activ, iar imobilizările nete însumează {fmt_money(ai2)} lei ({fmt_ratio_pct(pond_ai2)})."
        )

    if cp2 < 0:
        pasiv_msg = (
            f"Pe latura resurselor, {traj_cp}; valoarea negativă a capitalurilor proprii indică un activ net negativ, "
            f"finanțarea fiind asigurată integral din datorii curente ({fmt_money(dts2)} lei) și pe termen lung ({fmt_money(dtl2)} lei)."
        )
    elif dtl2 > 0:
        pasiv_msg = (
            f"În planul finanțării, {traj_cp}, iar obligațiile față de terți în perioada de raportare 2 se împart între "
            f"datorii exigibile sub un an ({fmt_money(dts2)} lei) și datorii peste un an ({fmt_money(dtl2)} lei)."
        )
    else:
        pasiv_msg = (
            f"În planul resurselor, {traj_cp}; societatea nu are datorii peste un an în perioada de raportare 2, "
            f"obligațiile față de terți fiind integral pe termen scurt ({fmt_money(dts2)} lei)."
        )

    add_paragraph(doc, f"Diagnostic bilanțier: {traj_ta[0].upper() + traj_ta[1:]}. {profil_activ} {pasiv_msg}")
