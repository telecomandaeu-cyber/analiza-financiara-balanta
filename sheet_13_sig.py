#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 13: Analiza Soldurilor Intermediare de Gestiune (Cascada SIG).

Cascada: Marja comercială → Producția exercițiului → Valoarea adăugată → EBE →
Rezultatul exploatării → Rezultatul curent → Rezultatul net. Toate soldurile sunt rulaje
cumulate de la începutul anului (P1, P2); P0 nu conține rulaje (n.d.).
Rezultatul exploatării și rezultatul net obținute prin cascadă sunt verificate față de CPP.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from docx import Document

from core_balanta import (
    BalantaData,
    add_callout,
    add_heading,
    add_paragraph,
    calc_delta,
    calc_growth_rate,
    calc_index,
    calc_ritm_mediu,
    describe_flow,
    fmt_money,
    fmt_ratio_pct,
    make_table,
    pick_variant,
    safe_div,
)
from sheet_02_cpp import CPPResult

Triple = Tuple[float, float, float]


@dataclass
class SIGRowData:
    name: str
    vals: Triple
    delta: Tuple[float, float]
    index: Tuple[Optional[float], Optional[float]]
    growth: Tuple[Optional[float], Optional[float]]
    ritm_mediu: Optional[float]
    is_marja: bool = False


@dataclass
class SIGResult:
    rows: List[SIGRowData] = field(default_factory=list)
    by_key: Dict[str, Triple] = field(default_factory=dict)
    check_re: Tuple[float, float] = (0.0, 0.0)
    check_rn: Tuple[float, float] = (0.0, 0.0)

    def _g(self, k: str) -> Triple:
        return self.by_key.get(k, (0.0, 0.0, 0.0))

    @property
    def marja_comerciala(self) -> Triple:
        return self._g("mc")

    @property
    def productia_exercitiului(self) -> Triple:
        return self._g("qe")

    @property
    def consumuri_terti(self) -> Triple:
        return self._g("ct")

    @property
    def valoarea_adaugata(self) -> Triple:
        return self._g("va")

    @property
    def ebe(self) -> Triple:
        return self._g("ebe")

    @property
    def rezultat_exploatare(self) -> Triple:
        return self._g("re")

    @property
    def rezultat_curent(self) -> Triple:
        return self._g("rc")

    @property
    def rezultat_net(self) -> Triple:
        return self._g("rn")


def _t(fn) -> Triple:
    return tuple(round(fn(p), 2) for p in range(3))  # type: ignore[return-value]


def compute_sheet_13_sig(cpp: CPPResult) -> SIGResult:
    """Calculează cascada Soldurilor Intermediare de Gestiune (SIG)."""
    res = SIGResult()
    g = lambda k, p: cpp.it(k)[p]  # noqa: E731

    vanz_marf = _t(lambda p: g("venit_marfuri", p))
    cost_marf = _t(lambda p: g("marfuri", p))
    mc = _t(lambda p: vanz_marf[p] - cost_marf[p])

    prod_vand = _t(lambda p: g("prod_vanduta", p) - g("reduceri_acordate", p))
    prod_stoc = _t(lambda p: g("var_stoc", p))
    prod_imob = _t(lambda p: g("prod_imob", p))
    qe = _t(lambda p: prod_vand[p] + prod_stoc[p] + prod_imob[p])

    # Consumurile de la terți: materiale, energie, prestații externe, minus reducerile comerciale primite (ct. 609).
    cons_terti = _t(
        lambda p: g("mat_prime", p) + g("alte_mat", p) + g("energie", p) + g("prest_ext", p) - g("reduceri_primite", p)
    )
    va = _t(lambda p: mc[p] + qe[p] - cons_terti[p])

    subv_expl = _t(lambda p: g("subv_ca", p) + g("subv_expl", p))
    imp_taxe = _t(lambda p: g("impozite_taxe", p))
    ch_pers = _t(lambda p: g("salarii", p) + g("asigurari", p))
    ebe = _t(lambda p: va[p] + subv_expl[p] - imp_taxe[p] - ch_pers[p])

    alte_ven_expl = _t(
        lambda p: g("alte_ven_expl", p) + g("ven_cedare", p) + g("ven_subv_inv", p) + g("ven_fc_neg", p)
    )
    reluari_expl = _t(
        lambda p: g("rel_ajust_imob", p) + g("rel_ajust_ac", p) + g("creante_reactivate", p) + g("rel_prov", p)
    )
    amort_calc = _t(
        lambda p: g("amortizare", p) + g("ajust_ac_ch", p) + g("pierderi_creante", p) + g("ajust_prov", p)
    )
    alte_ch_expl = _t(lambda p: g("alte_ch_expl", p) + g("vnc_cedate", p))
    re = _t(lambda p: ebe[p] + alte_ven_expl[p] + reluari_expl[p] - amort_calc[p] - alte_ch_expl[p])

    ven_fin = cpp.venituri_financiare
    ch_fin = cpp.cheltuieli_financiare
    rc = _t(lambda p: re[p] + ven_fin[p] - ch_fin[p])

    rez_extr = _t(lambda p: g("ven_extra", p) - g("ch_extra", p))
    imp_prof = cpp.impozit_profit
    rn = _t(lambda p: rc[p] + rez_extr[p] - imp_prof[p])

    res.by_key = {"mc": mc, "qe": qe, "ct": cons_terti, "va": va, "ebe": ebe, "re": re, "rc": rc, "rn": rn}
    re_cpp, rn_cpp = cpp.rezultat_exploatare, cpp.rezultat_net
    res.check_re = (round(re[1] - re_cpp[1], 2), round(re[2] - re_cpp[2], 2))
    res.check_rn = (round(rn[1] - rn_cpp[1], 2), round(rn[2] - rn_cpp[2], 2))

    raw_rows = [
        ("Vânzări de mărfuri (ct. 707)", vanz_marf, False),
        ("- Costul mărfurilor vândute (ct. 607)", cost_marf, False),
        ("= Marja comercială", mc, True),
        ("Producția vândută netă de reduceri (ct. 70x - 709)", prod_vand, False),
        ("± Producția stocată (variația stocurilor de produse)", prod_stoc, False),
        ("+ Producția imobilizată (ct. 72x)", prod_imob, False),
        ("= Producția exercițiului", qe, True),
        ("- Consumuri de la terți (ct. 60x fără 607, 61x, 62x, minus 609)", cons_terti, False),
        ("= Valoarea adăugată (MC + QE - consumuri de la terți)", va, True),
        ("+ Subvenții de exploatare (ct. 741x)", subv_expl, False),
        ("- Impozite, taxe și vărsăminte asimilate (ct. 63x)", imp_taxe, False),
        ("- Cheltuieli cu personalul (ct. 64x)", ch_pers, False),
        ("= Excedentul brut de exploatare (EBE)", ebe, True),
        ("+ Alte venituri din exploatare (ct. 75x fără 754, 7815)", alte_ven_expl, False),
        ("+ Reluări de ajustări și provizioane, creanțe reactivate (ct. 781x, 754)", reluari_expl, False),
        ("- Amortizări, ajustări, provizioane și pierderi din creanțe (ct. 681x, 654)", amort_calc, False),
        ("- Alte cheltuieli de exploatare (ct. 65x fără 654)", alte_ch_expl, False),
        ("= Rezultatul exploatării", re, True),
        ("+ Venituri financiare", ven_fin, False),
        ("- Cheltuieli financiare", ch_fin, False),
        ("= Rezultatul curent (brut)", rc, True),
        ("± Rezultatul extraordinar (ct. 77x - 67x)", rez_extr, False),
        ("- Impozitul pe profit / venit (ct. 69x)", imp_prof, False),
        ("= Rezultatul net al perioadei", rn, True),
    ]
    for name, vals, is_m in raw_rows:
        res.rows.append(
            SIGRowData(
                name=name,
                vals=vals,
                delta=calc_delta(vals),
                index=calc_index(vals),
                growth=calc_growth_rate(vals),
                ritm_mediu=calc_ritm_mediu(vals),
                is_marja=is_m,
            )
        )
    return res


def render_sheet_13_sig(doc: Document, bal: BalantaData, res: SIGResult) -> None:
    """Generează secțiunea Soldurilor Intermediare de Gestiune (SIG)."""
    add_heading(doc, "13. Analiza Soldurilor Intermediare de Gestiune (Cascada SIG)", level=1)

    fl = bal.period.flow_labels
    luna = bal.period.month_name or "Luna curentă"
    m1 = bal.period.months[1]
    headers = [
        "Element / Sold Intermediar de Gestiune",
        f"P1 cumulat {fl[1]}",
        f"P2 cumulat {fl[2]}",
        f"Rulaj {luna} (P2 - P1)",
        "Luna curentă vs. media lunară P1",
    ]
    rows_table = []
    tot_idx = []
    for r in res.rows:
        if not r.is_marja and abs(r.vals[1]) < 1e-6 and abs(r.vals[2]) < 1e-6:
            continue
        if r.is_marja:
            tot_idx.append(len(rows_table))
        luna_val = r.vals[2] - r.vals[1]
        avg1 = r.vals[1] / m1 if m1 > 0 else 0.0
        cmp = safe_div(luna_val - avg1, abs(avg1), default=None) if abs(avg1) > 1e-6 else None
        rows_table.append([
            r.name, fmt_money(r.vals[1]), fmt_money(r.vals[2]), fmt_money(luna_val), fmt_ratio_pct(cmp, plus_sign=True),
        ])

    make_table(
        doc,
        headers=headers,
        rows=rows_table,
        col_widths=[6.6, 2.5, 2.5, 2.4, 2.2],
        total_rows=tot_idx,
        font_size=7.6,
    )

    ok = all(abs(v) <= 0.05 for v in res.check_re + res.check_rn)
    add_callout(
        doc,
        (
            "Rezultatul exploatării și rezultatul net obținute prin cascada SIG sunt identice cu cele din Contul de profit și pierdere în P1 și P2."
            if ok
            else f"ATENȚIE: cascada SIG diferă de CPP (diferențe RE: {fmt_money(res.check_re[0])} / {fmt_money(res.check_re[1])} lei; RN: {fmt_money(res.check_rn[0])} / {fmt_money(res.check_rn[1])} lei)."
        ),
        title="Verificarea cascadei",
    )

    mc2 = res.marja_comerciala[2]
    pe2 = res.productia_exercitiului[2]
    va2 = res.valoarea_adaugata[2]
    ebe2 = res.ebe[2]
    re2 = res.rezultat_exploatare[2]
    rc2 = res.rezultat_curent[2]
    rn2 = res.rezultat_net[2]
    flow_va = describe_flow(res.valoarea_adaugata, bal, "valoarea adăugată")
    flow_ebe = describe_flow(res.ebe, bal, "excedentul brut de exploatare")

    if mc2 > pe2:
        motor_sig = f"formarea marjelor este susținută preponderent de activitatea comercială (marja comercială = {fmt_money(mc2)} lei, față de producția exercițiului = {fmt_money(pe2)} lei)"
    elif pe2 > 0 and mc2 > 0:
        motor_sig = f"baza generatoare de valoare îmbină producția exercițiului ({fmt_money(pe2)} lei) cu marja comercială ({fmt_money(mc2)} lei)"
    else:
        motor_sig = f"activitatea de bază se sprijină pe producția exercițiului ({fmt_money(pe2)} lei)"

    rata_va = safe_div(ebe2, va2, default=None) if abs(va2) > 1e-6 else None
    s13_end = pick_variant(
        bal,
        "s13_end",
        [
            f"EBE reține {fmt_ratio_pct(rata_va)} din valoarea adăugată; la treptele finale ale cascadei din perioada de raportare 2, rezultatul exploatării este de {fmt_money(re2)} lei, rezultatul curent de {fmt_money(rc2)} lei, iar rezultatul net de {fmt_money(rn2)} lei.",
            f"După amortizări, ajustări și fluxurile financiare și fiscale, cascada perioadei de raportare 2 se închide cu un rezultat al exploatării de {fmt_money(re2)} lei, un rezultat curent de {fmt_money(rc2)} lei și un rezultat net de {fmt_money(rn2)} lei (EBE = {fmt_money(ebe2)} lei, {fmt_ratio_pct(rata_va)} din VA).",
        ],
    )
    add_paragraph(
        doc,
        f"Interpretarea cascadei SIG: în perioada de raportare 2, {motor_sig}; după consumurile de la terți "
        f"({fmt_money(res.consumuri_terti[2])} lei), {flow_va}, iar {flow_ebe}. {s13_end}",
    )
