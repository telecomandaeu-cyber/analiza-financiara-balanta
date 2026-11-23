#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 14: Capacitatea de Autofinanțare (CAF — Metoda Deductivă și Metoda Aditivă).

Metoda deductivă (pornind de la EBE):
  CAF = EBE + alte venituri de exploatare încasabile - alte cheltuieli de exploatare plătibile
        + venituri financiare încasabile - cheltuieli financiare plătibile
        ± rezultat extraordinar - impozitul pe profit
  (se exclud veniturile din cedarea activelor 7583, subvențiile pentru investiții reluate 7584,
   fondul comercial negativ 7815 și valoarea contabilă a activelor cedate 6583)

Metoda aditivă (pornind de la rezultatul net):
  CAF = RN + cheltuieli calculate (681x fără 654, 686) - venituri calculate (781x, 786)
        + VNC active cedate (6583) - venituri din cedări (7583) - 7584 - 7815

Cele două metode trebuie să conducă la aceeași valoare; diferența este verificată și raportată.
Indicatorii sunt rulaje cumulate (P1, P2); P0 nu conține rulaje (n.d.).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Tuple

from docx import Document

from core_balanta import (
    BalantaData,
    add_heading,
    add_paragraph,
    calc_delta,
    calc_growth_rate,
    calc_index,
    calc_ritm_mediu,
    describe_flow,
    fmt_money,
    make_table,
    safe_div,
    fmt_ratio_pct,
)
from sheet_02_cpp import CPPResult
from sheet_13_sig import SIGResult

Triple = Tuple[float, float, float]


@dataclass
class CAFResult:
    caf_deductiv: Triple
    caf_aditiv: Triple
    delta_ded: Tuple[float, float]
    delta_ad: Tuple[float, float]
    index_ded: Tuple[Optional[float], Optional[float]]
    index_ad: Tuple[Optional[float], Optional[float]]
    growth_ded: Tuple[Optional[float], Optional[float]]
    growth_ad: Tuple[Optional[float], Optional[float]]
    ritm_ded: Optional[float]
    ritm_ad: Optional[float]
    componente_ded: Tuple[Tuple[str, Triple], ...] = ()
    componente_ad: Tuple[Tuple[str, Triple], ...] = ()

    @property
    def diferenta(self) -> Triple:
        return tuple(round(self.caf_deductiv[p] - self.caf_aditiv[p], 2) for p in range(3))  # type: ignore[return-value]


def _t(fn) -> Triple:
    return tuple(round(fn(p), 2) for p in range(3))  # type: ignore[return-value]


def compute_sheet_14_caf(cpp: CPPResult, sig: SIGResult) -> CAFResult:
    """Calculează CAF deductiv și CAF aditiv pentru P1 și P2 (P0 = n.d.)."""
    g = lambda k, p: cpp.it(k)[p]  # noqa: E731

    ebe = sig.ebe
    alte_ven_inc = _t(lambda p: g("alte_ven_expl", p) + g("creante_reactivate", p))
    alte_ch_plat = _t(lambda p: g("alte_ch_expl", p) + g("pierderi_creante", p))
    ven_fin_inc = _t(lambda p: g("ven_fin", p))
    ch_fin_plat = _t(lambda p: g("dobanzi", p) + g("alte_ch_fin", p))
    rez_extra = _t(lambda p: g("ven_extra", p) - g("ch_extra", p))
    imp = cpp.impozit_profit
    ded = _t(lambda p: ebe[p] + alte_ven_inc[p] - alte_ch_plat[p] + ven_fin_inc[p] - ch_fin_plat[p] + rez_extra[p] - imp[p])

    rn = cpp.rezultat_net
    ch_calc = _t(lambda p: g("amortizare", p) + g("ajust_ac_ch", p) + g("ajust_prov", p) + g("ajust_fin_ch", p))
    ven_calc = _t(lambda p: g("rel_ajust_imob", p) + g("rel_ajust_ac", p) + g("rel_prov", p) + g("rel_ajust_fin", p))
    cedari = _t(lambda p: g("vnc_cedate", p) - g("ven_cedare", p) - g("ven_subv_inv", p) - g("ven_fc_neg", p))
    ad = _t(lambda p: rn[p] + ch_calc[p] - ven_calc[p] + cedari[p])

    return CAFResult(
        caf_deductiv=ded,
        caf_aditiv=ad,
        delta_ded=calc_delta(ded),
        delta_ad=calc_delta(ad),
        index_ded=calc_index(ded),
        index_ad=calc_index(ad),
        growth_ded=calc_growth_rate(ded),
        growth_ad=calc_growth_rate(ad),
        ritm_ded=calc_ritm_mediu(ded),
        ritm_ad=calc_ritm_mediu(ad),
        componente_ded=(
            ("Excedentul brut de exploatare (EBE)", ebe),
            ("+ Alte venituri de exploatare încasabile (ct. 75x fără 7583, 7584; 754)", alte_ven_inc),
            ("- Alte cheltuieli de exploatare plătibile (ct. 65x fără 6583; 654)", alte_ch_plat),
            ("+ Venituri financiare încasabile (ct. 76x)", ven_fin_inc),
            ("- Cheltuieli financiare plătibile (ct. 66x)", ch_fin_plat),
            ("± Rezultatul extraordinar (ct. 77x - 67x)", rez_extra),
            ("- Impozitul pe profit / venit (ct. 69x)", imp),
        ),
        componente_ad=(
            ("Rezultatul net al perioadei", rn),
            ("+ Cheltuieli calculate: amortizări, ajustări, provizioane (ct. 681x, 686x)", ch_calc),
            ("- Venituri calculate: reluări de ajustări și provizioane (ct. 781x, 786x)", ven_calc),
            ("± Efectul cedărilor de active (ct. 6583 - 7583 - 7584 - 7815)", cedari),
        ),
    )


def render_sheet_14_caf(doc: Document, bal: BalantaData, cpp: CPPResult, sig: SIGResult, res: CAFResult) -> None:
    """Generează secțiunea Capacității de Autofinanțare (CAF)."""
    add_heading(doc, "14. Capacitatea de Autofinanțare (CAF — Metoda Deductivă și Metoda Aditivă)", level=1)

    fl = bal.period.flow_labels
    luna = bal.period.month_name or "Luna curentă"
    headers = [
        "Componentă de calcul a capacității de autofinanțare (lei)",
        f"P1 cumulat {fl[1]}",
        f"P2 cumulat {fl[2]}",
        f"Rulaj {luna} (P2 - P1)",
    ]
    rows = []
    tot_idx = []

    def _add(label: str, vals: Triple, is_tot: bool = False) -> None:
        if not is_tot and abs(vals[1]) < 1e-6 and abs(vals[2]) < 1e-6:
            return
        if is_tot:
            tot_idx.append(len(rows))
        rows.append([label, fmt_money(vals[1]), fmt_money(vals[2]), fmt_money(vals[2] - vals[1])])

    for lbl, v in res.componente_ded:
        _add(lbl, v)
    _add("= CAPACITATEA DE AUTOFINANȚARE — METODA DEDUCTIVĂ", res.caf_deductiv, True)
    for lbl, v in res.componente_ad:
        _add(lbl, v)
    _add("= CAPACITATEA DE AUTOFINANȚARE — METODA ADITIVĂ", res.caf_aditiv, True)
    _add("Diferența dintre metode (trebuie să fie 0)", res.diferenta, True)

    make_table(doc, headers=headers, rows=rows, col_widths=[8.2, 2.6, 2.6, 2.6], total_rows=tot_idx, font_size=7.7)

    ok = all(abs(res.diferenta[p]) <= 0.05 for p in (1, 2))
    flow_caf = describe_flow(res.caf_deductiv, bal, "capacitatea de autofinanțare")
    caf2 = res.caf_deductiv[2]
    rn2 = cpp.rezultat_net[2]
    amort2 = cpp.it("amortizare")[2]
    ca2 = cpp.cifra_afaceri[2]
    rata_caf = safe_div(caf2, ca2, default=None) if abs(ca2) > 1e-6 else None

    verif = (
        "Metoda deductivă și metoda aditivă conduc la aceeași valoare în ambele perioade"
        if ok
        else f"ATENȚIE: metodele diferă (P1: {fmt_money(res.diferenta[1])} lei; P2: {fmt_money(res.diferenta[2])} lei), ceea ce indică o încadrare incompletă a conturilor"
    )
    if caf2 > 0:
        concl = (
            f"Activitatea generează un surplus monetar intern de {fmt_money(caf2)} lei până la finele perioadei de raportare 2 "
            f"({fmt_ratio_pct(rata_caf)} din cifra de afaceri), format din rezultatul net ({fmt_money(rn2)} lei) și cheltuielile "
            f"nemonetare cu amortizarea ({fmt_money(amort2)} lei), disponibil pentru investiții, rambursarea datoriilor și dividende."
        )
    else:
        concl = (
            f"Activitatea nu generează surplus monetar intern până la finele perioadei de raportare 2 (CAF = {fmt_money(caf2)} lei), "
            f"ceea ce impune redresarea marjelor operaționale."
        )
    add_paragraph(doc, f"Sinteză CAF: {verif}; {flow_caf}. {concl}")
