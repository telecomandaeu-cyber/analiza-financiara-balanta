#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Orchestrator Principal — Analiza Poziției Financiare și a Performanțelor Financiare.
Extrage datele contabile din balanța de verificare (Balanta.pdf), rulează toate cele 20 de
metodologii de diagnostic financiar interconectate pe cele trei perioade de raportare (0, 1 și 2)
și generează raportul financiar în format Word (.docx).
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path
from typing import Optional, Union

CURRENT_DIR = Path(__file__).resolve().parent
ANALIZA_DIR = CURRENT_DIR if (CURRENT_DIR / "core_balanta.py").exists() else (CURRENT_DIR / "Analiza")
if str(ANALIZA_DIR) not in sys.path:
    sys.path.insert(0, str(ANALIZA_DIR))

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from core_balanta import (
    BalantaData,
    add_bullet,
    add_callout,
    fmt_pp,
    is_nd,
    add_heading,
    add_paragraph,
    add_part_banner,
    configure_document,
    describe_trajectory,
    fmt_money,
    fmt_num,
    fmt_pct,
    fmt_ratio_pct,
    make_table,
    parse_balanta_pdf,
    pick_variant,
    set_cell_margins,
    set_cell_shading,
    set_run_font,
    set_table_borders,
    set_table_full_width,
)
from sheet_01_bilant import compute_sheet_01_bilant, render_sheet_01_bilant
from sheet_02_cpp import compute_sheet_02_cpp, render_sheet_02_cpp
from sheet_03_bilant_financiar import compute_sheet_03_bilant_financiar, render_sheet_03_bilant_financiar
from sheet_04_ratele_activului import compute_sheet_04_ratele_activului, render_sheet_04_ratele_activului
from sheet_05_structura_financiara import compute_sheet_05_structura_financiara, render_sheet_05_structura_financiara
from sheet_06_ratele_pasivului import compute_sheet_06_ratele_pasivului, render_sheet_06_ratele_pasivului
from sheet_07_echilibrul_financiar import compute_sheet_07_echilibrul_financiar, render_sheet_07_echilibrul_financiar
from sheet_08_lichiditatea import compute_sheet_08_lichiditatea, render_sheet_08_lichiditatea
from sheet_09_solvabilitatea import compute_sheet_09_solvabilitatea, render_sheet_09_solvabilitatea
from sheet_10_bilant_simplificat import compute_sheet_10_bilant_simplificat, render_sheet_10_bilant_simplificat
from sheet_11_echilibru_static import compute_sheet_11_echilibru_static, render_sheet_11_echilibru_static
from sheet_12_echilibru_dinamic import compute_sheet_12_echilibru_dinamic, render_sheet_12_echilibru_dinamic
from sheet_13_sig import compute_sheet_13_sig, render_sheet_13_sig
from sheet_14_caf import compute_sheet_14_caf, render_sheet_14_caf
from sheet_15_capacitatea_indatorare import (
    compute_sheet_15_capacitatea_indatorare,
    render_sheet_15_capacitatea_indatorare,
)
from sheet_16_efectul_indatorare import compute_sheet_16_efectul_indatorare, render_sheet_16_efectul_indatorare
from sheet_17_riscul_faliment import compute_sheet_17_riscul_faliment, render_sheet_17_riscul_faliment
from sheet_18_indicatori_gestiune import compute_sheet_18_indicatori_gestiune, render_sheet_18_indicatori_gestiune
from sheet_19_indicatori_profitabilitate import (
    compute_sheet_19_indicatori_profitabilitate,
    render_sheet_19_indicatori_profitabilitate,
)
from sheet_20_corelatii_pearson_spearman import compute_sheet_20_corelatii, render_sheet_20_corelatii


def build_kpi_rows(bal, bf, sf, rp, ef, liq, solv, sig, caf, lev, risc, prof, bil_ok: bool = True) -> list:
    """Rândurile Tabloului de Bord KPI (folosite atât în raportul Word, cât și în interfața web)."""
    cap_soc2 = bf.capital_social[2]
    if sf.cp[2] <= 0:
        verdict_cp = "Semnal nefavorabil (activ net negativ)"
    elif cap_soc2 > 0 and sf.cp[2] < 0.5 * cap_soc2:
        verdict_cp = "Semnal nefavorabil (activ net < ½ capital subscris, art. 153^24 L. 31/1990)"
    else:
        verdict_cp = "Semnal favorabil (activ net ≥ ½ capital subscris)"
    z2 = risc.z_altman[2]
    if is_nd(z2):
        verdict_z = "n.d."
    elif z2 > 2.90:
        verdict_z = "Societate solvabilă (Z' > 2,90)"
    elif z2 >= 1.23:
        verdict_z = "Zonă de incertitudine (1,23–2,90)"
    else:
        verdict_z = "Faliment iminent / risc ridicat (Z' < 1,23)"

    kpi_rows = [
        [
            "1. Activ Total Bilanțier = Pasiv Total (lei)",
            fmt_money(bf.total_activ[0]),
            fmt_money(bf.total_activ[1]),
            fmt_money(bf.total_activ[2]),
            "Da (bilanț echilibrat)" if bil_ok else "Nu (diferențe Activ - Pasiv)",
        ],
        [
            "2. Capitaluri Proprii — Activ Net Contabil (lei)",
            fmt_money(sf.cp[0]),
            fmt_money(sf.cp[1]),
            fmt_money(sf.cp[2]),
            verdict_cp,
        ],
        [
            "3. Rata Stabilității Financiare — RSF (%)",
            fmt_ratio_pct(rp.rsf[0]),
            fmt_ratio_pct(rp.rsf[1]),
            fmt_ratio_pct(rp.rsf[2]),
            "Semnal favorabil (>50%)" if rp.rsf[2] >= 0.50 else "Sub 50%",
        ],
        [
            "4. Fondul de Rulment Patrimonial — FR (lei)",
            fmt_money(ef.fr_sus[0]),
            fmt_money(ef.fr_sus[1]),
            fmt_money(ef.fr_sus[2]),
            "EFTL (FR > 0)" if ef.fr_sus[2] >= 0 else "DFTL (FR < 0)",
        ],
        [
            "5. Necesarul de Fond de Rulment — NFR (lei)",
            fmt_money(ef.nfr[0]),
            fmt_money(ef.nfr[1]),
            fmt_money(ef.nfr[2]),
            "EFTS (FR > NFR)" if ef.fr_sus[2] >= ef.nfr[2] else "DFTS (FR < NFR)",
        ],
        [
            "6. Trezoreria Netă — TN = FR - NFR (lei)",
            fmt_money(ef.tn[0]),
            fmt_money(ef.tn[1]),
            fmt_money(ef.tn[2]),
            "EFC (TN > 0)" if ef.tn[2] >= 0 else "DFC (TN < 0)",
        ],
        [
            "7. Lichiditatea Curentă / Lichiditatea Imediată (%)",
            f"{fmt_ratio_pct(liq.l_curenta[0])} / {fmt_ratio_pct(liq.l_imediata[0])}",
            f"{fmt_ratio_pct(liq.l_curenta[1])} / {fmt_ratio_pct(liq.l_imediata[1])}",
            f"{fmt_ratio_pct(liq.l_curenta[2])} / {fmt_ratio_pct(liq.l_imediata[2])}",
            "Da (În plafon)" if 1.2 <= liq.l_curenta[2] <= 2.0 else "În afara intervalului [120%-200%]",
        ],
        [
            "8. Solvabilitatea Generală / Solvabilitatea Efectivă (%)",
            f"{fmt_ratio_pct(solv.solv_generala[0])} / {fmt_ratio_pct(solv.solv_efectiva[0])}",
            f"{fmt_ratio_pct(solv.solv_generala[1])} / {fmt_ratio_pct(solv.solv_efectiva[1])}",
            f"{fmt_ratio_pct(solv.solv_generala[2])} / {fmt_ratio_pct(solv.solv_efectiva[2])}",
            "Da (Solvabilitate ≥ 166%)" if solv.solv_generala[2] >= 1.66 else "Sub pragul de 166%",
        ],
        [
            "9. Cifra de Afaceri Netă / Rezultatul Net — cumulat (lei)",
            f"{fmt_money(prof.blocks[0].den_vals[0])} / {fmt_money(sig.rezultat_net[0])}",
            f"{fmt_money(prof.blocks[0].den_vals[1])} / {fmt_money(sig.rezultat_net[1])}",
            f"{fmt_money(prof.blocks[0].den_vals[2])} / {fmt_money(sig.rezultat_net[2])}",
            f"{'Profit' if sig.rezultat_net[2] > 0 else 'Pierdere'} (marja netă {fmt_pct(prof.marja_neta[2])})",
        ],
        [
            "10. Capacitatea de Autofinanțare — CAF cumulată (lei)",
            fmt_money(caf.caf_deductiv[0]),
            fmt_money(caf.caf_deductiv[1]),
            fmt_money(caf.caf_deductiv[2]),
            "Da (CAF > 0)" if caf.caf_deductiv[2] > 0 else "Nu (CAF ≤ 0)",
        ],
        [
            "11. ROA / ROE anualizate (%) / Efect de levier (pp)",
            f"{fmt_pct(prof.roa[0])} / {fmt_pct(prof.roe[0])} / {fmt_pp(lev.efect_levier[0])}",
            f"{fmt_pct(prof.roa[1])} / {fmt_pct(prof.roe[1])} / {fmt_pp(lev.efect_levier[1])}",
            f"{fmt_pct(prof.roa[2])} / {fmt_pct(prof.roe[2])} / {fmt_pp(lev.efect_levier[2])}",
            lev.verdict[2],
        ],
        [
            "12. Scorul Altman Z' / Scorul Conan-Holder Z",
            f"{fmt_num(risc.z_altman[0], 2)} / {fmt_num(risc.z_canon_holder[0], 3)}",
            f"{fmt_num(risc.z_altman[1], 2)} / {fmt_num(risc.z_canon_holder[1], 3)}",
            f"{fmt_num(risc.z_altman[2], 2)} / {fmt_num(risc.z_canon_holder[2], 3)}",
            verdict_z,
        ],
    ]
    return kpi_rows


def _render_cover_and_intro(
    doc: Document,
    bal: BalantaData,
    bf,
    sf,
    rp,
    ef,
    liq,
    solv,
    sig,
    caf,
    lev,
    risc,
    prof,
    bil_ok: bool = True,
) -> None:
    """Generează antetul executiv, Tabloul de Bord KPI și sinteza balanței de verificare."""
    banner = doc.add_table(rows=1, cols=1)
    set_table_full_width(banner)
    set_table_borders(banner, color="0F2C59")
    c = banner.cell(0, 0)
    set_cell_shading(c, "0F2C59")
    set_cell_margins(c, top=90, bottom=90, start=140, end=140)

    p0 = c.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(2)
    r_sub_top = p0.add_run("DIAGNOSTIC ECONOMICO-FINANCIAR PATRIMONIAL ȘI DE PERFORMANȚĂ\n")
    set_run_font(r_sub_top, size=8.0, bold=True, color=(147, 197, 253))
    r_main = p0.add_run("RAPORT DE ANALIZĂ A POZIȚIEI FINANCIARE ȘI A PERFORMANȚELOR FINANCIARE\n")
    set_run_font(r_main, size=12.5, bold=True, color=(255, 255, 255))
    r_comp = p0.add_run(
        f"{bal.denumire}   |   C.I.F.: {bal.cif}   |   Reg. Com.: {bal.reg_com}   |   Sediul: {bal.adresa}   |   Balanța la {bal.perioada}"
    )
    set_run_font(r_comp, size=8.8, bold=True, color=(226, 232, 240))

    tot_r = bal.total_row
    intro_txt = pick_variant(
        bal,
        "intro_opening",
        [
            f"Prezentul raport evaluează poziția financiară, echilibrul patrimonial, lichiditatea, solvabilitatea, rentabilitatea și riscul de insolvență ale societății {bal.denumire} (C.I.F. {bal.cif}) pe baza datelor contabile din balanța de verificare încheiată la {bal.perioada}. Analiza dinamică urmărește trei momente succesive: Perioada de raportare 0 (soldurile inițiale SID = SIC = {fmt_money(tot_r.sid if tot_r else 0)} lei), Perioada de raportare 1 (situația patrimonială și rezultatele cumulate la finele lunii precedente) și Perioada de raportare 2 (soldurile finale SFD = SFC = {fmt_money(tot_r.sfd if tot_r else 0)} lei și rulajele cumulate TSD = TSC = {fmt_money(tot_r.tsd if tot_r else 0)} lei).",
            f"Acest studiu de diagnostic economico-financiar investighează arhitectura patrimonială, dinamica rezultatelor, starea de echilibru și bonitatea bancară a societății {bal.denumire} (C.I.F. {bal.cif}), valorificând informațiile din balanța de verificare la {bal.perioada}. Evaluarea comparativă este structurată pe trei repere temporale: Perioada de raportare 0 (egalitatea soldurilor inițiale SID = SIC = {fmt_money(tot_r.sid if tot_r else 0)} lei), Perioada de raportare 1 (poziția intermediară cumulată) și Perioada de raportare 2 (totalul sumelor TSD = TSC = {fmt_money(tot_r.tsd if tot_r else 0)} lei și soldurile finale SFD = SFC = {fmt_money(tot_r.sfd if tot_r else 0)} lei).",
            f"Documentul sintetizează analiza multilaterală a poziției financiare și a performanțelor operaționale înregistrate de {bal.denumire} (C.I.F. {bal.cif}), conform balanței de verificare la {bal.perioada}. Traiectoria indicatorilor este urmărită pe parcursul a trei secvențe de raportare: Perioada de raportare 0 (solduri inițiale debitoare și creditoare de {fmt_money(tot_r.sid if tot_r else 0)} lei), Perioada de raportare 1 (nivelul cumulat precedent) și Perioada de raportare 2 (rulaje curente de {fmt_money(tot_r.rld if tot_r else 0)} lei, total sume de {fmt_money(tot_r.tsd if tot_r else 0)} lei și solduri finale de {fmt_money(tot_r.sfd if tot_r else 0)} lei).",
        ],
    )
    add_paragraph(
        doc,
        intro_txt,
        space_before=3.0,
        space_after=2.5,
    )

    per = bal.period
    fl = per.flow_labels
    nota = (
        f"Perioada de raportare 0 = soldurile de deschidere la {per.dates[0]} (fără rulaje de venituri și cheltuieli, de aceea "
        f"indicatorii de flux sunt „n.d.”); Perioada de raportare 1 = soldurile la {per.dates[1]} și rulajele cumulate {fl[1]}; "
        f"Perioada de raportare 2 = soldurile la {per.dates[2]} și rulajele cumulate {fl[2]}. Rulajele cumulate nu se compară "
        f"direct între P1 și P2 (acoperă {per.months[1]} și {per.months[2]} luni); se prezintă rulajul lunii curente (P2 - P1), iar "
        f"ratele care raportează un flux la un sold (ROA, ROE, Re, Rd, vitezele de rotație, scorurile de risc) folosesc valori "
        f"anualizate ({per.days_in_year} / {per.days[1]} zile pentru P1 și {per.days_in_year} / {per.days[2]} zile pentru P2). "
        f"Controale efectuate: balanța echilibrată, Activ = Pasiv în fiecare perioadă, rezultatul net din CPP = soldul contului 121, "
        f"cascada SIG = CPP, CAF deductiv = CAF aditiv, FR (sus) = FR (jos) și TN = FR - NFR."
    )
    add_callout(doc, nota, title="Notă metodologică")

    add_heading(doc, "0.1. Tablou de Bord Executiv — Sinteza Indicatorilor Cheie (Top 12 KPI)", level=2)
    kpi_headers = [
        "Indicator Cheie de Performanță și Echilibru (KPI)",
        *bal.period_labels,
        "Verdict și Încadrare",
    ]
    kpi_align = [
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.CENTER,
    ]
    kpi_rows = build_kpi_rows(bal, bf, sf, rp, ef, liq, solv, sig, caf, lev, risc, prof, bil_ok)
    make_table(
        doc,
        headers=kpi_headers,
        rows=kpi_rows,
        col_widths=[6.0, 2.8, 2.8, 2.8, 3.4],
        align_map=kpi_align,
        font_size=7.8,
    )

    add_heading(doc, "0.2. Sinteza Balanței de Verificare pe Clase de Conturi (Clasa 1 — Clasa 7)", level=2)
    cls_headers = [
        "Clasa",
        "Denumire clasă contabilă",
        "SID (lei)",
        "SIC (lei)",
        f"Rulaj D {bal.period.month_name[:3]}.",
        f"Rulaj C {bal.period.month_name[:3]}.",
        "Total Sume D",
        "Total Sume C",
        "SFD (lei)",
        "SFC (lei)",
    ]
    cls_align = [
        WD_ALIGN_PARAGRAPH.CENTER,
        WD_ALIGN_PARAGRAPH.LEFT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
        WD_ALIGN_PARAGRAPH.RIGHT,
    ]
    cls_rows = []
    for c_num in range(1, 8):
        key = f"Clasa {c_num}"
        if key in bal.accounts:
            a = bal.accounts[key]
            cls_rows.append([
                key,
                a.name,
                fmt_money(a.sid),
                fmt_money(a.sic),
                fmt_money(a.rld),
                fmt_money(a.rlc),
                fmt_money(a.tsd),
                fmt_money(a.tsc),
                fmt_money(a.sfd),
                fmt_money(a.sfc),
            ])
    if bal.total_row:
        t = bal.total_row
        cls_rows.append([
            "TOTAL",
            "TOTAL BALANȚĂ DE VERIFICARE",
            fmt_money(t.sid),
            fmt_money(t.sic),
            fmt_money(t.rld),
            fmt_money(t.rlc),
            fmt_money(t.tsd),
            fmt_money(t.tsc),
            fmt_money(t.sfd),
            fmt_money(t.sfc),
        ])
    make_table(
        doc,
        headers=cls_headers,
        rows=cls_rows,
        col_widths=[1.3, 3.7, 1.6, 1.6, 1.6, 1.6, 1.6, 1.6, 1.6, 1.6],
        align_map=cls_align,
        total_rows=[len(cls_rows) - 1] if cls_rows else None,
        font_size=7.6,
    )


def _render_executive_conclusions(
    doc: Document,
    bal: BalantaData,
    bf,
    sf,
    rp,
    ef,
    liq,
    solv,
    sig,
    caf,
    cap_ind,
    lev,
    risc,
    gest,
    prof,
    corr,
) -> None:
    """Generează capitolul final de sinteză executivă și recomandări strategice."""
    add_heading(doc, "21. Sinteza Diagnosticului Global și Recomandări Manageriale", level=1)

    g_cp = (sf.cp[2] - sf.cp[0]) / abs(sf.cp[0]) if abs(sf.cp[0]) > 1e-6 else 0.0
    acop_imob = sf.cp[2] / bf.active_imobilizate[2] if abs(bf.active_imobilizate[2]) > 1e-6 else 0.0
    traj_cp_exec = describe_trajectory(sf.cp, bal, "exec_cp", "capitalurile proprii")

    b1_txt = pick_variant(
        bal,
        "exec_b1",
        [
            f"{traj_cp_exec[0].upper() + traj_cp_exec[1:]}, acoperind activele imobilizate nete ({fmt_money(bf.active_imobilizate[2])} lei în perioada de raportare 2) în proporție de {fmt_ratio_pct(acop_imob)} și generând în perioada de raportare 2 un Fond de Rulment Patrimonial de {fmt_money(ef.fr_sus[2])} lei, un Fond de Rulment Propriu de {fmt_money(ef.frp[2])} lei și o Trezorerie Netă de {fmt_money(ef.tn[2])} lei.",
            f"Arhitectura financiară a societății este marcată de faptul că {traj_cp_exec}; la finele perioadei de raportare 2, resursele proprii finanțează în proporție de {fmt_ratio_pct(acop_imob)} imobilizările nete ({fmt_money(bf.active_imobilizate[2])} lei), susținând un Fond de Rulment de {fmt_money(ef.fr_sus[2])} lei (FRP = {fmt_money(ef.frp[2])} lei) și o Trezorerie Netă de {fmt_money(ef.tn[2])} lei.",
        ],
    )

    b2_txt = pick_variant(
        bal,
        "exec_b2",
        [
            f"În perioada de raportare 2, treptele de lichiditate — curentă ({fmt_ratio_pct(liq.l_curenta[2])}), rapidă ({fmt_ratio_pct(liq.l_rapida[2])}) și imediată ({fmt_ratio_pct(liq.l_imediata[2])}) — atestă gradul de acoperire a datoriilor pe termen scurt ({fmt_money(bf.datorii_termen_scurt[2])} lei) din activele circulante și din disponibilitățile bănești ({fmt_money(bf.disponibilitati[2])} lei). Solvabilitatea generală se situează la {fmt_ratio_pct(solv.solv_generala[2])}, iar raportul dintre rentabilitatea economică anualizată (Re = {fmt_ratio_pct(lev.re_rate[2])}) și costul mediu al îndatorării (Rd = {fmt_ratio_pct(lev.rd_rate[2])}) conduce la un efect de levier de {fmt_pp(lev.efect_levier[2])} ({lev.verdict[2]}).",
            f"Capacitatea de plată în perioada de raportare 2 este caracterizată de o lichiditate curentă de {fmt_ratio_pct(liq.l_curenta[2])}, o lichiditate rapidă de {fmt_ratio_pct(liq.l_rapida[2])} și o lichiditate imediată de {fmt_ratio_pct(liq.l_imediata[2])} (lichiditate imediată pe datoriile exigibile de {fmt_ratio_pct(liq.l_imediata_exig[2])}), dublată de o solvabilitate generală de {fmt_ratio_pct(solv.solv_generala[2])} și un efect de levier de {fmt_pp(lev.efect_levier[2])} ({lev.verdict[2]}).",
        ],
    )

    b3_txt = pick_variant(
        bal,
        "exec_b3",
        [
            f"Până la finele perioadei de raportare 2 ({bal.period.months[2]} luni cumulate), activitatea operațională degajă o cifră de afaceri netă de {fmt_money(prof.blocks[0].den_vals[2])} lei, o valoare adăugată de {fmt_money(sig.valoarea_adaugata[2])} lei, un rezultat al exploatării de {fmt_money(sig.rezultat_exploatare[2])} lei și un rezultat net de {fmt_money(sig.rezultat_net[2])} lei (ROA anualizat = {fmt_pct(prof.roa[2])}, ROE anualizat = {fmt_pct(prof.roe[2])}, marja operațională = {fmt_pct(prof.ros[2])}), generând o capacitate de autofinanțare de {fmt_money(caf.caf_deductiv[2])} lei; scorul Altman Z' este {fmt_num(risc.z_altman[2], 4)} ({risc.interp_altman[2]}).",
            f"Performanța cumulată în perioada de raportare 2 ({bal.period.months[2]} luni) se concretizează într-o cifră de afaceri netă de {fmt_money(prof.blocks[0].den_vals[2])} lei, un excedent brut de exploatare (EBE) de {fmt_money(sig.ebe[2])} lei, un rezultat net de {fmt_money(sig.rezultat_net[2])} lei și o capacitate de autofinanțare de {fmt_money(caf.caf_deductiv[2])} lei, în condițiile unor rate anualizate ROA = {fmt_pct(prof.roa[2])} și ROE = {fmt_pct(prof.roe[2])}, ale unei marje operaționale de {fmt_pct(prof.ros[2])} și ale unui scor Altman Z' de {fmt_num(risc.z_altman[2], 4)}.",
        ],
    )

    df2 = bf.datorii_financiare[2]
    if ef.tn[2] > 0 and df2 > 0:
        b4_txt = f"Având în vedere trezoreria netă pozitivă ({fmt_money(ef.tn[2])} lei) din perioada de raportare 2, se recomandă utilizarea excedentului monetar pentru reducerea datoriilor financiare purtătoare de dobândă ({fmt_money(df2)} lei, cu un cost mediu anualizat Rd = {fmt_ratio_pct(lev.rd_rate[2])}) sau plasarea lui în instrumente cu randament superior acestui cost, concomitent cu monitorizarea decalajului dintre durata de încasare a clienților ({fmt_num(gest.dur_clienti_411[2], 1)} zile) și durata de plată a furnizorilor ({fmt_num(gest.dur_furnizori_401[2], 1)} zile)."
    elif ef.tn[2] > 0:
        b4_txt = f"În contextul unei trezorerii nete excedentare ({fmt_money(ef.tn[2])} lei în perioada de raportare 2) și al absenței datoriilor financiare purtătoare de dobândă, se recomandă fructificarea disponibilităților ({fmt_money(bf.disponibilitati[2])} lei) și optimizarea rotației stocurilor ({fmt_num(gest.dur_stoc[2], 1)} zile) și a încasării clienților ({fmt_num(gest.dur_clienti_411[2], 1)} zile)."
    else:
        b4_txt = f"Pentru consolidarea Trezoreriei Nete ({fmt_money(ef.tn[2])} lei în perioada de raportare 2), se recomandă reducerea Necesarului de Fond de Rulment ({fmt_money(ef.nfr[2])} lei) prin accelerarea încasării creanțelor ({fmt_num(gest.dur_creante[2], 1)} zile), dimensionarea strictă a stocurilor ({fmt_num(gest.dur_stoc[2], 1)} zile) și renegocierea termenelor de plată cu furnizorii ({fmt_num(gest.dur_furnizori_401[2], 1)} zile)."

    add_bullet(doc, "1. Poziție patrimonială și echilibru financiar", b1_txt)
    add_bullet(doc, "2. Lichiditate, solvabilitate și capacitate de îndatorare bancară", b2_txt)
    add_bullet(doc, "3. Performanță operațională, autofinanțare și risc de insolvență", b3_txt)
    add_bullet(doc, "4. Recomandări manageriale adaptate profilului financiar", b4_txt)


def compute_all(bal: BalantaData) -> dict:
    """Rulează cele 20 de module de calcul în ordinea dependențelor și întoarce rezultatele."""
    per = bal.period
    af = (per.annual_factor(0), per.annual_factor(1), per.annual_factor(2))
    s01 = compute_sheet_01_bilant(bal)
    s02 = compute_sheet_02_cpp(bal, s01)
    s03 = compute_sheet_03_bilant_financiar(s01)
    s04 = compute_sheet_04_ratele_activului(s03, s02, per.months)
    s05 = compute_sheet_05_structura_financiara(s03)
    s06 = compute_sheet_06_ratele_pasivului(s05)
    s07 = compute_sheet_07_echilibrul_financiar(s03, s05)
    s08 = compute_sheet_08_lichiditatea(s03, s01)
    s09 = compute_sheet_09_solvabilitatea(s03, s05, s01)
    s10 = compute_sheet_10_bilant_simplificat(s03, s07)
    s11 = compute_sheet_11_echilibru_static(s07)
    s12 = compute_sheet_12_echilibru_dinamic(s11)
    s13 = compute_sheet_13_sig(s02)
    s14 = compute_sheet_14_caf(s02, s13)
    s15 = compute_sheet_15_capacitatea_indatorare(s01, s03, s05)
    s16 = compute_sheet_16_efectul_indatorare(s03, s05, s02, af)
    s17 = compute_sheet_17_riscul_faliment(s01, s02, s03, s05, s07, s13, af)
    s18 = compute_sheet_18_indicatori_gestiune(bal, s03, s02)
    s19 = compute_sheet_19_indicatori_profitabilitate(s03, s02, af)
    s20 = compute_sheet_20_corelatii(s01, s02, s03, s05, s07, s08, s09, s13, s17)
    return {
        "s01": s01, "s02": s02, "s03": s03, "s04": s04, "s05": s05, "s06": s06, "s07": s07,
        "s08": s08, "s09": s09, "s10": s10, "s11": s11, "s12": s12, "s13": s13, "s14": s14,
        "s15": s15, "s16": s16, "s17": s17, "s18": s18, "s19": s19, "s20": s20,
    }


def generate_full_financial_report(
    pdf_path: Union[str, Path],
    output_docx_path: Optional[Union[str, Path]] = None,
) -> Path:
    """
    Rulează întregul flux de analiză financiară pe cele 20 de capitole interconectate
    și salvează raportul Word (.docx).
    """
    bal = parse_balanta_pdf(pdf_path)

    # ETAPA 1: EXECUȚIA CELOR 20 DE MODULE INTERCONECTATE
    r = compute_all(bal)
    s01_bilant, s02_cpp, s03_bf, s04_ra = r["s01"], r["s02"], r["s03"], r["s04"]
    s05_sf, s06_rp, s07_ef, s08_liq = r["s05"], r["s06"], r["s07"], r["s08"]
    s09_solv, s10_bs, s11_es, s12_ed = r["s09"], r["s10"], r["s11"], r["s12"]
    s13_sig, s14_caf, s15_cap, s16_lev = r["s13"], r["s14"], r["s15"], r["s16"]
    s17_risc, s18_gest, s19_prof, s20_corr = r["s17"], r["s18"], r["s19"], r["s20"]

    # ETAPA 2: CONSTRUIREA DOCUMENTULUI WORD (.DOCX) COMPACT ȘI SINTETIC
    doc = Document()
    configure_document(doc, bal)

    _render_cover_and_intro(
        doc, bal, s03_bf, s05_sf, s06_rp, s07_ef, s08_liq, s09_solv, s13_sig, s14_caf, s16_lev, s17_risc, s19_prof,
        bil_ok=s01_bilant.echilibrat,
    )

    add_part_banner(
        doc,
        "PARTEA I  •  CAPITOLELE 1 – 3",
        "SITUAȚIILE FINANCIARE DE SINTEZĂ ȘI RETRATAREA PATRIMONIALĂ",
        "Bilanțul Contabil, Contul de Profit și Pierdere și Bilanțul Financiar retratat pe criterii de lichiditate–exigibilitate.",
        page_break=False,
    )
    render_sheet_01_bilant(doc, bal, s01_bilant)
    render_sheet_02_cpp(doc, bal, s02_cpp)
    render_sheet_03_bilant_financiar(doc, bal, s01_bilant, s03_bf)

    add_part_banner(
        doc,
        "PARTEA A II-A  •  CAPITOLELE 4 – 6",
        "ANALIZA STRUCTURII PATRIMONIALE ȘI FINANCIARE (ACTIV & PASIV)",
        "Ratele de structură ale activului, structura financiară a pasivului și ratele de autonomie și îndatorare bancară.",
        page_break=False,
    )
    render_sheet_04_ratele_activului(doc, bal, s03_bf, s02_cpp, s04_ra)
    render_sheet_05_structura_financiara(doc, bal, s05_sf)
    render_sheet_06_ratele_pasivului(doc, bal, s06_rp)

    add_part_banner(
        doc,
        "PARTEA A III-A  •  CAPITOLELE 7 – 12",
        "ECHILIBRUL FINANCIAR STATIC ȘI DINAMIC, LICHIDITATEA ȘI SOLVABILITATEA",
        "Fondul de Rulment (FR, FRP, FRS), Necesarul de Fond de Rulment (NFR), Trezoreria Netă (TN), Lichiditatea și Solvabilitatea.",
        page_break=False,
    )
    render_sheet_07_echilibrul_financiar(doc, bal, s03_bf, s05_sf, s07_ef)
    render_sheet_08_lichiditatea(doc, bal, s03_bf, s01_bilant, s08_liq)
    render_sheet_09_solvabilitatea(doc, bal, s03_bf, s05_sf, s01_bilant, s09_solv)
    render_sheet_10_bilant_simplificat(doc, bal, s10_bs)
    render_sheet_11_echilibru_static(doc, bal, s11_es)
    render_sheet_12_echilibru_dinamic(doc, bal, s12_ed)

    add_part_banner(
        doc,
        "PARTEA A IV-A  •  CAPITOLELE 13 – 14",
        "ANALIZA PERFORMANȚELOR FINANCIARE: CASCADA SIG ȘI CAPACITATEA DE AUTOFINANȚARE",
        "Soldurile Intermediare de Gestiune (MC, PE, VA, EBE, RE, RC, RN) și Capacitatea de Autofinanțare (Metoda Deductivă = Metoda Aditivă).",
        page_break=False,
    )
    render_sheet_13_sig(doc, bal, s13_sig)
    render_sheet_14_caf(doc, bal, s02_cpp, s13_sig, s14_caf)

    add_part_banner(
        doc,
        "PARTEA A V-A  •  CAPITOLELE 15 – 17",
        "CAPACITATEA DE ÎNDATORARE, EFECTUL DE LEVIER ȘI RISCUL DE INSOLVENȚĂ",
        "Marja de creditare bancară (CIG, CIT), Efectul de Levier Financiar (Re vs. Rd) și Modelele de Scoring Altman Z & Conan-Holder.",
        page_break=False,
    )
    render_sheet_15_capacitatea_indatorare(doc, bal, s15_cap)
    render_sheet_16_efectul_indatorare(doc, bal, s16_lev)
    render_sheet_17_riscul_faliment(doc, bal, s17_risc)

    add_part_banner(
        doc,
        "PARTEA A VI-A  •  CAPITOLELE 18 – 21",
        "GESTIUNEA CICLULUI OPERAȚIONAL, PROFITABILITATEA, CORELAȚIILE STATISTICE ȘI CONCLUZII",
        "Vitezele de rotație în zile, ratele de rentabilitate (ROA, ROE, ROS), Matricea Pearson & Spearman și Recomandări Executive.",
        page_break=False,
    )
    render_sheet_18_indicatori_gestiune(doc, bal, s03_bf, s02_cpp, s18_gest)
    render_sheet_19_indicatori_profitabilitate(doc, bal, s19_prof)
    render_sheet_20_corelatii(doc, bal, s20_corr)
    _render_executive_conclusions(
        doc, bal, s03_bf, s05_sf, s06_rp, s07_ef, s08_liq, s09_solv,
        s13_sig, s14_caf, s15_cap, s16_lev, s17_risc, s18_gest, s19_prof, s20_corr
    )

    if output_docx_path is None:
        utilitare_dir = ANALIZA_DIR.parent
        out_dir = utilitare_dir / "output_analiza_financiara"
        out_dir.mkdir(parents=True, exist_ok=True)
        clean_name = "".join(ch if ch.isalnum() else "_" for ch in bal.denumire).strip("_")
        output_docx_path = out_dir / f"Raport_Analiza_Financiara_{clean_name}_Balanta_{bal.period.year}.docx"

    out_path = Path(output_docx_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    copy_in_analiza = ANALIZA_DIR / out_path.name

    saved_primary = False
    try:
        doc.save(str(out_path))
        saved_primary = True
    except PermissionError:
        alt_out = out_path.with_name(f"{out_path.stem}_Actualizat{out_path.suffix}")
        doc.save(str(alt_out))
        out_path = alt_out

    if copy_in_analiza.resolve() != out_path.resolve():
        try:
            if saved_primary:
                shutil.copy2(out_path, copy_in_analiza)
            else:
                doc.save(str(copy_in_analiza))
        except PermissionError:
            alt_copy = copy_in_analiza.with_name(f"{copy_in_analiza.stem}_Actualizat{copy_in_analiza.suffix}")
            doc.save(str(alt_copy))

    return out_path


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    parser = argparse.ArgumentParser(
        description="Generează Raportul Word de Analiză Financiară din Balanta.pdf"
    )
    parser.add_argument(
        "pdf_path",
        nargs="?",
        default=None,
        help="Calea către fișierul Balanta.pdf (implicit: caută Balanta.pdf în folderul Utilitare)",
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Calea fișierului Word (.docx) generat",
    )
    args = parser.parse_args()

    if args.pdf_path:
        pdf_file = Path(args.pdf_path)
    else:
        candidates = [
            CURRENT_DIR / "Balanta.pdf",
            ANALIZA_DIR.parent / "Balanta.pdf",
            Path(r"c:\Users\Florentin\Desktop\Import\Utilitare\Balanta.pdf"),
        ]
        pdf_file = next((c for c in candidates if c.exists()), candidates[-1])

    out_path = generate_full_financial_report(pdf_file, args.output)
    print("[OK] Raportul Word de analiză financiară a fost generat cu succes:")
    print(f"     -> {out_path}")
    print(f"     -> {ANALIZA_DIR / out_path.name}")


if __name__ == "__main__":
    main()
