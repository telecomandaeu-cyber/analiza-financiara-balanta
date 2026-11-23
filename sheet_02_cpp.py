#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Capitolul 2: Contul de Profit și Pierdere Sintetic (structura OMFP nr. 1802/2014).

Veniturile și cheltuielile sunt rulaje CUMULATE de la începutul anului:
  - Perioada de raportare 1: de la 01.01 până la finele lunii precedente;
  - Perioada de raportare 2: de la 01.01 până la finele lunii de raportare.
Perioada de raportare 0 (soldurile de deschidere de la 01.01) nu conține rulaje de venituri
și cheltuieli, de aceea indicatorii de flux sunt marcați „n.d.” pentru P0 (nu se estimează).

Fiecare cont de clasa 6 și 7 din balanță este repartizat pe un rând al CPP; conturile
care nu au o încadrare specifică sunt incluse la „alte venituri/cheltuieli de exploatare”
și semnalate, astfel încât rezultatul net să fie complet.
Variația stocurilor de produse (ct. 711) se determină din modificarea soldurilor conturilor
de produse și producție în curs (331, 332, 341, 345, 346, 347, 348, 354) față de 01.01,
deoarece la închiderea lunară a contului 711 rulajele lui nu mai reflectă sensul variației.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Tuple

from docx import Document

from core_balanta import (
    ND,
    BalantaData,
    add_callout,
    add_heading,
    add_paragraph,
    all_ok,
    describe_flow,
    fmt_money,
    fmt_ratio_pct,
    make_table,
    pick_variant,
    safe_div,
)

TOL = 0.05

PRODUSE_STOC = ("331", "332", "341", "345", "346", "347", "348", "354")
SPLIT_PARENTS = {"681", "741", "758", "658", "781", "786", "686", "666", "766"}

ITEM_KEYS = (
    "prod_vanduta", "venit_marfuri", "reduceri_acordate", "subv_ca", "var_stoc", "prod_imob", "subv_expl",
    "alte_ven_expl", "ven_cedare", "ven_subv_inv", "ven_fc_neg", "creante_reactivate", "rel_prov",
    "rel_ajust_imob", "rel_ajust_ac", "ven_fin", "rel_ajust_fin", "ven_extra",
    "mat_prime", "alte_mat", "energie", "marfuri", "reduceri_primite", "prest_ext", "impozite_taxe",
    "salarii", "asigurari", "pierderi_creante", "alte_ch_expl", "vnc_cedate", "amortizare", "ajust_prov",
    "ajust_ac_ch", "ajust_fin_ch", "dobanzi", "alte_ch_fin", "ch_extra", "impozit_profit",
)


def _cpp_key(code: str) -> Tuple[str, str]:
    """(cheie, natura) — natura „C” = rulaj creditor (venit), „D” = rulaj debitor (cheltuială)."""
    c = code
    c2, c3 = c[:2], c[:3]
    if c[0] == "7":
        if c3 in ("701", "702", "703", "704", "705", "706", "708"):
            return "prod_vanduta", "C"
        if c3 == "707":
            return "venit_marfuri", "C"
        if c3 == "709":
            return "reduceri_acordate", "D"
        if c3 in ("711", "712"):
            return "var_stoc", "X"
        if c2 == "72":
            return "prod_imob", "C"
        if c3 == "741":
            return ("subv_ca", "C") if c.startswith("7411") else ("subv_expl", "C")
        if c3 == "754":
            return "creante_reactivate", "C"
        if c.startswith("7583"):
            return "ven_cedare", "C"
        if c.startswith("7584"):
            return "ven_subv_inv", "C"
        if c2 == "75":
            return "alte_ven_expl", "C"
        if c3 == "781":
            return {
                "7812": "rel_prov", "7813": "rel_ajust_imob", "7814": "rel_ajust_ac", "7815": "ven_fc_neg",
            }.get(c, "rel_ajust_imob"), "C"
        if c3 == "786":
            return "rel_ajust_fin", "C"
        if c2 == "76":
            return "ven_fin", "C"
        if c2 == "77":
            return "ven_extra", "C"
        return "alte_ven_expl", "C?"
    if c[0] == "6":
        if c3 in ("601", "602"):
            return "mat_prime", "D"
        if c3 in ("603", "604", "606", "608"):
            return "alte_mat", "D"
        if c3 == "605":
            return "energie", "D"
        if c3 == "607":
            return "marfuri", "D"
        if c3 == "609":
            return "reduceri_primite", "C"
        if c2 in ("61", "62"):
            return "prest_ext", "D"
        if c2 == "63":
            return "impozite_taxe", "D"
        if c3 in ("641", "642", "643", "644"):
            return "salarii", "D"
        if c2 == "64":
            return "asigurari", "D"
        if c3 == "654":
            return "pierderi_creante", "D"
        if c.startswith("6583"):
            return "vnc_cedate", "D"
        if c2 == "65":
            return "alte_ch_expl", "D"
        if c3 == "681":
            return {"6812": "ajust_prov", "6814": "ajust_ac_ch"}.get(c, "amortizare"), "D"
        if c3 == "686":
            return "ajust_fin_ch", "D"
        if c3 == "666":
            return "dobanzi", "D"
        if c2 == "66":
            return "alte_ch_fin", "D"
        if c2 == "67":
            return "ch_extra", "D"
        if c2 == "69":
            return "impozit_profit", "D"
        return "alte_ch_expl", "D?"
    return "", ""


def _flow_accounts(bal: BalantaData) -> List[str]:
    """Conturile de clasa 6/7 la nivelul la care se face încadrarea (cu divizarea pe analitice unde contează)."""
    out: List[str] = []
    for syn in bal.synthetic_accounts():
        if syn[0] not in "67":
            continue
        kids = bal.children(syn) if syn in SPLIT_PARENTS else []
        if kids:
            ok = all(
                abs(sum(bal.flow_d(k, p) for k in kids) - bal.flow_d(syn, p)) < TOL
                and abs(sum(bal.flow_c(k, p) for k in kids) - bal.flow_c(syn, p)) < TOL
                for p in (1, 2)
            )
            if ok:
                out.extend(kids)
                continue
        out.append(syn)
    return out


def compute_cpp_items(bal: BalantaData) -> Tuple[Dict[str, Tuple[float, float, float]], List[str]]:
    """Rulajele cumulate pe elementele de venituri și cheltuieli (P0 = n.d.) și observațiile de încadrare."""
    items: Dict[str, List[float]] = {k: [ND, 0.0, 0.0] for k in ITEM_KEYS}
    notes: List[str] = []
    for code in _flow_accounts(bal):
        key, nat = _cpp_key(code)
        if not key or nat == "X":
            continue
        if nat.endswith("?"):
            notes.append(f"contul {code} a fost inclus la {'alte venituri' if code[0] == '7' else 'alte cheltuieli'} de exploatare")
            nat = nat[0]
        for p in (1, 2):
            items[key][p] += bal.flow_c(code, p) if nat == "C" else bal.flow_d(code, p)

    # Variația stocurilor de produse = sold(perioadă) - sold(01.01) pentru conturile de produse.
    for p in (1, 2):
        items["var_stoc"][p] = sum(bal.net_d(c, p) - bal.net_d(c, 0) for c in PRODUSE_STOC)

    return {k: (v[0], round(v[1], 2), round(v[2], 2)) for k, v in items.items()}, notes


def _assemble(it: Dict[str, Tuple[float, float, float]]) -> Dict[int, List[float]]:
    r: Dict[int, List[float]] = {k: [ND, 0.0, 0.0] for k in range(1, 69)}
    for p in (1, 2):
        g = lambda k: it[k][p]  # noqa: E731
        r[2][p] = g("prod_vanduta")
        r[3][p] = g("venit_marfuri")
        r[4][p] = g("reduceri_acordate")
        r[5][p] = 0.0
        r[6][p] = g("subv_ca")
        r[1][p] = r[2][p] + r[3][p] - r[4][p] + r[5][p] + r[6][p]
        vs = g("var_stoc")
        r[7][p], r[8][p] = (vs, 0.0) if vs >= 0 else (0.0, -vs)
        r[9][p] = g("prod_imob")
        r[10][p] = g("alte_ven_expl") + g("ven_cedare") + g("ven_subv_inv") + g("ven_fc_neg")
        r[11][p] = g("subv_expl")
        r[12][p] = r[1][p] + r[7][p] - r[8][p] + r[9][p] + r[10][p] + r[11][p]

        r[13][p] = g("mat_prime")
        r[14][p] = g("alte_mat")
        r[15][p] = g("energie")
        r[16][p] = g("marfuri")
        r[17][p] = g("reduceri_primite")
        r[19][p] = g("salarii")
        r[20][p] = g("asigurari")
        r[18][p] = r[19][p] + r[20][p]
        r[22][p] = g("amortizare")
        r[23][p] = g("rel_ajust_imob")
        r[21][p] = r[22][p] - r[23][p]
        r[25][p] = g("ajust_ac_ch") + g("pierderi_creante")
        r[26][p] = g("rel_ajust_ac") + g("creante_reactivate")
        r[24][p] = r[25][p] - r[26][p]
        r[28][p] = g("prest_ext")
        r[29][p] = g("impozite_taxe")
        r[30][p] = g("alte_ch_expl") + g("vnc_cedate")
        r[31][p] = 0.0
        r[27][p] = r[28][p] + r[29][p] + r[30][p] + r[31][p]
        r[33][p] = g("ajust_prov")
        r[34][p] = g("rel_prov")
        r[32][p] = r[33][p] - r[34][p]
        r[35][p] = (
            r[13][p] + r[14][p] + r[15][p] + r[16][p] - r[17][p] + r[18][p] + r[21][p] + r[24][p] + r[27][p] + r[32][p]
        )
        rex = r[12][p] - r[35][p]
        r[36][p], r[37][p] = (rex, 0.0) if rex >= 0 else (0.0, -rex)

        r[45][p] = g("ven_fin")
        r[47][p] = g("ajust_fin_ch")
        r[48][p] = g("rel_ajust_fin")
        r[46][p] = r[47][p] - r[48][p]
        r[49][p] = g("dobanzi")
        r[51][p] = g("alte_ch_fin")
        r[52][p] = r[46][p] + r[49][p] + r[51][p]
        rfin = r[45][p] - r[52][p]
        r[53][p], r[54][p] = (rfin, 0.0) if rfin >= 0 else (0.0, -rfin)
        rcur = rex + rfin
        r[55][p], r[56][p] = (rcur, 0.0) if rcur >= 0 else (0.0, -rcur)

        r[57][p] = g("ven_extra")
        r[58][p] = g("ch_extra")
        rext = r[57][p] - r[58][p]
        r[59][p], r[60][p] = (rext, 0.0) if rext >= 0 else (0.0, -rext)

        r[61][p] = r[12][p] + r[45][p] + r[57][p]
        r[62][p] = r[35][p] + r[52][p] + r[58][p]
        rbrut = r[61][p] - r[62][p]
        r[63][p], r[64][p] = (rbrut, 0.0) if rbrut >= 0 else (0.0, -rbrut)
        r[65][p] = g("impozit_profit")
        r[66][p] = 0.0
        rnet = rbrut - r[65][p] - r[66][p]
        r[67][p], r[68][p] = (rnet, 0.0) if rnet >= 0 else (0.0, -rnet)
    for k in r:
        r[k] = [r[k][0], round(r[k][1], 2), round(r[k][2], 2)]
    return r


def compute_net_result(bal: BalantaData) -> Tuple[float, float, float]:
    """Rezultatul net al exercițiului curent (P0 = n.d.), folosit și de Bilanț pentru reconciliere."""
    items, _ = compute_cpp_items(bal)
    r = _assemble(items)
    return (ND, round(r[67][1] - r[68][1], 2), round(r[67][2] - r[68][2], 2))


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
    items: Dict[str, Tuple[float, float, float]] = field(default_factory=dict)
    note: List[str] = field(default_factory=list)
    rezultat_121: Tuple[float, float, float] = (ND, ND, ND)

    def rd(self, code: int) -> Tuple[float, float, float]:
        return self.rows_by_code.get(code, (ND, 0.0, 0.0))

    def it(self, key: str) -> Tuple[float, float, float]:
        return self.items.get(key, (ND, 0.0, 0.0))

    def _signed(self, pos: int, neg: int) -> Tuple[float, float, float]:
        a, b = self.rd(pos), self.rd(neg)
        return tuple(round(a[p] - b[p], 2) for p in range(3))  # type: ignore[return-value]

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
        return self._signed(36, 37)

    @property
    def venituri_financiare(self) -> Tuple[float, float, float]:
        return self.rd(45)

    @property
    def cheltuieli_financiare(self) -> Tuple[float, float, float]:
        return self.rd(52)

    @property
    def rezultat_financiar(self) -> Tuple[float, float, float]:
        return self._signed(53, 54)

    @property
    def venituri_totale(self) -> Tuple[float, float, float]:
        return self.rd(61)

    @property
    def cheltuieli_totale(self) -> Tuple[float, float, float]:
        return self.rd(62)

    @property
    def rezultat_brut(self) -> Tuple[float, float, float]:
        return self._signed(63, 64)

    @property
    def impozit_profit(self) -> Tuple[float, float, float]:
        return self.rd(65)

    @property
    def rezultat_net(self) -> Tuple[float, float, float]:
        return self._signed(67, 68)

    @property
    def dobanzi(self) -> Tuple[float, float, float]:
        return self.rd(49)


def compute_sheet_02_cpp(bal: BalantaData, bilant=None) -> CPPResult:
    """Calculează Contul de Profit și Pierdere pentru P1 și P2 (rulaje cumulate); P0 = n.d."""
    res = CPPResult()
    items, notes = compute_cpp_items(bal)
    res.items = items
    r = _assemble(items)
    for code, vals in r.items():
        res.rows_by_code[code] = (vals[0], vals[1], vals[2])

    # Reconcilierea cu soldul contului 121 (+ eventualele conturi 6/7 neînchise)
    rez_121 = []
    for p in (0, 1, 2):
        unclosed = sum(-bal.net_d(c, p) for c in bal.synthetic_accounts() if c[0] in "67")
        rez_121.append(round(bal.net_c("121", p) + unclosed, 2))
    res.rezultat_121 = (rez_121[0], rez_121[1], rez_121[2])
    res.note = notes

    rows_def = [
        ("1.", "Cifra de afaceri netă (rd. 02 + 03 - 04 + 06)", 1, True),
        ("", "  - Producția vândută (ct. 701..706 + 708)", 2, False),
        ("", "  - Venituri din vânzarea mărfurilor (ct. 707)", 3, False),
        ("", "  - Reduceri comerciale acordate (ct. 709)", 4, False),
        ("", "  - Subvenții de exploatare aferente cifrei de afaceri (ct. 7411)", 6, False),
        ("2.", "Variația stocurilor de produse — sold creditor (Δ ct. 331..354)", 7, False),
        ("", "Variația stocurilor de produse — sold debitor (Δ ct. 331..354)", 8, False),
        ("3.", "Producția realizată pentru scopuri proprii și capitalizată (ct. 72x)", 9, False),
        ("4.", "Alte venituri din exploatare (ct. 75x fără 754, 7815)", 10, False),
        ("", "Venituri din subvenții de exploatare (ct. 741x fără 7411)", 11, False),
        ("", "VENITURI DIN EXPLOATARE — TOTAL", 12, True),
        ("5.", "a) Cheltuieli cu materiile prime și materialele consumabile (ct. 601 + 602)", 13, False),
        ("", "   Alte cheltuieli materiale (ct. 603 + 604 + 606 + 608)", 14, False),
        ("", "b) Alte cheltuieli externe — energie și apă (ct. 605)", 15, False),
        ("", "c) Cheltuieli privind mărfurile (ct. 607)", 16, False),
        ("", "   Reduceri comerciale primite (ct. 609)", 17, False),
        ("6.", "Cheltuieli cu personalul (ct. 641..646)", 18, True),
        ("7.", "Ajustări de valoare privind imobilizările (ct. 6811 + 6813 - 7813)", 21, False),
        ("", "Ajustări de valoare privind activele circulante (ct. 654 + 6814 - 754 - 7814)", 24, False),
        ("8.", "Alte cheltuieli de exploatare (ct. 61x + 62x + 63x + 65x)", 27, True),
        ("", "  - Prestații externe (ct. 61x + 62x)", 28, False),
        ("", "  - Impozite, taxe și vărsăminte asimilate (ct. 63x)", 29, False),
        ("", "  - Alte cheltuieli (ct. 65x fără 654)", 30, False),
        ("", "Ajustări privind provizioanele (ct. 6812 - 7812)", 32, False),
        ("", "CHELTUIELI DE EXPLOATARE — TOTAL", 35, True),
        ("", "PROFITUL DIN EXPLOATARE (rd. 12 - 35)", 36, True),
        ("", "PIERDEREA DIN EXPLOATARE (rd. 35 - 12)", 37, True),
        ("", "VENITURI FINANCIARE — TOTAL (ct. 76x)", 45, False),
        ("", "Ajustări de valoare privind imobilizările financiare (ct. 686 - 786)", 46, False),
        ("", "Cheltuieli privind dobânzile (ct. 666)", 49, False),
        ("", "Alte cheltuieli financiare (ct. 663..668 fără 666)", 51, False),
        ("", "CHELTUIELI FINANCIARE — TOTAL", 52, False),
        ("", "PROFITUL FINANCIAR (rd. 45 - 52)", 53, False),
        ("", "PIERDEREA FINANCIARĂ (rd. 52 - 45)", 54, False),
        ("", "VENITURI TOTALE (rd. 12 + 45)", 61, True),
        ("", "CHELTUIELI TOTALE (rd. 35 + 52)", 62, True),
        ("", "PROFITUL BRUT (rd. 61 - 62)", 63, True),
        ("", "PIERDEREA BRUTĂ (rd. 62 - 61)", 64, True),
        ("17.", "Impozitul pe profit / venit (ct. 69x)", 65, False),
        ("", "PROFITUL NET AL PERIOADEI (rd. 63 - 65)", 67, True),
        ("", "PIERDEREA NETĂ A PERIOADEI", 68, True),
    ]
    paired = {36: 37, 37: 36, 53: 54, 54: 53, 63: 64, 64: 63, 67: 68, 68: 67}
    for grp, den, code, is_tot in rows_def:
        v0, v1, v2 = res.rd(code)
        zero = abs(v1) < 1e-6 and abs(v2) < 1e-6
        if zero and code in paired:
            other = res.rd(paired[code])
            if abs(other[1]) > 1e-6 or abs(other[2]) > 1e-6:
                continue  # se afișează doar varianta (profit / pierdere) care are valori
        if zero and not is_tot:
            continue
        res.table_rows.append(CPPRow(grupa=grp, denumire=den, rand=f"{code:02d}", v0=v0, v1=v1, v2=v2, is_total=is_tot))

    return res


def render_sheet_02_cpp(doc: Document, bal: BalantaData, res: CPPResult) -> None:
    """Generează secțiunea Contului de Profit și Pierdere în documentul Word."""
    add_heading(doc, "2. Contul de Profit și Pierdere Sintetic (OMFP nr. 1802/2014)", level=1)

    fl = bal.period.flow_labels
    m2 = max(bal.period.months[2], 1)
    luna = bal.period.month_name or "Luna curentă"
    headers = [
        "Grupa",
        "Indicatorul de venituri, cheltuieli și rezultate",
        "Rd.",
        f"P1 cumulat {fl[1]}",
        f"P2 cumulat {fl[2]}",
        f"Rulaj {luna} (P2 - P1)",
        "Medie lunară P2",
    ]
    rows_data = []
    total_indices = []
    for idx, r in enumerate(res.table_rows):
        if r.is_total:
            total_indices.append(idx)
        rows_data.append([
            r.grupa, r.denumire, r.rand, fmt_money(r.v1), fmt_money(r.v2), fmt_money(r.v2 - r.v1), fmt_money(r.v2 / m2),
        ])

    make_table(
        doc,
        headers=headers,
        rows=rows_data,
        col_widths=[0.9, 6.8, 0.8, 2.3, 2.3, 2.3, 2.1],
        total_rows=total_indices,
        font_size=7.8,
    )

    rn = res.rezultat_net
    rec_parts = []
    for p in (1, 2):
        dif = res.rezultat_121[p] - rn[p]
        if abs(dif) <= TOL:
            rec_parts.append(f"P{p}: rezultatul net din CPP ({fmt_money(rn[p])} lei) este egal cu soldul contului 121")
        else:
            rec_parts.append(
                f"P{p}: rezultatul net din CPP ({fmt_money(rn[p])} lei) diferă de soldul contului 121 "
                f"({fmt_money(res.rezultat_121[p])} lei) cu {fmt_money(dif)} lei — de regulă rezultatul exercițiului "
                f"precedent încă nerepartizat sau înregistrări directe în contul 121"
            )
    rec_txt = "; ".join(rec_parts) + "."
    rec_txt += (
        f" Perioada de raportare 0 reprezintă soldurile de deschidere ale exercițiului {bal.period.year} și nu conține "
        f"rulaje de venituri și cheltuieli; rezultatul exercițiului precedent (sold 121 la 01.01 = "
        f"{fmt_money(res.rezultat_121[0])} lei) apare doar în bilanț."
    )
    if res.note:
        rec_txt += " Încadrări generice: " + "; ".join(res.note) + "."
    add_callout(doc, rec_txt, title="Reconcilierea rezultatului și sfera de cuprindere")

    ca = res.cifra_afaceri
    ve2 = res.venituri_exploatare[2]
    ce2 = res.cheltuieli_exploatare[2]
    re2 = res.rezultat_exploatare[2]
    rn2 = rn[2]
    prod_vand2 = res.rd(2)[2]
    ven_marf2 = res.rd(3)[2]
    ch_pers2 = res.rd(18)[2]
    ch_ext2 = res.rd(27)[2]
    ch_mat2 = res.rd(13)[2] + res.rd(14)[2] + res.rd(15)[2] + res.rd(16)[2] - res.rd(17)[2]

    flow_ca = describe_flow(ca, bal, "cifra de afaceri netă")
    flow_rn = describe_flow(rn, bal, "rezultatul net")

    if all_ok(ca[2]) and ca[2] > 0 and ven_marf2 > prod_vand2 * 1.5:
        profil_ca = f"activitatea având un profil predominant comercial (vânzările de mărfuri reprezintă {fmt_ratio_pct(safe_div(ven_marf2, ca[2]))} din cifra de afaceri)"
    elif all_ok(ca[2]) and ca[2] > 0 and prod_vand2 > ven_marf2 * 1.5:
        profil_ca = f"activitatea fiind susținută prioritar de producția vândută și prestările de servicii ({fmt_ratio_pct(safe_div(prod_vand2, ca[2]))} din cifra de afaceri)"
    else:
        profil_ca = f"structura veniturilor combinând producția vândută ({fmt_money(prod_vand2)} lei) cu vânzările de mărfuri ({fmt_money(ven_marf2)} lei)"

    cost_dom = (
        f"cheltuielile materiale și cu mărfurile ({fmt_money(ch_mat2)} lei)"
        if ch_mat2 >= max(ch_pers2, ch_ext2)
        else (
            f"cheltuielile cu personalul ({fmt_money(ch_pers2)} lei)"
            if ch_pers2 >= ch_ext2
            else f"prestațiile externe, impozitele și alte cheltuieli de exploatare ({fmt_money(ch_ext2)} lei)"
        )
    )

    if re2 >= 0:
        rez_msg = pick_variant(
            bal,
            "s02_rez_poz",
            [
                f"Veniturile din exploatare cumulate până la finele perioadei de raportare 2 ({fmt_money(ve2)} lei) acoperă cheltuielile de exploatare ({fmt_money(ce2)} lei, dominate de {cost_dom}), degajând un rezultat din exploatare de {fmt_money(re2)} lei; după rezultatul financiar ({fmt_money(res.rezultat_financiar[2])} lei) și impozitul pe profit ({fmt_money(res.rd(65)[2])} lei), {flow_rn}.",
                f"Raportul venituri/cheltuieli de exploatare cumulate în perioada de raportare 2 ({fmt_money(ve2)} lei față de {fmt_money(ce2)} lei, elementul dominant de cost fiind {cost_dom}) asigură un rezultat din exploatare de {fmt_money(re2)} lei; pe ansamblu, {flow_rn}.",
            ],
        )
    else:
        rez_msg = (
            f"Cheltuielile de exploatare cumulate până la finele perioadei de raportare 2 ({fmt_money(ce2)} lei, dominate de {cost_dom}) "
            f"depășesc veniturile din exploatare ({fmt_money(ve2)} lei), rezultând o pierdere din exploatare de {fmt_money(-re2)} lei; {flow_rn}."
        )

    add_paragraph(doc, f"Diagnostic al performanței: {flow_ca[0].upper() + flow_ca[1:]}, {profil_ca}. {rez_msg}")
