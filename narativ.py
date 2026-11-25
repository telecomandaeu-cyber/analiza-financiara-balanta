#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Motorul de diversificare narativă al raportului.

Scop: două balanțe diferite să nu primească explicații identice, chiar dacă au aceeași
structură de activitate. Diversificarea se face pe trei niveluri, toate DETERMINISTE
(aceeași balanță produce mereu același text, balanțe diferite produc texte diferite):

1. Amprenta balanței — hash SHA-256 peste TOATE conturile și sumele (nu doar totaluri),
   astfel încât orice diferență de sold schimbă alegerile.
2. Parafrazare la nivel de expresie — fiecare apariție a unei expresii din lexicon
   (verbe, conectori, repere temporale, etichete de paragraf) își alege independent una
   dintre formulările echivalente; la 20–40 de alegeri pe paragraf, numărul de combinații
   este de ordinul milioanelor. Lexiconul conține doar substituții gramatical sigure
   (aceeași formă de persoană, număr și gen).
3. Observații specifice balanței — propoziții construite din conturile concrete ale
   societății (denumiri din balanță, cele mai mari variații, structura cheltuielilor,
   a datoriilor și a trezoreriei), care diferă în conținut de la o firmă la alta.
Numerele, simbolurile de cont și denumirile nu sunt modificate niciodată.
"""

from __future__ import annotations

import hashlib
import re
from typing import Callable, Dict, List, Optional, Sequence, Tuple

_LITERA = "A-Za-zĂÂÎȘȚăâîșțŞŢşţ0-9"


class Narator:
    """Alegeri deterministe de formulare, legate de amprenta unei balanțe."""

    def __init__(self, bal) -> None:
        self.bal = bal
        self._fp = self._fingerprint(bal)
        self._counter = 0
        self._lexicon = self._build_lexicon()
        keys = sorted(self._lexicon, key=len, reverse=True)
        self._re = re.compile(
            r"(?<![" + _LITERA + r"])(" + "|".join(re.escape(k) for k in keys) + r")(?![" + _LITERA + r"])",
            re.IGNORECASE,
        )

    # ------------------------------------------------------------------ amprenta
    @staticmethod
    def _fingerprint(bal) -> str:
        h = hashlib.sha256()
        h.update(f"{bal.cif}|{bal.denumire}|{bal.perioada}".encode("utf-8"))
        for sym in sorted(bal.accounts):
            a = bal.accounts[sym]
            h.update(f"{sym}:{a.sid:.2f}:{a.sic:.2f}:{a.rld:.2f}:{a.rlc:.2f}:{a.tsd:.2f}:{a.tsc:.2f}".encode("utf-8"))
        return h.hexdigest()

    def h(self, *parts: object) -> int:
        s = self._fp + "|" + "|".join(str(p) for p in parts)
        return int(hashlib.sha256(s.encode("utf-8")).hexdigest()[:15], 16)

    def alege(self, salt: str, variante: Sequence[str]) -> str:
        if not variante:
            return ""
        return variante[self.h("alege", salt) % len(variante)]

    def ordine(self, salt: str, parti: Sequence[str]) -> List[str]:
        """Permutare deterministă a unor propoziții independente."""
        return sorted(parti, key=lambda p: self.h("ordine", salt, p))

    # ------------------------------------------------------------------ lexicon
    def _build_lexicon(self) -> Dict[str, List[str]]:
        per = self.bal.period
        d0, d1, d2 = per.dates
        luna = (per.month_name or "").lower()
        an = per.year
        L: Dict[str, List[str]] = {
            # repere temporale
            "în perioada de raportare 2": [
                "în perioada de raportare 2", f"la {d2}", "la finele perioadei de raportare 2", "în P2",
                f"la sfârșitul lunii {luna} {an}", "la data de raportare curentă",
            ],
            "în perioada de raportare 1": [
                "în perioada de raportare 1", f"la {d1}", "la finele perioadei de raportare 1", "în P1",
                "la sfârșitul lunii precedente",
            ],
            "în perioada de raportare 0": [
                "în perioada de raportare 0", f"la {d0}", "la deschiderea exercițiului", "în P0",
                "la începutul anului",
            ],
            "din perioada de raportare 2": ["din perioada de raportare 2", "din P2", f"de la {d2}", "din perioada curentă"],
            "pe ansamblul intervalului": ["pe ansamblul intervalului", "pe întregul orizont analizat", "pe parcursul intervalului studiat", "în intervalul analizat"],
            "pe intervalul analizat": ["pe intervalul analizat", "pe orizontul studiat", "în perioada examinată"],
            # verbe la persoana a III-a (formă identică la singular și plural)
            "indică": ["indică", "arată", "relevă", "evidențiază", "denotă", "sugerează"],
            "atestă": ["atestă", "confirmă", "probează", "demonstrează"],
            "reflectă": ["reflectă", "exprimă", "oglindește"],
            "evidențiază": ["evidențiază", "pune în evidență", "scoate în relief", "relevă"],
            "semnalează": ["semnalează", "indică", "arată"],
            "se situează la": ["se situează la", "se plasează la", "ajunge la", "se cifrează la", "atinge nivelul de"],
            "se situează": ["se situează", "se plasează", "se poziționează"],
            "însumează": ["însumează", "totalizează", "se ridică la"],
            "reprezintă": ["reprezintă", "constituie", "echivalează cu"],
            "înregistrează": ["înregistrează", "consemnează", "marchează"],
            "determină": ["determină", "generează", "conduce la"],
            "asigură": ["asigură", "garantează", "furnizează"],
            "depășește": ["depășește", "devansează", "excede"],
            "crește": ["crește", "se majorează", "avansează", "urcă"],
            "scade": ["scade", "se diminuează", "coboară", "se reduce"],
            "rămâne": ["rămâne", "se menține"],
            # gerunzii
            "semnalând": ["semnalând", "sugerând", "indicând", "punând în evidență"],
            "generând": ["generând", "producând", "antrenând"],
            "reflectând": ["reflectând", "exprimând", "oglindind"],
            # substantive cu același gen și articol
            "creșterea": ["creșterea", "majorarea", "sporirea"],
            "reducerea": ["reducerea", "diminuarea", "contracția"],
            "o creștere": ["o creștere", "o majorare", "o sporire"],
            "o reducere": ["o reducere", "o diminuare", "o contracție"],
            "societatea": ["societatea", "entitatea", "compania", "firma"],
            "societății": ["societății", "entității", "companiei", "firmei"],
            "capacitatea de plată": ["capacitatea de plată", "solvența pe termen scurt", "capacitatea de onorare a obligațiilor"],
            # conectori și locuțiuni
            "în timp ce": ["în timp ce", "pe când", "iar"],
            "față de": ["față de", "comparativ cu", "în raport cu", "prin comparație cu"],
            "ceea ce": ["ceea ce", "fapt ce", "lucru care", "aspect care"],
            "pe ansamblu,": ["pe ansamblu,", "în ansamblu,", "per total,", "global,"],
            "concomitent": ["concomitent", "simultan", "în paralel"],
            "de asemenea": ["de asemenea", "totodată", "în egală măsură"],
            "se recomandă": ["se recomandă", "este indicată", "se impune", "ar fi oportună"],
            "în special": ["în special", "mai ales", "îndeosebi"],
            "în medie": ["în medie", "ca valoare medie"],
            "din care": ["din care", "dintre care", "inclusiv"],
            "se constată": ["se constată", "se observă", "se remarcă", "se poate observa"],
            "în principal": ["în principal", "preponderent", "cu precădere"],
            "integral": ["integral", "în totalitate", "complet"],
            "progresiv": ["progresiv", "treptat", "gradual"],
            "pe fondul": ["pe fondul", "în contextul", "în condițiile"],
            "în contextul": ["în contextul", "pe fondul", "în condițiile"],
            "ca urmare a": ["ca urmare a", "drept consecință a", "pe fondul"],
            "evoluția": ["evoluția", "dinamica", "traiectoria"],
            "evoluției": ["evoluției", "dinamicii", "traiectoriei"],
            "ponderea": ["ponderea", "cota", "proporția"],
            "nivelul": ["nivelul", "palierul"],
            "datoriile pe termen scurt": ["datoriile pe termen scurt", "datoriile curente", "obligațiile pe termen scurt"],
            "datoriilor pe termen scurt": ["datoriilor pe termen scurt", "datoriilor curente", "obligațiilor pe termen scurt"],
            "datorii pe termen scurt": ["datorii pe termen scurt", "datorii curente", "obligații pe termen scurt"],
            "pe parcursul celor trei perioade de raportare": [
                "pe parcursul celor trei perioade de raportare", "de-a lungul celor trei momente analizate",
                "în cele trei perioade de raportare", "pe cele trei repere temporale",
            ],
            "pe ansamblul celor trei perioade de raportare": [
                "pe ansamblul celor trei perioade de raportare", "luând în considerare toate cele trei momente",
                "în ansamblul intervalului analizat",
            ],
            "perioada 0": ["perioada 0", "P0"],
            "perioada 1": ["perioada 1", "P1"],
            "perioada 2": ["perioada 2", "P2"],
        }
        # curăță alternativele construite defensiv
        return {k.lower(): [a for a in v if a] for k, v in L.items()}

    # etichetele de paragraf („Diagnostic bilanțier: ...”) au și ele variante
    ETICHETE: Dict[str, List[str]] = {
        "Diagnostic bilanțier": ["Diagnostic bilanțier", "Lectura bilanțului", "Interpretarea poziției patrimoniale", "Concluzii privind bilanțul"],
        "Diagnostic al performanței": ["Diagnostic al performanței", "Interpretarea contului de profit și pierdere", "Evaluarea rezultatelor", "Lectura performanței"],
        "Diagnostic după retratarea financiară": ["Diagnostic după retratarea financiară", "Interpretarea bilanțului financiar", "Lectura bilanțului retratat"],
        "Analiza structurii activului": ["Analiza structurii activului", "Interpretarea ratelor de activ", "Structura activului — concluzii"],
        "Dinamica structurii financiare": ["Dinamica structurii financiare", "Evoluția surselor de finanțare", "Interpretarea structurii pasivului"],
        "Evaluarea ratelor de structură ale pasivului": ["Evaluarea ratelor de structură ale pasivului", "Autonomie și îndatorare — interpretare", "Lectura ratelor de pasiv"],
        "Diagnostic de echilibru financiar": ["Diagnostic de echilibru financiar", "Interpretarea echilibrului financiar", "Echilibrul financiar — concluzii"],
        "Evaluarea capacității de plată pe termen scurt": ["Evaluarea capacității de plată pe termen scurt", "Interpretarea lichidității", "Lichiditatea — concluzii"],
        "Diagnostic de solvabilitate": ["Diagnostic de solvabilitate", "Interpretarea solvabilității", "Solvabilitatea — concluzii"],
        "Sinteză structurală simplificată": ["Sinteză structurală simplificată", "Lectura bilanțului simplificat", "Structura simplificată — concluzii"],
        "Concluzie privind echilibrul static": ["Concluzie privind echilibrul static", "Echilibrul static — interpretare", "Lectura echilibrului static"],
        "Interpretarea cascadei SIG": ["Interpretarea cascadei SIG", "Lectura soldurilor intermediare de gestiune", "SIG — concluzii"],
        "Sinteză CAF": ["Sinteză CAF", "Interpretarea capacității de autofinanțare", "Autofinanțarea — concluzii"],
        "Verdict bancar": ["Verdict bancar", "Interpretarea capacității de îndatorare", "Îndatorarea — concluzii"],
        "Evaluarea efectului de levier": ["Evaluarea efectului de levier", "Interpretarea levierului financiar", "Levierul financiar — concluzii"],
        "Evaluarea riscului de insolvență": ["Evaluarea riscului de insolvență", "Interpretarea scorurilor de risc", "Riscul de insolvență — concluzii"],
        "Sinteză privind rentabilitatea": ["Sinteză privind rentabilitatea", "Interpretarea ratelor de rentabilitate", "Rentabilitatea — concluzii"],
        "Sinteză statistică": ["Sinteză statistică", "Interpretarea corelațiilor", "Corelațiile — concluzii"],
        "Eficiența gestiunii ciclului operațional": ["Eficiența gestiunii ciclului operațional", "Interpretarea vitezelor de rotație", "Gestiunea ciclului operațional"],
        "Analiza vitezelor de rotație": ["Analiza vitezelor de rotație", "Interpretarea duratelor de rotație", "Rotația elementelor circulante"],
    }

    CONECTORI = ["Totodată, ", "De asemenea, ", "În plus, ", "Complementar, ", "Concomitent, ", "Pe lângă aceasta, ", "Mai departe, ", "Suplimentar, "]
    _NU_LOWER = ("Altman", "Conan", "Pearson", "Spearman", "Legea", "EBE", "CAF", "ROA", "ROE", "ROS", "SIG", "NFR", "TN", "FR", "CPP", "TVA", "ATENȚIE")

    # ------------------------------------------------------------------ parafrazare
    def paraphrase(self, text: str, salt: str = "") -> str:
        if not text or len(text) < 60:
            return text
        self._counter += 1
        pid = f"{salt}#{self._counter}"
        text = self._eticheta(text, pid)
        k = [0]

        def _sub(m: "re.Match[str]") -> str:
            k[0] += 1
            orig = m.group(1)
            alts = self._lexicon.get(orig.lower())
            if not alts:
                return orig
            # verbele folosite reflexiv („se înregistrează”) rămân neschimbate
            if m.string[max(0, m.start() - 3):m.start()].lower() == "se " and not orig.lower().startswith("se "):
                return orig
            alt = alts[self.h("lex", pid, orig.lower(), k[0]) % len(alts)]
            if orig[0].isupper() and alt in ("iar", "inclusiv"):
                return orig  # nu pot deschide o propoziție
            if orig[0].isupper():
                alt = alt[0].upper() + alt[1:]
            return alt

        text = self._re.sub(_sub, text)
        return self._conectori(text, pid)

    def _eticheta(self, text: str, pid: str) -> str:
        m = re.match(r"^([^:]{5,60}):\s", text)
        if not m:
            return text
        lbl = m.group(1)
        alts = self.ETICHETE.get(lbl)
        if not alts:
            return text
        return alts[self.h("eticheta", pid, lbl) % len(alts)] + text[len(lbl):]

    def _conectori(self, text: str, pid: str) -> str:
        parts = re.split(r"(?<=[a-zăâîșț0-9)%»])\.\s+(?=[A-ZĂÂÎȘȚ])", text)
        if len(parts) < 2:
            return text
        out = [parts[0]]
        used = set()
        for i, s in enumerate(parts[1:], start=1):
            first = s.split(" ", 1)[0].rstrip(",:")
            ok = (
                self.h("con?", pid, i) % 100 < 45
                and not first.startswith(self._NU_LOWER)
                and not any(s.startswith(c.strip(", ")) for c in self.CONECTORI)
                and len(first) > 1
                and first[1:2].islower()
                and ":" not in s[:40]
            )
            if ok:
                c = self.CONECTORI[self.h("con", pid, i) % len(self.CONECTORI)]
                if c not in used:
                    used.add(c)
                    s = c + s[0].lower() + s[1:]
            out.append(s)
        return ". ".join(out)


# ==============================================================================
# OBSERVAȚII SPECIFICE BALANȚEI (conținut diferit de la o societate la alta)
# ==============================================================================
def _nume(bal, sym: str) -> str:
    a = bal.accounts.get(sym)
    n = (a.name if a else "").strip()
    n = re.sub(r"\s+", " ", n)
    if len(n) > 70:
        n = n[:67].rsplit(" ", 1)[0] + "…"
    return n[:1].upper() + n[1:].lower() if n else f"contul {sym}"


def _top(items: List[Tuple[str, float]], k: int) -> List[Tuple[str, float]]:
    return sorted(items, key=lambda x: abs(x[1]), reverse=True)[:k]


def _lista(parts: List[str]) -> str:
    if len(parts) <= 1:
        return "".join(parts)
    return ", ".join(parts[:-1]) + " și " + parts[-1]


def observatii_bilant(nar: Narator, fmt_money: Callable[[float], str]) -> Optional[str]:
    bal = nar.bal
    items = []
    for sym in bal.synthetic_accounts():
        if len(sym) != 3 or sym[0] not in "12345":
            continue
        d = bal.net_d(sym, 2) - bal.net_d(sym, 0)
        if abs(d) >= 1.0:
            items.append((sym, d))
    top = _top(items, 3)
    if not top:
        return None
    desc = [
        f"contul {s} «{_nume(bal, s)}» ({'+' if d > 0 else '-'}{fmt_money(abs(d))} lei în sold net debitor)" for s, d in top
    ]
    return nar.alege(
        "obs_bil",
        [
            f"Observații specifice balanței: cele mai ample modificări de sold între {bal.period.dates[0]} și {bal.period.dates[2]} privesc {_lista(desc)}.",
            f"Particularități ale balanței: dinamica patrimonială este explicată în principal de {_lista(desc)}.",
            f"Elemente care diferențiază poziția financiară: variațiile cu cel mai mare impact asupra bilanțului provin din {_lista(desc)}.",
            f"Analiza pe conturi: între deschiderea exercițiului și {bal.period.dates[2]}, mișcările nete cele mai importante se regăsesc la {_lista(desc)}.",
        ],
    )


def observatii_cpp(nar: Narator, fmt_money: Callable[[float], str], fmt_ratio_pct: Callable[[float], str]) -> Optional[str]:
    bal = nar.bal
    ch = [(s, bal.flow_d(s, 2)) for s in bal.synthetic_accounts() if len(s) == 3 and s[0] == "6" and s != "609"]
    ven = [(s, bal.flow_c(s, 2)) for s in bal.synthetic_accounts() if len(s) == 3 and s[0] == "7" and s not in ("711", "709")]
    tot_ch = sum(v for _, v in ch if v > 0)
    tot_ven = sum(v for _, v in ven if v > 0)
    if tot_ch <= 0 or tot_ven <= 0:
        return None
    tch = [(s, v) for s, v in _top(ch, 3) if v > 0]
    tv = _top(ven, 1)[0]
    lch = _lista([f"{s} «{_nume(bal, s)}» ({fmt_ratio_pct(v / tot_ch)})" for s, v in tch])
    conc = sum(v for _, v in tch) / tot_ch
    # contul de cheltuieli cu cea mai mare abatere a lunii curente față de media lunară anterioară
    m1 = bal.period.months[1]
    dev = ""
    if m1 > 0:
        cand = []
        for s, v2 in ch:
            v1 = bal.flow_d(s, 1)
            avg = v1 / m1 if m1 else 0.0
            luna = v2 - v1
            if avg > 0 and v2 > 0.02 * tot_ch:
                cand.append((s, (luna - avg) / avg, luna))
        if cand:
            s, r, luna = max(cand, key=lambda x: abs(x[1]))
            sens = "peste" if r > 0 else "sub"
            dev = nar.alege(
                "obs_cpp_dev",
                [
                    f" În luna curentă, cea mai mare abatere de la ritmul obișnuit apare la contul {s} «{_nume(bal, s)}»: {fmt_money(luna)} lei, cu {fmt_ratio_pct(abs(r))} {sens} media lunară anterioară.",
                    f" Ritmul lunar se modifică cel mai mult la contul {s} «{_nume(bal, s)}», cu un rulaj lunar de {fmt_money(luna)} lei ({fmt_ratio_pct(abs(r))} {sens} medie).",
                    f" Ca abatere lunară notabilă, contul {s} «{_nume(bal, s)}» înregistrează {fmt_money(luna)} lei în ultima lună, {sens} media primelor {m1} luni cu {fmt_ratio_pct(abs(r))}.",
                ],
            )
    return nar.alege(
        "obs_cpp",
        [
            f"Observații specifice balanței: cheltuielile sunt concentrate în {lch}, care împreună reprezintă {fmt_ratio_pct(conc)} din total, iar principala sursă de venit este contul {tv[0]} «{_nume(bal, tv[0])}» ({fmt_ratio_pct(tv[1] / tot_ven)} din venituri).",
            f"Structura costurilor pe conturi: primele poziții sunt {lch} ({fmt_ratio_pct(conc)} cumulat); pe partea de venituri domină contul {tv[0]} «{_nume(bal, tv[0])}», cu {fmt_ratio_pct(tv[1] / tot_ven)}.",
            f"Particularități ale contului de rezultat: veniturile provin în proporție de {fmt_ratio_pct(tv[1] / tot_ven)} din contul {tv[0]} «{_nume(bal, tv[0])}», iar cheltuielile cele mai mari se înregistrează la {lch}.",
            f"Analiza pe conturi a performanței: gradul de concentrare a cheltuielilor este de {fmt_ratio_pct(conc)} pe primele trei conturi — {lch}; veniturile depind în principal de contul {tv[0]} «{_nume(bal, tv[0])}» ({fmt_ratio_pct(tv[1] / tot_ven)}).",
        ],
    ) + dev


def observatii_datorii(nar: Narator, fmt_money: Callable[[float], str], fmt_ratio_pct: Callable[[float], str]) -> Optional[str]:
    bal = nar.bal
    items = []
    for sym in bal.leaf_accounts():
        if sym[0] == "4" and sym[:3] not in ("472", "475", "478"):
            v = bal.net_c(sym, 2)
            if v > 0.5:
                items.append((sym, v))
    tot = sum(v for _, v in items)
    if tot <= 0:
        return None
    top = _top(items, 3)
    lst = _lista([f"{s} «{_nume(bal, s)}» ({fmt_ratio_pct(v / tot)})" for s, v in top])
    return nar.alege(
        "obs_dat",
        [
            f"Componența datoriilor curente din conturi de terți ({fmt_money(tot)} lei) este dominată de {lst}.",
            f"Pe conturi, obligațiile față de terți ({fmt_money(tot)} lei) se concentrează în {lst}.",
            f"Datoriile din clasa 4 însumează {fmt_money(tot)} lei, principalele poziții fiind {lst}.",
        ],
    )


def observatii_trezorerie(nar: Narator, fmt_money: Callable[[float], str], fmt_ratio_pct: Callable[[float], str]) -> Optional[str]:
    bal = nar.bal
    items = [(s, bal.net_d(s, 2)) for s in bal.leaf_accounts() if s[0] == "5" and s[:3] in ("511", "512", "531", "532", "541", "542")]
    items = [(s, v) for s, v in items if v > 0.5]
    tot = sum(v for _, v in items)
    if tot <= 0:
        return None
    lst = _lista([f"{s} «{_nume(bal, s)}» ({fmt_ratio_pct(v / tot)})" for s, v in _top(items, 3)])
    d0 = sum(max(bal.net_d(s, 0), 0) for s, _ in items)
    var = tot - d0
    return nar.alege(
        "obs_trez",
        [
            f"Disponibilitățile de {fmt_money(tot)} lei sunt păstrate în {lst}; față de deschiderea exercițiului, nivelul lor s-a modificat cu {fmt_money(var)} lei.",
            f"Pe conturi de trezorerie, lichiditățile ({fmt_money(tot)} lei) se repartizează astfel: {lst}, cu o variație de {fmt_money(var)} lei față de 01.01.",
            f"Structura numerarului ({fmt_money(tot)} lei): {lst}; modificarea netă de la începutul anului este de {fmt_money(var)} lei.",
        ],
    )
