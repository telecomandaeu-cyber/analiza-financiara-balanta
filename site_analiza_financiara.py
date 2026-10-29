#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Aplicație Web Locală — Analiza Poziției Financiare și a Performanțelor Financiare.
Rulează local în browser (fără dependențe externe suplimentare), permite încărcarea
oricărui fișier Balanta.pdf (sau rularea fișierului existent), afișează pe ecran
Tabloul de Bord Executiv pe cele trei perioade de raportare (0, 1 și 2) și exportă
documentul complet în format Word (.docx).
"""

from __future__ import annotations

import argparse
import io
import json
import os
import shutil
import socket
import sys
import tempfile
import threading
import traceback
import urllib.parse
import uuid
import webbrowser
from email.parser import BytesParser
from email.policy import default as email_default_policy
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

CURRENT_DIR = Path(__file__).resolve().parent
ANALIZA_DIR = CURRENT_DIR if (CURRENT_DIR / "core_balanta.py").exists() else (CURRENT_DIR / "Analiza")
if str(ANALIZA_DIR) not in sys.path:
    sys.path.insert(0, str(ANALIZA_DIR))

from analiza_financiara import generate_full_financial_report
from core_balanta import fmt_money, fmt_num, fmt_pct, fmt_ratio_pct, parse_balanta_pdf
from sheet_01_bilant import compute_sheet_01_bilant
from sheet_02_cpp import compute_sheet_02_cpp
from sheet_03_bilant_financiar import compute_sheet_03_bilant_financiar
from sheet_05_structura_financiara import compute_sheet_05_structura_financiara
from sheet_06_ratele_pasivului import compute_sheet_06_ratele_pasivului
from sheet_07_echilibrul_financiar import compute_sheet_07_echilibrul_financiar
from sheet_08_lichiditatea import compute_sheet_08_lichiditatea
from sheet_09_solvabilitatea import compute_sheet_09_solvabilitatea
from sheet_13_sig import compute_sheet_13_sig
from sheet_14_caf import compute_sheet_14_caf
from sheet_16_efectul_indatorare import compute_sheet_16_efectul_indatorare
from sheet_17_riscul_faliment import compute_sheet_17_riscul_faliment
from sheet_19_indicatori_profitabilitate import compute_sheet_19_indicatori_profitabilitate

# Registru în memorie cu rapoartele generate și lacăt pentru execuție concurentă în rețea
GENERATED_FILES: Dict[str, Path] = {}
REPORT_LOCK = threading.Lock()
SERVER_PORT: int = 8080


def _get_lan_ips() -> List[str]:
    """Returnează lista adreselor IPv4 din rețeaua locală (LAN / Wi-Fi) ale acestui calculator."""
    ips: List[str] = []
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("10.255.255.255", 1))
            primary_ip = s.getsockname()[0]
            if primary_ip and not primary_ip.startswith("127."):
                ips.append(primary_ip)
    except Exception:
        pass

    try:
        hostname = socket.gethostname()
        for ip in socket.gethostbyname_ex(hostname)[2]:
            if ip and not ip.startswith("127.") and ip not in ips:
                ips.append(ip)
    except Exception:
        pass

    return ips or ["127.0.0.1"]


def _find_default_pdf() -> Optional[Path]:
    candidates = [
        ANALIZA_DIR.parent / "Balanta.pdf",
        ANALIZA_DIR / "Balanta.pdf",
        CURRENT_DIR / "Balanta.pdf",
        Path(r"c:\Users\Florentin\Desktop\Import\Utilitare\Balanta.pdf"),
    ]
    for c in candidates:
        if c.exists():
            return c
    return None


def run_analysis_and_build_payload(pdf_path: Path) -> Dict[str, Any]:
    """
    Execută analiza financiară completă, generează fișierul Word (.docx)
    și returnează structura JSON pentru afișarea Tabloului de Bord în pagina web.
    """
    bal = parse_balanta_pdf(pdf_path)

    s01_bil = compute_sheet_01_bilant(bal)
    s02_cpp = compute_sheet_02_cpp(bal, s01_bil)
    s03_bf = compute_sheet_03_bilant_financiar(s01_bil)
    s05_sf = compute_sheet_05_structura_financiara(s03_bf)
    s06_rp = compute_sheet_06_ratele_pasivului(s05_sf)
    s07_ef = compute_sheet_07_echilibrul_financiar(s03_bf, s05_sf)
    s08_liq = compute_sheet_08_lichiditatea(s03_bf, s01_bil)
    s09_solv = compute_sheet_09_solvabilitatea(s03_bf, s05_sf, s01_bil)
    s13_sig = compute_sheet_13_sig(s02_cpp)
    s14_caf = compute_sheet_14_caf(s02_cpp, s13_sig)
    s16_lev = compute_sheet_16_efectul_indatorare(s03_bf, s05_sf, s02_cpp)
    s17_risc = compute_sheet_17_riscul_faliment(s01_bil, s02_cpp, s03_bf, s05_sf, s07_ef, s13_sig)
    s19_prof = compute_sheet_19_indicatori_profitabilitate(s03_bf, s02_cpp)

    with REPORT_LOCK:
        docx_path = generate_full_financial_report(pdf_path)
        file_id = f"{uuid.uuid4().hex[:8]}_{docx_path.name}"
        GENERATED_FILES[file_id] = docx_path

    g_cp = (s05_sf.cp[2] - s05_sf.cp[0]) / abs(s05_sf.cp[0]) if abs(s05_sf.cp[0]) > 1e-6 else 0.0
    g_rn = (s13_sig.rezultat_net[2] - s13_sig.rezultat_net[0]) / abs(s13_sig.rezultat_net[0]) if abs(s13_sig.rezultat_net[0]) > 1e-6 else 0.0
    g_caf = (s14_caf.caf_deductiv[2] - s14_caf.caf_deductiv[0]) / abs(s14_caf.caf_deductiv[0]) if abs(s14_caf.caf_deductiv[0]) > 1e-6 else 0.0

    cards = [
        {
            "title": "Activ Total = Pasiv Total",
            "p0": f"{fmt_money(s03_bf.total_activ[0])} lei",
            "p1": f"{fmt_money(s03_bf.total_activ[1])} lei",
            "p2": f"{fmt_money(s03_bf.total_activ[2])} lei",
            "badge": "Echilibru Bilanțier",
            "tone": "good",
        },
        {
            "title": "Capitaluri Proprii (Activ Net)",
            "p0": f"{fmt_money(s05_sf.cp[0])} lei",
            "p1": f"{fmt_money(s05_sf.cp[1])} lei",
            "p2": f"{fmt_money(s05_sf.cp[2])} lei",
            "badge": f"Dinamică: {fmt_ratio_pct(g_cp, plus_sign=True)}",
            "tone": "good" if s05_sf.cp[2] > 0 else "bad",
        },
        {
            "title": "Cifra de Afaceri Netă",
            "p0": f"{fmt_money(s19_prof.blocks[0].den_vals[0])} lei",
            "p1": f"{fmt_money(s19_prof.blocks[0].den_vals[1])} lei",
            "p2": f"{fmt_money(s19_prof.blocks[0].den_vals[2])} lei",
            "badge": f"ROS: {fmt_pct(s19_prof.ros[2])}",
            "tone": "good",
        },
        {
            "title": "Profit Net (Rezultatul Exercițiului)",
            "p0": f"{fmt_money(s13_sig.rezultat_net[0])} lei",
            "p1": f"{fmt_money(s13_sig.rezultat_net[1])} lei",
            "p2": f"{fmt_money(s13_sig.rezultat_net[2])} lei",
            "badge": f"ROE: {fmt_pct(s19_prof.roe[2])} | ROA: {fmt_pct(s19_prof.roa[2])}",
            "tone": "good" if s13_sig.rezultat_net[2] > 0 else "bad",
        },
        {
            "title": "Fond de Rulment / Trezoreria Netă",
            "p0": f"TN: {fmt_money(s07_ef.tn[0])} lei",
            "p1": f"TN: {fmt_money(s07_ef.tn[1])} lei",
            "p2": f"TN: {fmt_money(s07_ef.tn[2])} lei",
            "badge": f"FR P2: {fmt_money(s07_ef.fr_sus[2])} lei",
            "tone": "good" if s07_ef.tn[2] >= 0 else "warn",
        },
        {
            "title": "Capacitatea de Autofinanțare (CAF)",
            "p0": f"{fmt_money(s14_caf.caf_deductiv[0])} lei",
            "p1": f"{fmt_money(s14_caf.caf_deductiv[1])} lei",
            "p2": f"{fmt_money(s14_caf.caf_deductiv[2])} lei",
            "badge": f"Scor Altman Z: {fmt_num(s17_risc.z_altman_ajustat[2], 2)}",
            "tone": "good" if s14_caf.caf_deductiv[2] > 0 else "bad",
        },
    ]

    kpi_rows = [
        [
            "1. Activ Total Bilanțier = Pasiv Total (lei)",
            fmt_money(s03_bf.total_activ[0]),
            fmt_money(s03_bf.total_activ[1]),
            fmt_money(s03_bf.total_activ[2]),
            "Corelație optimă",
        ],
        [
            "2. Capitaluri Proprii — Activ Net Contabil (lei)",
            fmt_money(s05_sf.cp[0]),
            fmt_money(s05_sf.cp[1]),
            fmt_money(s05_sf.cp[2]),
            f"{'EFTL' if s05_sf.cp[2] > 0 else 'Dezechilibru'} ({fmt_ratio_pct(g_cp, plus_sign=True)})",
        ],
        [
            "3. Rata Stabilității Financiare — RSF (%)",
            fmt_ratio_pct(s06_rp.rsf[0]),
            fmt_ratio_pct(s06_rp.rsf[1]),
            fmt_ratio_pct(s06_rp.rsf[2]),
            "Semnal favorabil (>50%)" if s06_rp.rsf[2] >= 0.50 else "Sub 50%",
        ],
        [
            "4. Fondul de Rulment Patrimonial — FR (lei)",
            fmt_money(s07_ef.fr_sus[0]),
            fmt_money(s07_ef.fr_sus[1]),
            fmt_money(s07_ef.fr_sus[2]),
            "EFTL (FR > 0)" if s07_ef.fr_sus[2] >= 0 else "DFTL (FR < 0)",
        ],
        [
            "5. Necesarul de Fond de Rulment — NFR (lei)",
            fmt_money(s07_ef.nfr[0]),
            fmt_money(s07_ef.nfr[1]),
            fmt_money(s07_ef.nfr[2]),
            "EFTS (FR > NFR)" if s07_ef.fr_sus[2] >= s07_ef.nfr[2] else "DFTS (FR < NFR)",
        ],
        [
            "6. Trezoreria Netă — TN = FR - NFR (lei)",
            fmt_money(s07_ef.tn[0]),
            fmt_money(s07_ef.tn[1]),
            fmt_money(s07_ef.tn[2]),
            "EFC (TN > 0)" if s07_ef.tn[2] >= 0 else "DFC (TN < 0)",
        ],
        [
            "7. Lichiditatea Curentă / Lichiditatea Imediată (%)",
            f"{fmt_ratio_pct(s08_liq.l_curenta[0])} / {fmt_ratio_pct(s08_liq.l_imediata[0])}",
            f"{fmt_ratio_pct(s08_liq.l_curenta[1])} / {fmt_ratio_pct(s08_liq.l_imediata[1])}",
            f"{fmt_ratio_pct(s08_liq.l_curenta[2])} / {fmt_ratio_pct(s08_liq.l_imediata[2])}",
            "Da (În plafon)" if 1.2 <= s08_liq.l_curenta[2] <= 2.0 else "În afara intervalului [120%-200%]",
        ],
        [
            "8. Solvabilitatea Generală / Solvabilitatea Efectivă (%)",
            f"{fmt_ratio_pct(s09_solv.solv_generala[0])} / {fmt_ratio_pct(s09_solv.solv_efectiva[0])}",
            f"{fmt_ratio_pct(s09_solv.solv_generala[1])} / {fmt_ratio_pct(s09_solv.solv_efectiva[1])}",
            f"{fmt_ratio_pct(s09_solv.solv_generala[2])} / {fmt_ratio_pct(s09_solv.solv_efectiva[2])}",
            "Da (Solvabilitate >150%)" if s09_solv.solv_generala[2] >= 1.5 else "Sub plafonul de 150%",
        ],
        [
            "9. Cifra de Afaceri Netă / Profitul Net (lei)",
            f"{fmt_money(s19_prof.blocks[0].den_vals[0])} / {fmt_money(s13_sig.rezultat_net[0])}",
            f"{fmt_money(s19_prof.blocks[0].den_vals[1])} / {fmt_money(s13_sig.rezultat_net[1])}",
            f"{fmt_money(s19_prof.blocks[0].den_vals[2])} / {fmt_money(s13_sig.rezultat_net[2])}",
            f"{'Profitabil' if s13_sig.rezultat_net[2] > 0 else 'Pierdere'} ({fmt_ratio_pct(g_rn, plus_sign=True)})",
        ],
        [
            "10. Capacitatea de Autofinanțare — CAF (lei)",
            fmt_money(s14_caf.caf_deductiv[0]),
            fmt_money(s14_caf.caf_deductiv[1]),
            fmt_money(s14_caf.caf_deductiv[2]),
            f"{'Da' if s14_caf.caf_deductiv[2] > 0 else 'Nu'} ({fmt_ratio_pct(g_caf, plus_sign=True)})",
        ],
        [
            "11. Rentabilități: ROA / ROE / Efect de Levier",
            f"{fmt_pct(s19_prof.roa[0])} / {fmt_pct(s19_prof.roe[0])} / {fmt_ratio_pct(s16_lev.efect_levier[0], plus_sign=True)}",
            f"{fmt_pct(s19_prof.roa[1])} / {fmt_pct(s19_prof.roe[1])} / {fmt_ratio_pct(s16_lev.efect_levier[1], plus_sign=True)}",
            f"{fmt_pct(s19_prof.roa[2])} / {fmt_pct(s19_prof.roe[2])} / {fmt_ratio_pct(s16_lev.efect_levier[2], plus_sign=True)}",
            s16_lev.verdict[2],
        ],
        [
            "12. Scorul Z Altman (Standard / Ajustat fără ct. 472+475)",
            f"{fmt_num(s17_risc.z_altman[0], 2)} / {fmt_num(s17_risc.z_altman_ajustat[0], 2)}",
            f"{fmt_num(s17_risc.z_altman[1], 2)} / {fmt_num(s17_risc.z_altman_ajustat[1], 2)}",
            f"{fmt_num(s17_risc.z_altman[2], 2)} / {fmt_num(s17_risc.z_altman_ajustat[2], 2)}",
            "Solvabil (Z > 3,00)" if s17_risc.z_altman_ajustat[2] > 3.0 else ("Zonă incertă [1,8-3,0]" if s17_risc.z_altman_ajustat[2] >= 1.8 else "Risc insolvență (<1,8)"),
        ],
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

    return {
        "ok": True,
        "company": {
            "denumire": bal.denumire,
            "cif": bal.cif,
            "reg_com": bal.reg_com,
            "adresa": bal.adresa,
            "perioada": bal.perioada,
            "nr_conturi": len([k for k in bal.accounts if not k.startswith("Clasa")]),
        },
        "download_url": f"/download?file={urllib.parse.quote(file_id)}",
        "file_name": docx_path.name,
        "saved_path": str(docx_path),
        "cards": cards,
        "kpi_rows": kpi_rows,
        "cls_rows": cls_rows,
    }


HTML_PAGE = r"""<!DOCTYPE html>
<html lang="ro">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Platformă Locală — Analiza Poziției și Performanțelor Financiare</title>
  <style>
    :root {
      --navy: #0f2c59;
      --navy-light: #1e3a8a;
      --blue-accent: #2563eb;
      --teal: #0d9488;
      --bg: #f1f5f9;
      --card: #ffffff;
      --text: #0f172a;
      --muted: #475569;
      --border: #cbd5e1;
      --success-bg: #dcfce7;
      --success-txt: #166534;
      --warn-bg: #fef3c7;
      --warn-txt: #92400e;
      --bad-bg: #fee2e2;
      --bad-txt: #991b1b;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.5;
      padding-bottom: 48px;
    }
    header {
      background: linear-gradient(135deg, var(--navy) 0%, var(--navy-light) 100%);
      color: #fff;
      padding: 24px 32px;
      box-shadow: 0 4px 16px rgba(15, 44, 89, 0.18);
    }
    .header-inner {
      max-width: 1240px;
      margin: 0 auto;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
    }
    .brand-tag {
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 1.4px;
      color: #93c5fd;
      font-weight: 700;
      margin-bottom: 4px;
    }
    h1 {
      font-size: 22px;
      font-weight: 700;
      letter-spacing: -0.3px;
    }
    .header-sub {
      font-size: 13px;
      color: #e2e8f0;
      margin-top: 2px;
    }
    .badge-local {
      background: rgba(255, 255, 255, 0.14);
      border: 1px solid rgba(255, 255, 255, 0.28);
      padding: 8px 14px;
      border-radius: 999px;
      font-size: 12px;
      font-weight: 600;
      color: #f8fafc;
    }
    .container {
      max-width: 1240px;
      margin: 24px auto 0;
      padding: 0 20px;
    }
    .upload-panel {
      background: var(--card);
      border-radius: 14px;
      padding: 24px;
      box-shadow: 0 2px 10px rgba(15, 23, 42, 0.06);
      border: 1px solid var(--border);
      display: grid;
      grid-template-columns: 1.4fr 1fr;
      gap: 22px;
      align-items: stretch;
    }
    @media (max-width: 860px) {
      .upload-panel { grid-template-columns: 1fr; }
    }
    .dropzone {
      border: 2px dashed #64748b;
      border-radius: 12px;
      background: #f8fafc;
      padding: 28px 20px;
      text-align: center;
      cursor: pointer;
      transition: all 0.2s ease;
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      gap: 10px;
    }
    .dropzone:hover, .dropzone.dragover {
      border-color: var(--blue-accent);
      background: #eff6ff;
    }
    .dropzone-icon {
      width: 46px;
      height: 46px;
      border-radius: 50%;
      background: #dbeafe;
      color: var(--blue-accent);
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 22px;
      font-weight: bold;
    }
    .dropzone-title {
      font-size: 15px;
      font-weight: 700;
      color: var(--navy);
    }
    .dropzone-hint {
      font-size: 12.5px;
      color: var(--muted);
    }
    .selected-file {
      margin-top: 6px;
      font-size: 13px;
      font-weight: 600;
      color: var(--teal);
      background: #ccfbf1;
      padding: 4px 12px;
      border-radius: 6px;
      display: none;
    }
    .actions-box {
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      background: #f8fafc;
      border-radius: 12px;
      padding: 20px;
      border: 1px solid #e2e8f0;
      gap: 14px;
    }
    .actions-box h3 {
      font-size: 15px;
      color: var(--navy);
      margin-bottom: 4px;
    }
    .actions-box p {
      font-size: 12.5px;
      color: var(--muted);
    }
    .checkbox-row {
      display: flex;
      align-items: center;
      gap: 8px;
      font-size: 13px;
      color: var(--text);
      font-weight: 500;
      cursor: pointer;
    }
    .btn-group {
      display: flex;
      flex-direction: column;
      gap: 10px;
    }
    .btn {
      border: none;
      border-radius: 9px;
      padding: 12px 18px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.18s ease;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      text-decoration: none;
    }
    .btn-primary {
      background: var(--navy);
      color: #ffffff;
    }
    .btn-primary:hover:not(:disabled) {
      background: var(--navy-light);
      transform: translateY(-1px);
    }
    .btn-secondary {
      background: #e2e8f0;
      color: var(--navy);
    }
    .btn-secondary:hover:not(:disabled) {
      background: #cbd5e1;
    }
    .btn-success {
      background: #15803d;
      color: #ffffff;
      box-shadow: 0 4px 12px rgba(21, 128, 61, 0.25);
    }
    .btn-success:hover {
      background: #166534;
      transform: translateY(-1px);
    }
    .btn:disabled {
      opacity: 0.6;
      cursor: not-allowed;
    }
    .status-bar {
      margin-top: 16px;
      padding: 14px 18px;
      border-radius: 10px;
      font-size: 13.5px;
      font-weight: 600;
      display: none;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
    }
    .status-loading {
      display: flex;
      background: #dbeafe;
      color: #1e40af;
      border: 1px solid #93c5fd;
    }
    .status-error {
      display: flex;
      background: var(--bad-bg);
      color: var(--bad-txt);
      border: 1px solid #fca5a5;
    }
    .status-ok {
      display: flex;
      background: var(--success-bg);
      color: var(--success-txt);
      border: 1px solid #86efac;
    }
    #resultsArea {
      display: none;
      margin-top: 24px;
    }
    .company-banner {
      background: var(--card);
      border-radius: 12px;
      padding: 18px 22px;
      border-left: 6px solid var(--navy);
      box-shadow: 0 2px 8px rgba(15, 23, 42, 0.05);
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      margin-bottom: 20px;
    }
    .company-info h2 {
      font-size: 19px;
      color: var(--navy);
      margin-bottom: 4px;
    }
    .company-meta {
      font-size: 13px;
      color: var(--muted);
      display: flex;
      flex-wrap: wrap;
      gap: 14px;
    }
    .company-meta span strong {
      color: var(--text);
    }
    .cards-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(340px, 1fr));
      gap: 16px;
      margin-bottom: 24px;
    }
    .metric-card {
      background: var(--card);
      border-radius: 12px;
      padding: 16px 18px;
      border: 1px solid var(--border);
      box-shadow: 0 1px 6px rgba(15, 23, 42, 0.04);
    }
    .metric-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
    }
    .metric-title {
      font-size: 13px;
      font-weight: 700;
      color: var(--navy);
      text-transform: uppercase;
      letter-spacing: 0.3px;
    }
    .metric-badge {
      font-size: 11.5px;
      font-weight: 700;
      padding: 3px 9px;
      border-radius: 999px;
    }
    .tone-good { background: var(--success-bg); color: var(--success-txt); }
    .tone-warn { background: var(--warn-bg); color: var(--warn-txt); }
    .tone-bad { background: var(--bad-bg); color: var(--bad-txt); }
    .metric-periods {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 8px;
      background: #f8fafc;
      padding: 10px;
      border-radius: 8px;
      border: 1px solid #e2e8f0;
    }
    .period-col {
      text-align: center;
    }
    .period-label {
      font-size: 11px;
      color: var(--muted);
      font-weight: 600;
    }
    .period-val {
      font-size: 13.5px;
      font-weight: 700;
      color: var(--text);
      margin-top: 2px;
    }
    .section-card {
      background: var(--card);
      border-radius: 12px;
      padding: 20px;
      border: 1px solid var(--border);
      box-shadow: 0 2px 8px rgba(15, 23, 42, 0.04);
      margin-bottom: 22px;
      overflow-x: auto;
    }
    .section-card h3 {
      font-size: 16px;
      color: var(--navy);
      margin-bottom: 12px;
      padding-bottom: 8px;
      border-bottom: 2px solid #e2e8f0;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12.8px;
    }
    th {
      background: var(--navy);
      color: #ffffff;
      font-weight: 600;
      padding: 9px 10px;
      text-align: right;
      border: 1px solid #1e3a8a;
    }
    th:first-child { text-align: left; }
    td {
      padding: 7px 10px;
      border: 1px solid #e2e8f0;
      text-align: right;
    }
    td:first-child {
      text-align: left;
      font-weight: 600;
      color: #1e293b;
    }
    tr:nth-child(even) td { background: #f8fafc; }
    tr.total-row td {
      background: #dbeafe !important;
      font-weight: 700;
      color: var(--navy);
    }
  </style>
</head>
<body>
  <header>
    <div class="header-inner">
      <div>
        <div class="brand-tag">Diagnostic Economico-Financiar Automatizat</div>
        <h1>Generator Raport de Analiză Financiară (.DOCX)</h1>
        <div class="header-sub">20 de capitole interconectate • Perioadele de raportare 0, 1 și 2 • Accesibil în Rețeaua Locală (LAN / Wi-Fi)</div>
      </div>
      <div style="display:flex; flex-direction:column; align-items:flex-end; gap:6px;">
        <div class="badge-local" id="lanBadge">● Acces Rețea Locală: se detectează...</div>
        <button onclick="copyLanUrl()" id="copyLanBtn" style="background:rgba(255,255,255,0.18); color:#fff; border:1px solid rgba(255,255,255,0.35); border-radius:6px; padding:4px 10px; font-size:11.5px; font-weight:600; cursor:pointer;">
          📋 Copiază link-ul pentru colegii din rețea
        </button>
      </div>
    </div>
  </header>

  <div class="container">
    <div class="upload-panel">
      <div class="dropzone" id="dropzone" onclick="document.getElementById('fileInput').click()">
        <input type="file" id="fileInput" accept=".pdf" style="display:none" />
        <div class="dropzone-icon">📄</div>
        <div class="dropzone-title">Trage aici fișierul Balanta.pdf sau apasă pentru a-l selecta</div>
        <div class="dropzone-hint">Acceptă orice balanță de verificare sintetică/analitică în format PDF (cu SID, SIC, Rulaje, TSD, TSC, SFD, SFC)</div>
        <div class="selected-file" id="selectedFileLabel"></div>
      </div>

      <div class="actions-box">
        <div>
          <h3>Execuție și Export Raport Word</h3>
          <p>La apăsarea butonului, scriptul extrage conturile din PDF, calculează toate cele 20 de capitole și generează raportul Word compact.</p>
        </div>

        <label class="checkbox-row">
          <input type="checkbox" id="autoDownloadCheck" checked />
          Descarcă automat fișierul Word (.docx) la finalul analizei
        </label>

        <div class="btn-group">
          <button class="btn btn-primary" id="btnRunUpload" onclick="runUploadedPdf()">
            ⚡ Analizează PDF-ul Selectat și Exportă Word
          </button>
          <button class="btn btn-secondary" id="btnRunDefault" onclick="runDefaultPdf()">
            📂 Rulează direct pe Balanta.pdf din folderul Utilitare
          </button>
        </div>
      </div>
    </div>

    <div id="statusBar" class="status-bar">
      <span id="statusText"></span>
      <a id="statusDownloadBtn" class="btn btn-success" href="#" style="display:none; padding: 8px 14px; font-size: 13px;">
        ⬇️ Descarcă Raportul Word (.docx)
      </a>
    </div>

    <div id="resultsArea">
      <div class="company-banner">
        <div class="company-info">
          <h2 id="compName">—</h2>
          <div class="company-meta">
            <span><strong>C.I.F.:</strong> <span id="compCif">—</span></span>
            <span><strong>Reg. Com.:</strong> <span id="compReg">—</span></span>
            <span><strong>Balanța la:</strong> <span id="compPerioada">—</span></span>
            <span><strong>Conturi analizate:</strong> <span id="compConturi">—</span></span>
          </div>
        </div>
        <a id="mainDownloadBtn" class="btn btn-success" href="#">
          📥 Descarcă Raportul Word (.docx)
        </a>
      </div>

      <div class="cards-grid" id="cardsGrid"></div>

      <div class="section-card">
        <h3>0.1. Tablou de Bord Executiv — Sinteza Indicatorilor Cheie (Top 12 KPI)</h3>
        <table>
          <thead>
            <tr>
              <th>Indicator Cheie de Performanță și Echilibru (KPI)</th>
              <th>Perioada de raportare 0</th>
              <th>Perioada de raportare 1</th>
              <th>Perioada de raportare 2</th>
              <th style="text-align:center">Verdict și Încadrare</th>
            </tr>
          </thead>
          <tbody id="kpiTableBody"></tbody>
        </table>
      </div>

      <div class="section-card">
        <h3>0.2. Sinteza Balanței de Verificare pe Clase de Conturi (Clasa 1 — Clasa 7)</h3>
        <table>
          <thead>
            <tr>
              <th>Clasa</th>
              <th style="text-align:left">Denumire clasă contabilă</th>
              <th>SID (lei)</th>
              <th>SIC (lei)</th>
              <th>Rulaj D</th>
              <th>Rulaj C</th>
              <th>Total Sume D</th>
              <th>Total Sume C</th>
              <th>SFD (lei)</th>
              <th>SFC (lei)</th>
            </tr>
          </thead>
          <tbody id="clsTableBody"></tbody>
        </table>
      </div>
    </div>
  </div>

  <script>
    const dropzone = document.getElementById('dropzone');
    const fileInput = document.getElementById('fileInput');
    const selectedFileLabel = document.getElementById('selectedFileLabel');
    let selectedFile = null;
    let primaryLanUrl = window.location.origin;

    async function loadNetworkInfo() {
      try {
        const res = await fetch('/api/network-info');
        const info = await res.json();
        if (info && info.lan_urls && info.lan_urls.length > 0) {
          primaryLanUrl = info.lan_urls[0];
          document.getElementById('lanBadge').textContent = '● Accesibil în rețea: ' + info.lan_urls.join('  |  ');
        }
      } catch (e) {
        document.getElementById('lanBadge').textContent = '● Accesibil la: ' + window.location.origin;
      }
    }
    loadNetworkInfo();

    function copyLanUrl() {
      const btn = document.getElementById('copyLanBtn');
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(primaryLanUrl);
      } else {
        const tempInput = document.createElement('input');
        tempInput.value = primaryLanUrl;
        document.body.appendChild(tempInput);
        tempInput.select();
        document.execCommand('copy');
        tempInput.remove();
      }
      btn.textContent = '✅ Link copiat: ' + primaryLanUrl;
      setTimeout(() => {
        btn.textContent = '📋 Copiază link-ul pentru colegii din rețea';
      }, 2800);
    }

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        selectedFile = e.target.files[0];
        selectedFileLabel.style.display = 'inline-block';
        selectedFileLabel.textContent = 'Fișier pregătit: ' + selectedFile.name;
      }
    });

    dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropzone.classList.add('dragover');
    });
    dropzone.addEventListener('dragleave', () => {
      dropzone.classList.remove('dragover');
    });
    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        selectedFile = e.dataTransfer.files[0];
        fileInput.files = e.dataTransfer.files;
        selectedFileLabel.style.display = 'inline-block';
        selectedFileLabel.textContent = 'Fișier pregătit: ' + selectedFile.name;
      }
    });

    function setStatus(mode, message, downloadUrl) {
      const bar = document.getElementById('statusBar');
      const txt = document.getElementById('statusText');
      const dlBtn = document.getElementById('statusDownloadBtn');
      bar.className = 'status-bar';
      dlBtn.style.display = 'none';
      if (mode === 'loading') {
        bar.classList.add('status-loading');
      } else if (mode === 'error') {
        bar.classList.add('status-error');
      } else if (mode === 'ok') {
        bar.classList.add('status-ok');
        if (downloadUrl) {
          dlBtn.href = downloadUrl;
          dlBtn.style.display = 'inline-flex';
        }
      }
      txt.textContent = message;
    }

    function setButtonsDisabled(disabled) {
      document.getElementById('btnRunUpload').disabled = disabled;
      document.getElementById('btnRunDefault').disabled = disabled;
    }

    async function runUploadedPdf() {
      if (!selectedFile) {
        setStatus('error', 'Te rugăm să selectezi mai întâi un fișier PDF (sau apasă pe butonul gri pentru a rula Balanta.pdf existentă).');
        return;
      }
      const formData = new FormData();
      formData.append('pdf_file', selectedFile);
      await executeAnalysis('/api/analyze-upload', { method: 'POST', body: formData });
    }

    async function runDefaultPdf() {
      await executeAnalysis('/api/analyze-default', { method: 'POST' });
    }

    async function executeAnalysis(url, fetchOptions) {
      try {
        setButtonsDisabled(true);
        setStatus('loading', '⏳ Se extrag datele din Balanta.pdf și se calculează cele 20 de capitole financiare...');
        const response = await fetch(url, fetchOptions);
        const data = await response.json();
        if (!response.ok || !data.ok) {
          throw new Error(data.error || 'Eroare la procesarea fișierului PDF.');
        }
        renderDashboard(data);
        setStatus('ok', '✅ Raportul Word (' + data.file_name + ') a fost generat cu succes!', data.download_url);

        if (document.getElementById('autoDownloadCheck').checked && data.download_url) {
          const a = document.createElement('a');
          a.href = data.download_url;
          a.download = data.file_name;
          document.body.appendChild(a);
          a.click();
          a.remove();
        }
      } catch (err) {
        setStatus('error', '❌ Eroare: ' + err.message);
      } finally {
        setButtonsDisabled(false);
      }
    }

    function renderDashboard(data) {
      document.getElementById('resultsArea').style.display = 'block';
      document.getElementById('compName').textContent = data.company.denumire;
      document.getElementById('compCif').textContent = data.company.cif;
      document.getElementById('compReg').textContent = data.company.reg_com;
      document.getElementById('compPerioada').textContent = data.company.perioada;
      document.getElementById('compConturi').textContent = data.company.nr_conturi;
      document.getElementById('mainDownloadBtn').href = data.download_url;

      const cardsGrid = document.getElementById('cardsGrid');
      cardsGrid.innerHTML = '';
      data.cards.forEach(c => {
        const toneClass = c.tone === 'good' ? 'tone-good' : (c.tone === 'warn' ? 'tone-warn' : 'tone-bad');
        const div = document.createElement('div');
        div.className = 'metric-card';
        div.innerHTML = `
          <div class="metric-header">
            <span class="metric-title">${c.title}</span>
            <span class="metric-badge ${toneClass}">${c.badge}</span>
          </div>
          <div class="metric-periods">
            <div class="period-col">
              <div class="period-label">Perioada 0</div>
              <div class="period-val">${c.p0}</div>
            </div>
            <div class="period-col">
              <div class="period-label">Perioada 1</div>
              <div class="period-val">${c.p1}</div>
            </div>
            <div class="period-col">
              <div class="period-label">Perioada 2</div>
              <div class="period-val">${c.p2}</div>
            </div>
          </div>
        `;
        cardsGrid.appendChild(div);
      });

      const kpiBody = document.getElementById('kpiTableBody');
      kpiBody.innerHTML = '';
      data.kpi_rows.forEach(r => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${r[0]}</td>
          <td>${r[1]}</td>
          <td>${r[2]}</td>
          <td><strong>${r[3]}</strong></td>
          <td style="text-align:center; font-weight:600; color:#0f2c59;">${r[4]}</td>
        `;
        kpiBody.appendChild(tr);
      });

      const clsBody = document.getElementById('clsTableBody');
      clsBody.innerHTML = '';
      data.cls_rows.forEach((r, idx) => {
        const tr = document.createElement('tr');
        if (idx === data.cls_rows.length - 1) tr.className = 'total-row';
        tr.innerHTML = `
          <td style="text-align:center">${r[0]}</td>
          <td style="text-align:left">${r[1]}</td>
          <td>${r[2]}</td>
          <td>${r[3]}</td>
          <td>${r[4]}</td>
          <td>${r[5]}</td>
          <td>${r[6]}</td>
          <td>${r[7]}</td>
          <td>${r[8]}</td>
          <td>${r[9]}</td>
        `;
        clsBody.appendChild(tr);
      });
    }
  </script>
</body>
</html>
"""


def _parse_multipart_pdf(content_type: str, body: bytes) -> Tuple[Optional[str], Optional[bytes]]:
    """Extrage numele fișierului și conținutul binar din multipart/form-data folosind biblioteca standard."""
    raw_headers = f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8")
    msg = BytesParser(policy=email_default_policy).parsebytes(raw_headers + body)
    if not msg.is_multipart():
        return None, None
    for part in msg.iter_parts():
        fn = part.get_filename()
        if fn:
            return fn, part.get_payload(decode=True)
    return None, None


class FinancialWebHandler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: Any) -> None:
        # Afișare curată în consolă
        sys.stdout.write(f"[WEB] {self.address_string()} - {format % args}\n")

    def _send_json(self, status: int, payload: Dict[str, Any]) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        if parsed.path in ("/", "/index.html"):
            body = HTML_PAGE.encode("utf-8")
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if parsed.path == "/api/network-info":
            lan_ips = _get_lan_ips()
            lan_urls = [f"http://{ip}:{SERVER_PORT}" for ip in lan_ips]
            self._send_json(
                HTTPStatus.OK,
                {
                    "ok": True,
                    "port": SERVER_PORT,
                    "lan_ips": lan_ips,
                    "lan_urls": lan_urls,
                },
            )
            return

        if parsed.path == "/download":
            qs = urllib.parse.parse_qs(parsed.query)
            file_id = qs.get("file", [""])[0]
            docx_path = GENERATED_FILES.get(file_id)
            if not docx_path or not docx_path.exists():
                self.send_error(HTTPStatus.NOT_FOUND, "Fisierul solicitat nu a fost gasit.")
                return

            data = docx_path.read_bytes()
            self.send_response(HTTPStatus.OK)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{docx_path.name}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
            return

        self.send_error(HTTPStatus.NOT_FOUND, "Ruta nu exista.")

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == "/api/analyze-default":
            try:
                default_pdf = _find_default_pdf()
                if not default_pdf or not default_pdf.exists():
                    self._send_json(
                        HTTPStatus.BAD_REQUEST,
                        {"ok": False, "error": "Nu a fost găsit fișierul Balanta.pdf în folderul Utilitare."},
                    )
                    return
                payload = run_analysis_and_build_payload(default_pdf)
                self._send_json(HTTPStatus.OK, payload)
            except Exception as exc:
                traceback.print_exc()
                self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"ok": False, "error": str(exc)})
            return

        if parsed.path == "/api/analyze-upload":
            try:
                content_type = self.headers.get("Content-Type", "")
                content_len = int(self.headers.get("Content-Length", "0"))
                if content_len <= 0:
                    self._send_json(HTTPStatus.BAD_REQUEST, {"ok": False, "error": "Fișierul încărcat este gol."})
                    return

                body = self.rfile.read(content_len)
                filename, file_bytes = _parse_multipart_pdf(content_type, body)
                if not file_bytes:
                    self._send_json(
                        HTTPStatus.BAD_REQUEST,
                        {"ok": False, "error": "Nu s-a putut citi fișierul PDF încărcat."},
                    )
                    return

                temp_dir = Path(tempfile.mkdtemp(prefix="balanta_upload_"))
                safe_name = Path(filename or "Balanta.pdf").name
                temp_pdf = temp_dir / safe_name
                temp_pdf.write_bytes(file_bytes)

                try:
                    payload = run_analysis_and_build_payload(temp_pdf)
                finally:
                    shutil.rmtree(temp_dir, ignore_errors=True)

                self._send_json(HTTPStatus.OK, payload)
            except Exception as exc:
                traceback.print_exc()
                self._send_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"ok": False, "error": str(exc)})
            return

        self._send_json(HTTPStatus.NOT_FOUND, {"ok": False, "error": "Endpoint necunoscut."})


def _find_free_port(preferred: int = 8080) -> int:
    for port in range(preferred, preferred + 20):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            if s.connect_ex(("127.0.0.1", port)) != 0:
                return port
    return preferred


def main() -> None:
    global SERVER_PORT
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    default_port = int(os.environ.get("PORT", "8080"))
    parser = argparse.ArgumentParser(description="Pornește site-ul în rețeaua locală pentru Analiza Financiară (.docx)")
    parser.add_argument("--host", type=str, default="0.0.0.0", help="Adresa de ascultare (implicit: 0.0.0.0 pentru toată rețeaua locală)")
    parser.add_argument("--port", type=int, default=default_port, help="Portul local (implicit: 8080 sau variabila de mediu PORT)")
    parser.add_argument("--no-browser", action="store_true", help="Nu deschide automat browserul la pornire")
    args = parser.parse_args()

    is_cloud_env = bool(os.environ.get("PORT") or os.environ.get("CODESPACES") or os.environ.get("GITHUB_ACTIONS"))
    port = args.port if is_cloud_env else _find_free_port(args.port)
    SERVER_PORT = port
    local_url = f"http://127.0.0.1:{port}"
    lan_ips = _get_lan_ips()
    lan_urls = [f"http://{ip}:{port}" for ip in lan_ips]

    server = ThreadingHTTPServer((args.host, port), FinancialWebHandler)
    print("=" * 76)
    print("  PLATFORMA WEB — ANALIZA POZIȚIEI ȘI PERFORMANȚELOR FINANCIARE (LAN + LOCAL)")
    print("-" * 76)
    print(f"  -> Acces de pe acest calculator : {local_url}")
    for u in lan_urls:
        print(f"  -> Acces pentru colegii din rețea: {u}")
    print("-" * 76)
    print("  Notă: Dacă alte calculatoare din rețea nu pot accesa linkul de mai sus,")
    print("        rulează 'Permite_Acces_Retea_Firewall.bat' (Click dreapta -> Run as administrator).")
    print("  Apasă Ctrl+C în această fereastră pentru a opri serverul.")
    print("=" * 76)

    open_target = lan_urls[0] if lan_urls else local_url
    if not args.no_browser and not is_cloud_env:
        threading.Timer(0.8, lambda: webbrowser.open(open_target)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[INFO] Serverul a fost oprit.")
        server.server_close()


if __name__ == "__main__":
    main()

