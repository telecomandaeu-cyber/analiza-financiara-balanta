#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Aplicație Web Streamlit (pentru rulare directă din GitHub prin Streamlit Community Cloud).
Permite încărcarea oricărui fișier Balanta.pdf în browser, afișează pe ecran întregul
conținut al raportului financiar (Tabloul de Bord KPI, Sinteza pe Clase și toate cele
21 de Capitole cu tabelele și interpretările lor complete) și permite descărcarea
documentului în format Microsoft Word (.docx).
"""

from __future__ import annotations

import html
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict

import streamlit as st

CURRENT_DIR = Path(__file__).resolve().parent
ANALIZA_DIR = CURRENT_DIR if (CURRENT_DIR / "core_balanta.py").exists() else (CURRENT_DIR / "Analiza")
if str(ANALIZA_DIR) not in sys.path:
    sys.path.insert(0, str(ANALIZA_DIR))

from site_analiza_financiara import _find_default_pdf, run_analysis_and_build_payload

st.set_page_config(
    page_title="Analiză Financiară din Balanța de Verificare (.DOCX)",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
    .report-table-wrap { overflow-x: auto; margin-bottom: 12px; }
    table.fin-tbl { width: 100%; border-collapse: collapse; font-size: 13px; background: #ffffff; }
    table.fin-tbl th {
        background: #0f2c59; color: #ffffff; font-weight: 600;
        padding: 8px 10px; border: 1px solid #1e3a8a; text-align: right;
    }
    table.fin-tbl th:first-child { text-align: left; }
    table.fin-tbl td {
        padding: 6px 10px; border: 1px solid #e2e8f0; text-align: right; color: #0f172a;
    }
    table.fin-tbl td:first-child { text-align: left; font-weight: 600; color: #1e293b; }
    table.fin-tbl tr:nth-child(even) td { background: #f8fafc; }
    table.fin-tbl tr.total-row td {
        background: #dbeafe !important; font-weight: 700; color: #0f2c59;
    }
    .part-banner-st {
        background: linear-gradient(135deg, #1b365d 0%, #1e4e8c 100%);
        color: #ffffff; padding: 15px 20px; border-radius: 10px;
        margin: 26px 0 14px;
    }
    .part-kicker-st { font-size: 11.5px; font-weight: 700; color: #93c5fd; text-transform: uppercase; letter-spacing: 0.7px; }
    .part-title-st { font-size: 16px; font-weight: 700; margin: 2px 0; color: #ffffff; }
    .part-sub-st { font-size: 12.5px; font-style: italic; color: #e2e8f0; }
    .analysis-para-st {
        background: #f8fafc; border-left: 4px solid #2563eb;
        padding: 12px 16px; border-radius: 0 8px 8px 0;
        font-size: 13.8px; color: #1e293b; text-align: justify; margin: 10px 0 18px;
    }
    .bullet-st {
        background: #f8fafc; border-left: 4px solid #0f2c59;
        padding: 12px 16px; border-radius: 0 8px 8px 0;
        font-size: 13.8px; color: #1e293b; text-align: justify; margin-bottom: 10px;
    }
    </style>
    <div style="background:linear-gradient(135deg,#0f2c59,#1e3a8a);padding:20px 26px;border-radius:12px;color:white;margin-bottom:20px;">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:1.3px;color:#93c5fd;font-weight:700;">
            Diagnostic Economico-Financiar Automatizat
        </div>
        <h2 style="margin:4px 0;color:white;">Generator Raport de Analiză Financiară (.DOCX & Web)</h2>
        <div style="font-size:13px;color:#e2e8f0;">
            21 de capitole interconectate • Perioadele de raportare 0, 1 și 2 • Afișare completă în pagină și export în Microsoft Word
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


def _render_html_table(blk: Dict[str, Any]) -> str:
    headers = blk.get("headers", [])
    rows = blk.get("rows", [])
    first_short = len(headers[0]) <= 6 if headers else False

    th_Parts = []
    for idx, h in enumerate(headers):
        align = "left" if (idx == 0 or (idx == 1 and first_short)) else "right"
        th_Parts.append(f'<th style="text-align:{align}">{html.escape(str(h))}</th>')

    tr_parts = []
    for r_obj in rows:
        tr_cls = ' class="total-row"' if r_obj.get("is_total") else ""
        td_parts = []
        for idx, c_obj in enumerate(r_obj.get("cells", [])):
            txt = html.escape(str(c_obj.get("text", "")))
            bg = c_obj.get("bg")
            fg = c_obj.get("fg")
            align = "left" if (idx == 0 or (idx == 1 and first_short)) else "right"
            if bg and fg:
                style = f"background-color:{bg} !important;color:{fg} !important;font-weight:700;text-align:center;"
            else:
                style = f"text-align:{align};"
            td_parts.append(f'<td style="{style}">{txt}</td>')
        tr_parts.append(f"<tr{tr_cls}>{''.join(td_parts)}</tr>")

    return (
        '<div class="report-table-wrap"><table class="fin-tbl">'
        f"<thead><tr>{''.join(th_Parts)}</tr></thead>"
        f"<tbody>{''.join(tr_parts)}</tbody>"
        "</table></div>"
    )


col_up, col_act = st.columns([2, 1])

with col_up:
    uploaded_file = st.file_uploader(
        "Încarcă fișierul Balanta.pdf (sintetică sau analitică)",
        type=["pdf"],
        help="Selectează balanța de verificare în format PDF pentru a afișa și descărca raportul complet (.docx).",
    )

with col_act:
    st.write("")
    st.write("")
    default_pdf = _find_default_pdf()
    use_default = False
    if default_pdf and default_pdf.exists():
        use_default = st.button("📂 Rulează pe Balanta.pdf demonstrativă", use_container_width=True)

target_pdf_path: Path | None = None
temp_dir_to_clean: Path | None = None

if uploaded_file is not None:
    temp_dir_to_clean = Path(tempfile.mkdtemp(prefix="st_balanta_"))
    target_pdf_path = temp_dir_to_clean / uploaded_file.name
    target_pdf_path.write_bytes(uploaded_file.getvalue())
elif use_default and default_pdf:
    target_pdf_path = default_pdf

if target_pdf_path is not None:
    try:
        with st.spinner("⏳ Se extrag conturile din PDF și se generează raportul complet pe cele 21 de capitole..."):
            data = run_analysis_and_build_payload(target_pdf_path)
            docx_path = Path(data["saved_path"])
            docx_bytes = docx_path.read_bytes()

        comp = data["company"]
        st.success(f"✅ Raportul financiar complet pentru **{comp['denumire']}** ({comp['perioada']}) a fost generat!")

        st.download_button(
            label=f"📥 Descarcă Raportul Word Complet ({data['file_name']})",
            data=docx_bytes,
            file_name=data["file_name"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )

        st.markdown(
            f"**Societate:** {comp['denumire']} &nbsp;|&nbsp; **C.I.F.:** {comp['cif']} &nbsp;|&nbsp; "
            f"**Reg. Com.:** {comp['reg_com']} &nbsp;|&nbsp; **Balanța la:** {comp['perioada']} &nbsp;|&nbsp; "
            f"**Conturi analizate:** {comp['nr_conturi']}"
        )

        rc = data.get("report_content", {})
        if rc.get("intro_text"):
            st.markdown(f'<div class="analysis-para-st">{html.escape(rc["intro_text"])}</div>', unsafe_allow_html=True)

        st.subheader("0.1. Tablou de Bord Executiv — Sinteza Indicatorilor Cheie (Top 12 KPI)")
        kpi_table = [
            {
                "Indicator Cheie (KPI)": r[0],
                "Perioada de raportare 0": r[1],
                "Perioada de raportare 1": r[2],
                "Perioada de raportare 2": r[3],
                "Verdict și Încadrare": r[4],
            }
            for r in data["kpi_rows"]
        ]
        st.dataframe(kpi_table, use_container_width=True, hide_index=True)

        st.subheader("0.2. Sinteza Balanței de Verificare pe Clase de Conturi (Clasa 1 — Clasa 7)")
        cls_table = [
            {
                "Clasa": r[0],
                "Denumire clasă contabilă": r[1],
                "SID (lei)": r[2],
                "SIC (lei)": r[3],
                "Rulaj D": r[4],
                "Rulaj C": r[5],
                "Total Sume D": r[6],
                "Total Sume C": r[7],
                "SFD (lei)": r[8],
                "SFC (lei)": r[9],
            }
            for r in data["cls_rows"]
        ]
        st.dataframe(cls_table, use_container_width=True, hide_index=True)

        # Afișarea integrală a tuturor celor 6 Părți și 21 de Capitole din documentul Word
        for blk in rc.get("blocks", []):
            btype = blk.get("type")
            if btype == "part_banner":
                st.markdown(
                    f"""
                    <div class="part-banner-st">
                        <div class="part-kicker-st">{html.escape(blk.get('kicker', ''))}</div>
                        <div class="part-title-st">{html.escape(blk.get('title', ''))}</div>
                        <div class="part-sub-st">{html.escape(blk.get('subtitle', ''))}</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            elif btype == "heading":
                st.subheader(blk.get("text", ""))
            elif btype == "table":
                st.markdown(_render_html_table(blk), unsafe_allow_html=True)
            elif btype == "paragraph":
                st.markdown(
                    f'<div class="analysis-para-st">{html.escape(blk.get("text", ""))}</div>',
                    unsafe_allow_html=True,
                )
            elif btype == "bullet":
                lead = html.escape(blk.get("lead", ""))
                body = html.escape(blk.get("body", ""))
                st.markdown(
                    f'<div class="bullet-st"><strong>▪ {lead}:</strong> {body}</div>',
                    unsafe_allow_html=True,
                )

    except Exception as exc:
        st.error(f"❌ Eroare la procesarea fișierului PDF: {exc}")
    finally:
        if temp_dir_to_clean and temp_dir_to_clean.exists():
            shutil.rmtree(temp_dir_to_clean, ignore_errors=True)
