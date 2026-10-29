#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Aplicație Web Streamlit (pentru rulare directă din GitHub prin Streamlit Community Cloud).
Permite încărcarea oricărui fișier Balanta.pdf în browser, vizualizarea Tabloului de Bord KPI
și descărcarea raportului financiar complet în format Word (.docx).
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

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
    <div style="background:linear-gradient(135deg,#0f2c59,#1e3a8a);padding:20px 26px;border-radius:12px;color:white;margin-bottom:20px;">
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:1.3px;color:#93c5fd;font-weight:700;">
            Diagnostic Economico-Financiar Automatizat
        </div>
        <h2 style="margin:4px 0;color:white;">Generator Raport de Analiză Financiară (.DOCX)</h2>
        <div style="font-size:13px;color:#e2e8f0;">
            20 de capitole interconectate • Perioadele de raportare 0, 1 și 2 • Export direct în Microsoft Word
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

col_up, col_act = st.columns([2, 1])

with col_up:
    uploaded_file = st.file_uploader(
        "Încarcă fișierul Balanta.pdf (sintetică sau analitică)",
        type=["pdf"],
        help="Selectează balanța de verificare în format PDF pentru a genera raportul Word (.docx).",
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
        with st.spinner("⏳ Se extrag conturile din PDF și se calculează toate cele 20 de capitole financiare..."):
            data = run_analysis_and_build_payload(target_pdf_path)
            docx_path = Path(data["saved_path"])
            docx_bytes = docx_path.read_bytes()

        comp = data["company"]
        st.success(f"✅ Raportul Word pentru **{comp['denumire']}** ({comp['perioada']}) a fost generat cu succes!")

        st.download_button(
            label=f"📥 Descarcă Raportul Word ({data['file_name']})",
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

    except Exception as exc:
        st.error(f"❌ Eroare la procesarea fișierului PDF: {exc}")
    finally:
        if temp_dir_to_clean and temp_dir_to_clean.exists():
            shutil.rmtree(temp_dir_to_clean, ignore_errors=True)

