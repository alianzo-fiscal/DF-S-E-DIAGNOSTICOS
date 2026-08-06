#!/usr/bin/env python3
"""
Diagnóstico Financeiro Consolidado — app genérico (fora do Claude)
Sobe os balancetes de verificação de UMA empresa ou de VÁRIAS empresas de um
mesmo grupo, consolida Balanço + DRE + DFC, calcula os indicadores e gera
Excel + Word para download. Não depende de nenhuma conta/skill do Claude —
roda localmente (desktop) ou hospedado (link compartilhável), ver README.md.

Este app não é específico de nenhuma empresa ou grupo: use-o para qualquer
cliente enviando os balancetes correspondentes. A lógica de leitura por
código de classificação (plano de contas) foi validada com o Grupo
CLA/Claudio — para uma empresa nova, sempre confira a "Checagem de
consistência" abaixo antes de confiar no resultado (ver nota no topo de
pipeline.py).
"""
import os
import tempfile
import shutil

import streamlit as st
import pandas as pd

import pipeline as p
import docx_report
import pptx_report

st.set_page_config(page_title="Diagnóstico Financeiro Consolidado", layout="wide")

st.title("📊 Diagnóstico Financeiro Consolidado")
st.caption(
    "Envie os balancetes de verificação (.xls/.xlsx) de uma empresa, ou de várias empresas de um "
    "mesmo grupo, um arquivo por empresa por período. O app identifica cada empresa pelo CNPJ e "
    "cada período pela data informada no próprio balancete — não precisa organizar em pastas. "
    "Funciona para qualquer empresa/grupo, não só para um cliente específico."
)

with st.sidebar:
    st.header("1. Empresa ou grupo")
    group_name = st.text_input(
        "Nome para aparecer nos relatórios",
        value="", placeholder="Ex.: Grupo CLA/Claudio, Empresa XYZ Ltda.",
    ) or "Empresa/Grupo"

    st.header("2. Envie os balancetes")
    uploaded = st.file_uploader(
        "Arquivos .xls / .xlsx (pode selecionar vários de uma vez, de todos os períodos)",
        type=["xls", "xlsx"], accept_multiple_files=True,
    )
    n_periodos_override = st.number_input(
        "Número de períodos esperado por empresa (deixe 0 para detectar automaticamente)",
        min_value=0, value=0, step=1,
        help="Ex.: 2 se cada empresa deveria mandar 1º e 2º trimestre. Se algumas empresas "
             "mandarem um número diferente de arquivos, o app avisa qual.",
    )
    run = st.button("Processar", type="primary", use_container_width=True)

if "statements" not in st.session_state:
    st.session_state.statements = None
    st.session_state.indicators = None
    st.session_state.warnings = []

if run:
    if not uploaded:
        st.error("Envie pelo menos um arquivo de balancete antes de processar.")
    else:
        with st.spinner("Lendo balancetes e consolidando…"):
            tmpdir = tempfile.mkdtemp(prefix="app_balancetes_")
            paths = []
            for f in uploaded:
                dest = os.path.join(tmpdir, f.name)
                with open(dest, "wb") as out:
                    out.write(f.getbuffer())
                paths.append(dest)
            companies, load_warnings = p.load_all_balancetes(file_paths=paths)
            if not companies:
                st.error("Não consegui identificar nenhuma empresa nos arquivos enviados. "
                         "Confira se são balancetes de verificação com 'C.N.P.J' e 'Classificação' visíveis.")
            else:
                n_periodos = n_periodos_override or None
                statements = p.consolidate(companies, n_periodos=n_periodos)
                indicators = p.compute_indicators(statements)
                st.session_state.statements = statements
                st.session_state.indicators = indicators
                st.session_state.warnings = load_warnings + statements["warnings"]
                st.session_state.companies_info = {
                    cnpj: {"nome": info["nome"], "n_periodos": len(info["periodos"])}
                    for cnpj, info in companies.items()
                }
            shutil.rmtree(tmpdir, ignore_errors=True)

statements = st.session_state.statements
indicators = st.session_state.indicators

if statements is None:
    st.info("Envie os balancetes na barra lateral e clique em **Processar** para começar.")
    st.stop()

for w in st.session_state.warnings:
    st.warning(w)

st.subheader("Empresas identificadas")
df_companies = pd.DataFrame([
    {"CNPJ": cnpj, "Empresa": info["nome"], "Períodos enviados": info["n_periodos"]}
    for cnpj, info in st.session_state.companies_info.items()
])
st.dataframe(df_companies, use_container_width=True, hide_index=True)

st.subheader("Checagem de consistência")
st.caption("Toda linha deve ficar em ~0,00 — diferenças maiores indicam código contábil não mapeado ou conta nova.")
df_checks = pd.DataFrame(statements["checks"], columns=["Período", "diff BP", "diff DRE", "diff DFC"])
st.dataframe(df_checks.style.format({"diff BP": "{:.2f}", "diff DRE": "{:.2f}", "diff DFC": "{:.2f}"}),
             use_container_width=True, hide_index=True)
max_diff = max(abs(c[1]) for c in statements["checks"]) if statements["checks"] else 0
if max_diff > 1.0:
    st.error(
        "Alguma diferença acima de R$ 1,00 — isso normalmente significa que esta empresa usa um "
        "plano de contas diferente do que o app espera (ver nota no topo de pipeline.py). Números "
        "abaixo podem estar incorretos para esta empresa; não use para apresentação sem investigar."
    )
elif "primeira_vez" not in st.session_state:
    st.info(
        "Primeira vez usando com esta empresa/grupo? Compare pelo menos um período contra o "
        "Balanço/DRE oficial dela antes de confiar nos números — o app foi validado com o Grupo "
        "CLA/Claudio; outra empresa pode usar um plano de contas um pouco diferente."
    )

labels = statements["period_labels"]
ind = indicators["por_periodo"]

st.subheader("Indicadores-chave (último período coberto)")
last = ind[indicators["last_flow_idx"]]
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Receita Líquida", docx_report.fmt_mi(last["receita_liquida_mi"]))
c2.metric("EBITDA", docx_report.fmt_mi(last["ebitda_mi"]))
c3.metric("Margem EBITDA", docx_report.fmt_pct(last["margem_ebitda_pct"]))
c4.metric("Liquidez Corrente", docx_report.fmt_x(last["liquidez_corrente"]))
c5.metric("Endividamento Geral", docx_report.fmt_pct(last["endividamento_geral_pct"]))
c6.metric("Alavancagem (Dív.Líq/EBITDA)", docx_report.fmt_x(indicators["alavancagem_x"]))

st.subheader("Indicadores por período")
rows = []
metric_labels = [
    ("receita_liquida_mi", "Receita Líquida", docx_report.fmt_mi),
    ("ebitda_mi", "EBITDA", docx_report.fmt_mi),
    ("margem_ebitda_pct", "Margem EBITDA", docx_report.fmt_pct),
    ("resultado_liquido_mi", "Resultado Líquido", docx_report.fmt_mi),
    ("ativo_total_mi", "Ativo Total", docx_report.fmt_mi),
    ("pl_total_mi", "Patrimônio Líquido", docx_report.fmt_mi),
    ("liquidez_corrente", "Liquidez Corrente", docx_report.fmt_x),
    ("endividamento_geral_pct", "Endividamento Geral", docx_report.fmt_pct),
    ("cobertura_juros_x", "Cobertura de Juros", docx_report.fmt_x),
    ("caixa_operacional_mi", "Caixa Operacional", docx_report.fmt_mi),
]
for key, label, fmt in metric_labels:
    rows.append([label] + [fmt(x[key]) for x in ind])
df_ind = pd.DataFrame(rows, columns=["Indicador"] + labels)
st.dataframe(df_ind, use_container_width=True, hide_index=True)

with st.expander("Ver Balanço, DRE e DFC completos"):
    tab1, tab2, tab3 = st.tabs(["Balanço", "DRE", "DFC"])
    def render_statement(tab, row_template, values):
        with tab:
            data = []
            for label, depth, key in row_template:
                if key is None:
                    data.append([label] + [""] * len(values))
                else:
                    data.append([("    " * depth) + label] + [f"{v.get(key, 0.0):,.2f}" for v in values])
            st.dataframe(pd.DataFrame(data, columns=["Descrição"] + labels), use_container_width=True, hide_index=True)
    render_statement(tab1, statements["bp_rows"], statements["bp"])
    render_statement(tab2, statements["dre_rows"], statements["dre"])
    render_statement(tab3, statements["dfc_rows"], statements["dfc"])

st.subheader("2. Baixe os entregáveis")
col_a, col_b, col_c = st.columns(3)

safe_name = "".join(c if c.isalnum() or c in " -_" else "" for c in group_name).strip().replace(" ", "_") or "Empresa"

with col_a:
    if st.button("Gerar Excel consolidado (Balanço + DRE + DFC)"):
        with st.spinner("Gerando Excel…"):
            out_path = os.path.join(tempfile.mkdtemp(), "consolidado.xlsx")
            p.build_excel(statements, out_path, group_name=group_name)
            with open(out_path, "rb") as f:
                st.session_state["excel_bytes"] = f.read()
    if "excel_bytes" in st.session_state:
        st.download_button("⬇️ Baixar Excel", st.session_state["excel_bytes"],
                            file_name=f"BP_DRE_DFC_{safe_name}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

with col_b:
    if st.button("Gerar relatório Word (diagnóstico)"):
        with st.spinner("Gerando Word…"):
            out_path = os.path.join(tempfile.mkdtemp(), "diagnostico.docx")
            docx_report.build_report(statements, indicators, out_path, group_name=group_name)
            with open(out_path, "rb") as f:
                st.session_state["docx_bytes"] = f.read()
    if "docx_bytes" in st.session_state:
        st.download_button("⬇️ Baixar Word", st.session_state["docx_bytes"],
                            file_name=f"Diagnostico_Financeiro_{safe_name}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with col_c:
    if st.button("Gerar apresentação PPT (diagnóstico)"):
        with st.spinner("Gerando PPT…"):
            out_path = os.path.join(tempfile.mkdtemp(), "diagnostico.pptx")
            pptx_report.build_presentation(statements, indicators, out_path, group_name=group_name)
            with open(out_path, "rb") as f:
                st.session_state["pptx_bytes"] = f.read()
    if "pptx_bytes" in st.session_state:
        st.download_button("⬇️ Baixar PPT", st.session_state["pptx_bytes"],
                            file_name=f"Diagnostico_Financeiro_{safe_name}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

st.divider()
st.caption(
    "Este relatório é gerado automaticamente a partir dos balancetes enviados. Os pontos de atenção "
    "seguem regras de referência de mercado (ver THRESHOLDS em docx_report.py) e não substituem uma "
    "revisão humana — especialmente para explicar reclassificações contábeis, baixas de fornecedor ou "
    "correções pontuais de saldo ocorridas no período."
)
