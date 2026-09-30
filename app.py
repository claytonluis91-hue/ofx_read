from __future__ import annotations

import hashlib
from pathlib import Path

import pandas as pd
import streamlit as st

from src.ofx_service import build_excel, parse_ofx
from src.pdf_service import build_ofx, parse_pdf


st.set_page_config(page_title="Nexus • Conferência Financeira", page_icon="↗", layout="wide")

st.markdown(
    """
    <style>
    :root { --navy:#17324D; --blue:#1F6E8C; --teal:#159A9C; --mist:#F4F8FA; }
    .stApp { background: linear-gradient(180deg, #F7FAFC 0%, #FFFFFF 45%); }
    .block-container { max-width: 1180px; padding-top: 2rem; padding-bottom: 4rem; }
    h1, h2, h3 { color: var(--navy); letter-spacing: -.02em; }
    div[data-testid="stMetric"] { background:#fff; border:1px solid #E4E7EC; padding:16px 18px; border-radius:14px; box-shadow:0 6px 20px rgba(23,50,77,.05); }
    div[data-testid="stFileUploader"] { background:#fff; border:1px solid #D6E2E8; border-radius:16px; padding:10px; }
    .hero { background:linear-gradient(120deg,#17324D,#1F6E8C); color:white; border-radius:20px; padding:28px 32px; margin-bottom:24px; box-shadow:0 14px 36px rgba(23,50,77,.18); }
    .hero h1 { color:white; margin:0 0 6px 0; font-size:2.1rem; }
    .hero p { margin:0; color:#DDECF1; font-size:1.02rem; }
    .note { padding:14px 16px; border-left:4px solid #159A9C; background:#EFF8F8; border-radius:8px; color:#344054; }
    .stButton > button, .stDownloadButton > button { border-radius:10px; font-weight:650; }
    div[data-testid="stTabs"] button { font-weight:650; }
    footer { visibility:hidden; }
    </style>
    <div class="hero">
      <h1>Nexus • Conferência Financeira</h1>
      <p>Converta extratos com segurança, revise os lançamentos e baixe arquivos prontos para o seu fluxo financeiro.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.markdown("### Central de conversão")
    st.caption("Os arquivos são processados somente durante a sessão e não são enviados para serviços externos.")
    st.divider()
    st.markdown("**Formatos disponíveis**")
    st.markdown("OFX → Excel (.xlsx)  \nPDF Bradesco → OFX")
    st.divider()
    st.caption("Revise sempre os totais e os dados bancários antes da importação.")


def money(value: float) -> str:
    text = f"{abs(value):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{'-' if value < 0 else ''}R$ {text}"


def safe_stem(filename: str) -> str:
    return Path(filename).stem.replace(" ", "_")


def restart_conversion() -> None:
    """Limpa uploads, edições e resultados da sessão atual."""
    st.session_state.clear()
    st.rerun()


tab_excel, tab_ofx, tab_help = st.tabs(["OFX → Excel", "PDF → OFX", "Como usar"])

with tab_excel:
    st.subheader("Gerar planilha de conferência")
    st.write("Envie um ou mais arquivos OFX. Cada arquivo gera uma planilha formatada com filtros, totais e marcação de conferência.")
    ofx_files = st.file_uploader("Selecione os arquivos OFX", type=["ofx"], accept_multiple_files=True, key="ofx_upload")
    if ofx_files:
        for uploaded in ofx_files:
            try:
                statement = parse_ofx(uploaded.getvalue())
                frame = statement.transactions
                credits = float(frame.loc[frame["Valor"] > 0, "Valor"].sum())
                debits = float(frame.loc[frame["Valor"] < 0, "Valor"].sum())
                with st.container(border=True):
                    st.markdown(f"#### {uploaded.name}")
                    c1, c2, c3, c4 = st.columns(4)
                    c1.metric("Lançamentos", len(frame))
                    c2.metric("Entradas", money(credits))
                    c3.metric("Saídas", money(debits))
                    c4.metric("Resultado", money(credits + debits))
                    preview = frame[["Data", "Tipo", "Valor", "Descrição", "Histórico"]].head(12).copy()
                    st.dataframe(
                        preview,
                        hide_index=True,
                        use_container_width=True,
                        column_config={
                            "Data": st.column_config.DateColumn(format="DD/MM/YYYY"),
                            "Valor": st.column_config.NumberColumn(format="R$ %.2f"),
                        },
                    )
                    excel = build_excel(statement, uploaded.name)
                    download_col, restart_col = st.columns([3, 1])
                    with download_col:
                        st.download_button(
                            "Baixar planilha de conferência",
                            data=excel,
                            file_name=f"{safe_stem(uploaded.name)}_conferencia.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key=f"xlsx_{hashlib.md5(uploaded.getvalue()).hexdigest()}",
                            type="primary",
                            use_container_width=True,
                        )
                    with restart_col:
                        if st.button(
                            "Reiniciar",
                            key=f"restart_xlsx_{hashlib.md5(uploaded.getvalue()).hexdigest()}",
                            use_container_width=True,
                        ):
                            restart_conversion()
            except Exception as exc:
                st.error(f"Não foi possível converter **{uploaded.name}**: {exc}")
    else:
        st.info("Envie um arquivo .ofx para visualizar os lançamentos e liberar o download da planilha.")

with tab_ofx:
    st.subheader("Criar OFX a partir do extrato em PDF")
    st.write("O conversor reconhece o layout de extrato Bradesco Empresas usado nos arquivos de exemplo.")
    pdf_file = st.file_uploader("Selecione o extrato em PDF", type=["pdf"], key="pdf_upload")
    if pdf_file:
        try:
            digest = hashlib.md5(pdf_file.getvalue()).hexdigest()
            cache_key = f"pdf_statement_{digest}"
            if cache_key not in st.session_state:
                st.session_state[cache_key] = parse_pdf(pdf_file.getvalue())
            pdf_statement = st.session_state[cache_key]
            frame = pdf_statement.transactions.copy()
            credits = float(frame.loc[frame["Valor"] > 0, "Valor"].sum())
            debits = float(frame.loc[frame["Valor"] < 0, "Valor"].sum())
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Lançamentos", len(frame))
            c2.metric("Entradas", money(credits))
            c3.metric("Saídas", money(debits))
            latest_date = frame["Data"].max()
            closing_balance = float(frame.loc[frame["Data"] == latest_date].iloc[-1]["Saldo"])
            c4.metric("Saldo final", money(closing_balance))
            for warning in pdf_statement.warnings:
                st.warning(warning)

            st.markdown("#### Dados bancários")
            m1, m2, m3 = st.columns(3)
            bank = m1.text_input("Código do banco", value=pdf_statement.metadata.get("Banco", "237"), key=f"bank_{digest}")
            branch = m2.text_input("Agência", value=pdf_statement.metadata.get("Agência", ""), key=f"branch_{digest}")
            account = m3.text_input("Conta", value=pdf_statement.metadata.get("Conta", ""), key=f"account_{digest}")
            holder = st.text_input("Titular", value=pdf_statement.metadata.get("Titular", ""), key=f"holder_{digest}")

            st.markdown("#### Revise os lançamentos")
            st.caption("Você pode corrigir data, descrição, documento, valor e saldo antes de gerar o OFX.")
            edited = st.data_editor(
                frame,
                hide_index=True,
                use_container_width=True,
                num_rows="dynamic",
                key=f"editor_{digest}",
                column_config={
                    "Data": st.column_config.DateColumn(required=True, format="DD/MM/YYYY"),
                    "Tipo": st.column_config.TextColumn(disabled=True),
                    "Valor": st.column_config.NumberColumn(required=True, format="R$ %.2f"),
                    "Saldo": st.column_config.NumberColumn(format="R$ %.2f"),
                },
            )
            edited["Tipo"] = edited["Valor"].apply(lambda value: "Crédito" if float(value) >= 0 else "Débito")
            metadata = dict(pdf_statement.metadata)
            metadata.update({"Banco": bank, "Agência": branch, "Conta": account, "Titular": holder})
            generated_ofx = build_ofx(edited, metadata)
            download_col, restart_col = st.columns([3, 1])
            with download_col:
                st.download_button(
                    "Baixar arquivo OFX",
                    data=generated_ofx,
                    file_name=f"{safe_stem(pdf_file.name)}_convertido.ofx",
                    mime="application/x-ofx",
                    key=f"ofx_{digest}",
                    type="primary",
                    use_container_width=True,
                )
            with restart_col:
                if st.button("Reiniciar", key=f"restart_ofx_{digest}", use_container_width=True):
                    restart_conversion()
        except Exception as exc:
            st.error(f"Não foi possível converter o PDF: {exc}")
    else:
        st.info("Envie um PDF com texto selecionável. PDFs escaneados ainda precisam passar por OCR antes da conversão.")

with tab_help:
    st.subheader("Fluxo recomendado")
    st.markdown(
        """
        1. **OFX para Excel:** envie o OFX, confira o resumo e baixe a planilha. No Excel, marque cada linha como “Sim” na coluna **Conferido**.
        2. **PDF para OFX:** envie o extrato, confira agência, conta e lançamentos, corrija o que for necessário e baixe o OFX.
        3. **Antes de importar:** compare quantidade de lançamentos, entradas, saídas e saldo final com o documento original.
        """
    )
    st.markdown('<div class="note">O PDF → OFX foi calibrado para o extrato Bradesco Empresas presente na pasta <b>Exemplos</b>. Outros layouts podem exigir um adaptador específico.</div>', unsafe_allow_html=True)
