import streamlit as st
import pdfplumber
import pandas as pd
import re
import tempfile
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Folha Analítica", layout="centered")

# -------------------------------
# ESTILO PREMIUM
# -------------------------------
st.markdown("""
<style>
    .block-container {
        padding-top: 2rem;
    }

    .stButton>button {
        background: linear-gradient(90deg, #0f62fe, #4589ff);
        color: white;
        border-radius: 10px;
        height: 3em;
        font-weight: 600;
        border: none;
    }

    .stDownloadButton>button {
        background: linear-gradient(90deg, #24a148, #42be65);
        color: white;
        border-radius: 10px;
        height: 3em;
        font-weight: 600;
        border: none;
    }

    .card {
        padding: 1rem;
        border-radius: 12px;
        background-color: #f8f9fa;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------------
# TÍTULO (MANTIDO ❤️)
# -------------------------------
st.title("📄 Processador de Folha Analítica pra Minha Preta (Karem 💍♥️)")
st.caption("Mor, envie um ou mais PDFs e baixe o Excel pronto")

# -------------------------------
# SESSION STATE
# -------------------------------
if "historico" not in st.session_state:
    st.session_state.historico = []

if "arquivos_processados" not in st.session_state:
    st.session_state.arquivos_processados = {}

# -------------------------------
# EXTRAÇÃO
# -------------------------------
def extrair_folha_analitica(pdf_path):

    dados = []

    with pdfplumber.open(pdf_path) as pdf:

        total_paginas = len(pdf.pages)
        progresso = st.progress(0)
        status = st.empty()

        for page_num, page in enumerate(pdf.pages):

            status.text(f"📄 Página {page_num+1} de {total_paginas}")
            progresso.progress((page_num + 1) / total_paginas)

            texto = page.extract_text() or page.extract_text(layout=True)

            if not texto:
                continue

            linhas = texto.split("\n")

            nome_atual = None
            matricula_atual = None

            for linha in linhas:

                linha = linha.strip()
                linha = re.sub(r"\s{2,}", " ", linha)

                mat_match = re.search(r"MAT\.?\s*:?\s*(\d{5,7})", linha)
                if mat_match:
                    matricula_atual = mat_match.group(1)

                nome_match = re.search(
                    r"NOME\s*:?\s*([A-ZÀ-Ú\s]+?)(?:FUNCAO|FUNC|DT|$)",
                    linha
                )
                if nome_match:
                    nome_atual = nome_match.group(1).strip()

                if "|" in linha and re.search(r"\d{3}", linha):

                    partes = linha.split("|")

                    for idx, parte in enumerate(partes):

                        parte = parte.strip()
                        if not parte:
                            continue

                        evento_match = re.match(
                            r"(\d{3})\s+(.+?)\s+([\d,]+)?\s+([\d.,]+)",
                            parte
                        )

                        if evento_match:

                            codigo, desc, ref, valor = evento_match.groups()
                            tipo = "PROVENTO" if idx == 0 else "DESCONTO"

                            valor = valor.replace(".", "").replace(",", ".")

                            try:
                                valor = float(valor)
                            except:
                                continue

                            dados.append({
                                "pagina": page_num + 1,
                                "nome": nome_atual or "SEM_NOME",
                                "matricula": matricula_atual or "SEM_MAT",
                                "tipo": tipo,
                                "codigo": codigo,
                                "descricao": desc.strip(),
                                "valor": valor
                            })

        status.text("✅ Concluído")

    df = pd.DataFrame(dados)

    if not df.empty:
        df["nome"] = df["nome"].str.replace(r"[:\s]+$", "", regex=True).str.strip()

    return df

# -------------------------------
# PROCESSAMENTO PRINCIPAL
# -------------------------------
def processar_pdf(file):

    st.info("🔄 Iniciando processamento... (preta, tenha um pouco de paciência 😂♥️)")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(file.read())
        pdf_path = tmp.name

    df = extrair_folha_analitica(pdf_path)

    pivot = df.pivot_table(
        values="valor",
        index=["nome", "matricula", "tipo"],
        columns="codigo",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    mapa = {
        "455": "Assistência Médica Titular",
        "454": "Assistência Odontológica Titular",
        "458": "Coparticipação",
        "456": "Assistência Médica Dependente",
        "461": "Assistência Odontológica Dependente"
    }

    for col in mapa:
        if col not in pivot.columns:
            pivot[col] = 0

    analise = pivot[["nome", "matricula"] + list(mapa.keys())].rename(columns=mapa)
    analise = analise.groupby(["nome", "matricula"]).sum().reset_index()

    linhas = []

    for _, row in analise.iterrows():

        med_tit = row["Assistência Médica Titular"]
        odo_tit = row["Assistência Odontológica Titular"]
        med_dep = row["Assistência Médica Dependente"]
        odo_dep = row["Assistência Odontológica Dependente"]

        qtd_med = int(round(med_dep / med_tit)) if med_tit > 0 else 0
        qtd_odo = int(round(odo_dep / odo_tit)) if odo_tit > 0 else 0
        qtd = max(qtd_med, qtd_odo)

        linhas.append({
            "nome": row["nome"],
            "matricula": row["matricula"],
            "tipo_registro": "TITULAR",
            "dependente_id": 0,
            "vlr_medico": med_tit,
            "vlr_odonto": odo_tit,
            "vlr_copart": row["Coparticipação"],
            "total": med_tit + odo_tit + row["Coparticipação"]
        })

        for i in range(qtd):

            vlr_med = med_dep / qtd_med if i < qtd_med and qtd_med > 0 else 0
            vlr_odo = odo_dep / qtd_odo if i < qtd_odo and qtd_odo > 0 else 0

            linhas.append({
                "nome": row["nome"],
                "matricula": row["matricula"],
                "tipo_registro": "DEPENDENTE",
                "dependente_id": i + 1,
                "vlr_medico": round(vlr_med, 2),
                "vlr_odonto": round(vlr_odo, 2),
                "vlr_copart": 0,
                "total": round(vlr_med + vlr_odo, 2)
            })

    df_totvs = pd.DataFrame(linhas)

    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Detalhamento", index=False)
        pivot.to_excel(writer, sheet_name="Pivot", index=False)
        analise.to_excel(writer, sheet_name="Analise", index=False)
        df_totvs.to_excel(writer, sheet_name="Base_TOTVS", index=False)

    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb["Base_TOTVS"]

    ultima_linha = ws.max_row

    ws[f"D{ultima_linha+2}"] = "TITULAR"
    ws[f"D{ultima_linha+3}"] = "DEPENDENTE"

    final = BytesIO()
    wb.save(final)
    final.seek(0)

    return final, df_totvs

# -------------------------------
# CACHE (ANTI REPROCESSAMENTO)
# -------------------------------
@st.cache_data(show_spinner=False)
def processar_pdf_cache(file_bytes, file_name):
    return processar_pdf(BytesIO(file_bytes))

# -------------------------------
# UPLOAD
# -------------------------------
uploaded_files = st.file_uploader(
    "📤 Envie um ou mais PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

# -------------------------------
# PROCESSAMENTO UI
# -------------------------------
if uploaded_files:

    for file in uploaded_files:

        st.markdown('<div class="card">', unsafe_allow_html=True)

        col1, col2 = st.columns([3,1])

        with col1:
            st.subheader(f"📄 {file.name}")

        with col2:
            processar = st.button("🚀", key=file.name)

        if processar:

            with st.spinner("Processando com carinho pra você 💙"):
                resultado, df_totvs = processar_pdf_cache(file.getvalue(), file.name)

            st.session_state.arquivos_processados[file.name] = resultado

            st.session_state.historico.append({
                "arquivo": file.name,
                "data": datetime.now().strftime("%d/%m %H:%M"),
                "linhas": len(df_totvs)
            })

            st.success("Prontinho, meu amor! 💚")

            with st.expander("👀 Visualizar dados"):
                st.dataframe(df_totvs.head(50), use_container_width=True)

        if file.name in st.session_state.arquivos_processados:

            st.download_button(
                "⬇️ Baixar Excel",
                st.session_state.arquivos_processados[file.name],
                file_name=f"{file.name.replace('.pdf','')}.xlsx"
            )

        st.markdown('</div>', unsafe_allow_html=True)

# -------------------------------
# HISTÓRICO
# -------------------------------
if st.session_state.historico:
    st.divider()
    st.subheader("📊 Histórico")

    df_hist = pd.DataFrame(st.session_state.historico)

    col1, col2 = st.columns(2)

    col1.metric("Arquivos processados", len(df_hist))
    col2.metric("Total de linhas", df_hist["linhas"].sum())

    st.dataframe(df_hist, use_container_width=True)
