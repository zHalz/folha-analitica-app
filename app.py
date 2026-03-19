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
# CSS PREMIUM DARK 💖
# -------------------------------
st.markdown("""
<style>

body {
    background-color: #0e1117;
}

.block-container {
    max-width: 750px;
    margin: auto;
    padding-top: 2rem;
}

/* HERO */
.hero {
    text-align: center;
    margin-bottom: 2rem;
}
.hero h1 {
    color: white;
}
.hero p {
    color: #aaa;
}

/* CARDS */
.card {
    background: #161b22;
    padding: 1.2rem;
    border-radius: 14px;
    margin-bottom: 1rem;
    border: 1px solid #222;
}

/* UPLOADER */
[data-testid="stFileUploader"] {
    border: none !important;
    background: transparent !important;
}

/* BOTÕES */
.stButton>button {
    background: linear-gradient(90deg, #ff4d6d, #ff758f);
    color: white;
    border-radius: 10px;
    height: 2.6em;
    width: 100%;
    border: none;
}

.stDownloadButton>button {
    background: #24a148;
    color: white;
    border-radius: 10px;
    height: 2.6em;
    width: 100%;
    border: none;
}

/* TEXTO */
h3, h2 {
    color: white;
}

</style>
""", unsafe_allow_html=True)

# -------------------------------
# HERO 💖
# -------------------------------
st.markdown("""
<div class="hero">
    <h1>📄 Processador de Folha Analítica pra Minha Preta (Karem 💍♥️)</h1>
    <p>Mor, envia aqui que eu resolvo tudo pra você rapidinho 😘</p>
</div>
""", unsafe_allow_html=True)

# -------------------------------
# SESSION
# -------------------------------
if "historico" not in st.session_state:
    st.session_state.historico = []

if "arquivos_processados" not in st.session_state:
    st.session_state.arquivos_processados = {}

# -------------------------------
# EXTRAÇÃO (COM PROGRESSO)
# -------------------------------
def extrair_folha_analitica(pdf_path):

    dados = []

    with pdfplumber.open(pdf_path) as pdf:

        total_paginas = len(pdf.pages)
        progress_bar = st.progress(0)
        status = st.empty()

        for page_num, page in enumerate(pdf.pages):

            status.info(f"📄 Página {page_num+1} de {total_paginas}")
            progress_bar.progress((page_num + 1) / total_paginas)

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

        status.success("✅ Processamento concluído!")

    return pd.DataFrame(dados)

# -------------------------------
# PROCESSAMENTO
# -------------------------------
def processar_pdf(file):

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
        df_totvs.to_excel(writer, sheet_name="Base_TOTVS", index=False)

    buffer.seek(0)

    wb = load_workbook(buffer)
    ws = wb["Base_TOTVS"]

    ultima_linha = ws.max_row

    linha_titular = ultima_linha + 2
    linha_dependente = ultima_linha + 3

    ws[f"D{linha_titular}"] = "TITULAR"
    ws[f"D{linha_dependente}"] = "DEPENDENTE"

    range_e = f"E2:E{ultima_linha}"
    range_f = f"F2:F{ultima_linha}"
    range_g = f"G2:G{ultima_linha}"
    range_h = f"H2:H{ultima_linha}"
    range_c = f"C2:C{ultima_linha}"

    # TITULAR
    ws[f"E{linha_titular}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(E2,ROW({range_e})-ROW(E2),0)),({range_c}=D{linha_titular})+0)'
    ws[f"F{linha_titular}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(F2,ROW({range_f})-ROW(F2),0)),({range_c}=D{linha_titular})+0)'
    ws[f"G{linha_titular}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(G2,ROW({range_g})-ROW(G2),0)),({range_c}=D{linha_titular})+0)'
    ws[f"H{linha_titular}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(H2,ROW({range_h})-ROW(H2),0)),({range_c}=D{linha_titular})+0)'

    # DEPENDENTE
    ws[f"E{linha_dependente}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(E2,ROW({range_e})-ROW(E2),0)),({range_c}=D{linha_dependente})+0)'
    ws[f"F{linha_dependente}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(F2,ROW({range_f})-ROW(F2),0)),({range_c}=D{linha_dependente})+0)'
    ws[f"G{linha_dependente}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(G2,ROW({range_g})-ROW(G2),0)),({range_c}=D{linha_dependente})+0)'
    ws[f"H{linha_dependente}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(H2,ROW({range_h})-ROW(H2),0)),({range_c}=D{linha_dependente})+0)'

    final = BytesIO()
    wb.save(final)
    final.seek(0)

    return final, df_totvs

# -------------------------------
# UPLOAD CARD
# -------------------------------
st.markdown('<div class="card">', unsafe_allow_html=True)
st.subheader("💌 Envie seus PDFs")

uploaded_files = st.file_uploader(
    "Arraste aqui",
    type=["pdf"],
    accept_multiple_files=True
)

st.markdown('</div>', unsafe_allow_html=True)

# -------------------------------
# ARQUIVOS
# -------------------------------
if uploaded_files:

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📄 Arquivos")

    for file in uploaded_files:

        col1, col2, col3 = st.columns([4,1,1])

        col1.write(file.name)

        if col2.button("Processar", key=file.name):

            status = st.empty()
            status.info("🔄 Iniciando processamento... (preta, tenha um pouco de paciência 😂♥️)")

            resultado, df_totvs = processar_pdf(file)

            status.success("Prontinho, meu amor 💚")

            st.session_state.arquivos_processados[file.name] = resultado

        if file.name in st.session_state.arquivos_processados:

            col3.download_button(
                "Baixar",
                st.session_state.arquivos_processados[file.name],
                file_name=file.name.replace(".pdf", ".xlsx")
            )

    st.markdown('</div>', unsafe_allow_html=True)

# -------------------------------
# HISTÓRICO
# -------------------------------
if st.session_state.historico:

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📊 Histórico")

    df_hist = pd.DataFrame(st.session_state.historico)
    st.dataframe(df_hist, use_container_width=True)

    st.markdown('</div>', unsafe_allow_html=True)

