import streamlit as st
import pdfplumber
import pandas as pd
import re
import tempfile
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook

st.set_page_config(page_title="Folha Analítica", layout="wide")

# -------------------------------
# CSS MODERNO DARK E FLUIDO 💖
# -------------------------------
st.markdown("""
<style>
:root {
    --bg-dark: #0e1117;
    --surface-dark: #161b22;
    --secondary-text: #bbb;
    --accent-start: #ff4d6d;
    --accent-end: #ff758f;
    --success: #24a148;
}

body {
    background-color: var(--bg-dark);
    color: white;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}

/* Espaçamento e altura mínima */
.block-container {
    min-height: 100vh;
    padding-top: 1rem;
    padding-bottom: 1rem;
}

/* TÍTULO PRINCIPAL */
h1, h2, h3 {
    color: white;
}

/* HERO */
.hero {
    text-align: center;
    margin-bottom: 1.5rem;
}

.hero h1 {
    font-size: 1.8rem;
    font-weight: 600;
}

.hero p {
    color: var(--secondary-text);
    font-size: 1.1rem;
}

/* BOTÕES */
.stButton>button {
    background: linear-gradient(90deg, var(--accent-start), var(--accent-end));
    color: white;
    border-radius: 8px;
    height: 2.6em;
    border: none;
    font-weight: 500;
    transition: transform 0.1s ease, box-shadow 0.1s ease;
}

.stButton>button:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(255, 100, 130, 0.3);
}

.stButton>button:disabled {
    background: linear-gradient(90deg, #88334d, #99445f);
    opacity: 0.7;
    cursor: not-allowed;
    transform: none;
    box-shadow: none;
}

/* BOTÃO DE DOWNLOAD */
.stDownloadButton>button {
    background-color: var(--success) !important;
    border-radius: 8px;
    height: 2.6em;
    border: none;
    font-weight: 500;
    transition: transform 0.1s ease, box-shadow 0.1s ease;
}

.stDownloadButton>button:hover {
    transform: translateY(-1px);
    box-shadow: 0 4px 12px rgba(36, 161, 72, 0.3);
}

/* TABELA DE HISTÓRICO */
.stDataFrame {
    border-radius: 8px;
    overflow: hidden;
}

/* ESPAÇAMENTO ENTRE SESSÕES */
.section-title {
    margin-top: 1rem;
    margin-bottom: 0.8rem;
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

            status.markdown(f"**🔄 Página {page_num+1} de {total_paginas}**")
            progress_bar.progress((page_num + 1) / total_paginas)

            texto = page.extract_text() or page.extract_text(layout=True)

            if not texto:
                continue

            linhas = texto.split("\\n")

            nome_atual = None
            matricula_atual = None

            for linha in linhas:

                linha = linha.strip()
                linha = re.sub(r"\\s{2,}", " ", linha)

                mat_match = re.search(r"MAT\\.?\\s*:?\\s*(\\d{5,7})", linha)
                if mat_match:
                    matricula_atual = mat_match.group(1)

                nome_match = re.search(
                    r"NOME\\s*:?\\s*([A-ZÀ-Ú\\s]+?)(?:FUNCAO|FUNC|DT|$)",
                    linha
                )
                if nome_match:
                    nome_atual = nome_match.group(1).strip()

                if "|" in linha and re.search(r"\\d{3}", linha):

                    partes = linha.split("|")

                    for idx, parte in enumerate(partes):

                        parte = parte.strip()
                        if not parte:
                            continue

                        evento_match = re.match(
                            r"(\\d{3})\\s+(.+?)\\s+([\\d,]+)?\\s+([\\d.,]+)",
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
                "total": round(vlr_med + vlr_odonto, 2)
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
# LAYOUT 2 COLUNAS: Processamento | Histórico 💖
# -------------------------------
col_process, col_hist = st.columns([2, 1], gap="medium")

with col_process:

    st.markdown('<div class="section-title"><h2>💌 Processamento dos PDFs</h2></div>', unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Arraste seus PDFs aqui",
        type=["pdf"],
        accept_multiple_files=True,
        label_visibility="collapsed"
    )

    if uploaded_files:
        for file in uploaded_files:
            aux_cols = st.columns([4, 1, 1])

            if file.name in st.session_state.arquivos_processados:
                aux_cols[0].markdown(f"✅ {file.name}")
                aux_cols[1].button("Processar", key=file.name, disabled=True)
                aux_cols[2].download_button(
                    "Baixar",
                    st.session_state.arquivos_processados[file.name],
                    file_name=file.name.replace(".pdf", ".xlsx"),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                aux_cols[0].markdown(f"📎 {file.name}")
                if aux_cols[1].button("Processar", key=file.name):
                    status_p = st.empty()
                    status_p.info("🔄 Iniciando processamento... (preta, tenha um pouco de paciência 😂♥️)")

                    resultado, df_totvs = processar_pdf(file)

                    status_p.success("Prontinho, meu amor 💚")

                    st.session_state.arquivos_processados[file.name] = resultado


with col_hist:

    st.markdown('<div class="section-title"><h2>📋 Histórico de Processamentos</h2></div>', unsafe_allow_html=True)

    if st.session_state.historico:
        df_hist = pd.DataFrame(st.session_state.historico)
        st.dataframe(df_hist, use_container_width=True)
    else:
        st.caption("Ainda não há processamentos.")
