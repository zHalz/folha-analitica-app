import streamlit as st
import pdfplumber
import pandas as pd
import re
import tempfile
from datetime import datetime
from openpyxl import load_workbook

st.set_page_config(page_title="Folha Analítica", layout="centered")

# -------------------------------
# ESTILO
# -------------------------------
st.markdown("""
    <style>
        .stButton>button {
            background-color: #0f62fe;
            color: white;
            border-radius: 8px;
            height: 3em;
            width: 100%;
            font-size: 16px;
        }
        .stDownloadButton>button {
            background-color: #24a148;
            color: white;
            border-radius: 8px;
            height: 3em;
            width: 100%;
            font-size: 16px;
        }
    </style>
""", unsafe_allow_html=True)

# -------------------------------
# TÍTULO
# -------------------------------
st.title("📄 Processador de Folha Analítica pra minha preta (Karem 💍♥️)")
st.caption("Mor, envie o PDF e receba o Excel pronto")

# -------------------------------
# HISTÓRICO
# -------------------------------
if "historico" not in st.session_state:
    st.session_state.historico = []

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

            status.text(f"📄 Página {page_num+1}/{total_paginas}")
            progresso.progress((page_num + 1) / total_paginas)

            texto = page.extract_text() or ""

            if not texto:
                continue

            linhas = texto.split("\n")

            nome_atual = None
            matricula_atual = None

            for linha in linhas:

                linha = re.sub(r"\s{2,}", " ", linha.strip())

                mat = re.search(r"MAT\.?\s*:?\s*(\d+)", linha)
                if mat:
                    matricula_atual = mat.group(1)

                nome = re.search(r"NOME\s*:?\s*([A-ZÀ-Ú\s]+)", linha)
                if nome:
                    nome_atual = nome.group(1).strip()

                if "|" in linha:

                    partes = linha.split("|")

                    for idx, parte in enumerate(partes):

                        match = re.match(r"(\d{3})\s+(.+?)\s+([\d.,]+)", parte.strip())

                        if match:
                            cod, desc, valor = match.groups()

                            valor = float(valor.replace(".", "").replace(",", "."))

                            dados.append({
                                "nome": nome_atual,
                                "matricula": matricula_atual,
                                "tipo": "PROVENTO" if idx == 0 else "DESCONTO",
                                "codigo": cod,
                                "descricao": desc,
                                "valor": valor
                            })

    return pd.DataFrame(dados)

# -------------------------------
# UPLOAD
# -------------------------------
uploaded_file = st.file_uploader("📤 Envie o PDF", type=["pdf"])

if uploaded_file:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        pdf_path = tmp.name

    df = extrair_folha_analitica(pdf_path)

    if df.empty:
        st.error("❌ Nenhum dado encontrado")
        st.stop()

    pivot = df.pivot_table(
        values="valor",
        index=["nome", "matricula", "tipo"],
        columns="codigo",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    mapa = {
        "455": "Med_Tit",
        "454": "Odo_Tit",
        "458": "Cop",
        "456": "Med_Dep",
        "461": "Odo_Dep"
    }

    for col in mapa:
        if col not in pivot.columns:
            pivot[col] = 0

    analise = pivot[["nome","matricula"]+list(mapa.keys())].rename(columns=mapa)
    analise = analise.groupby(["nome","matricula"]).sum().reset_index()

    linhas = []

    for _, r in analise.iterrows():

        linhas.append({
            "nome": r["nome"],
            "matricula": r["matricula"],
            "tipo_registro": "TITULAR",
            "dependente_id": 0,
            "vlr_medico": r["Med_Tit"],
            "vlr_odonto": r["Odo_Tit"],
            "vlr_copart": r["Cop"],
            "total": r["Med_Tit"] + r["Odo_Tit"] + r["Cop"]
        })

    df_totvs = pd.DataFrame(linhas)

    # -------------------------------
    # EXCEL
    # -------------------------------
    output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

    with pd.ExcelWriter(output.name, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Detalhamento", index=False)
        pivot.to_excel(writer, sheet_name="Pivot", index=False)
        analise.to_excel(writer, sheet_name="Analise", index=False)
        df_totvs.to_excel(writer, sheet_name="TOTVS", index=False)

    # -------------------------------
    # PÓS PROCESSAMENTO
    # -------------------------------
    wb = load_workbook(output.name)
    ws = wb["TOTVS"]

    ultima_linha = max(cell.row for cell in ws["C"] if cell.value)

    lt = ultima_linha + 2
    ld = ultima_linha + 3

    ws[f"D{lt}"] = "TITULAR"
    ws[f"D{ld}"] = "DEPENDENTE"

    range_e = f"E2:E{ultima_linha}"
    range_f = f"F2:F{ultima_linha}"
    range_g = f"G2:G{ultima_linha}"
    range_h = f"H2:H{ultima_linha}"
    range_c = f"C2:C{ultima_linha}"

    # TITULAR
    ws[f"E{lt}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(E2,ROW({range_e})-ROW(E2),0)),({range_c}=D{lt})+0)'
    ws[f"F{lt}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(F2,ROW({range_f})-ROW(F2),0)),({range_c}=D{lt})+0)'
    ws[f"G{lt}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(G2,ROW({range_g})-ROW(G2),0)),({range_c}=D{lt})+0)'
    ws[f"H{lt}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(H2,ROW({range_h})-ROW(H2),0)),({range_c}=D{lt})+0)'

    # DEPENDENTE
    ws[f"E{ld}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(E2,ROW({range_e})-ROW(E2),0)),({range_c}=D{ld})+0)'
    ws[f"F{ld}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(F2,ROW({range_f})-ROW(F2),0)),({range_c}=D{ld})+0)'
    ws[f"G{ld}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(G2,ROW({range_g})-ROW(G2),0)),({range_c}=D{ld})+0)'
    ws[f"H{ld}"] = f'=SUMPRODUCT(SUBTOTAL(9,OFFSET(H2,ROW({range_h})-ROW(H2),0)),({range_c}=D{ld})+0)'

    wb.calculation.fullCalcOnLoad = True
    wb.save(output.name)

    # -------------------------------
    # DOWNLOAD
    # -------------------------------
    with open(output.name, "rb") as f:
        st.success("✅ Pronto!")
        st.download_button("⬇️ Baixar Excel", f, file_name="folha.xlsx")
