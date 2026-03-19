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
st.title("📄 Processador de Folha Analítica pra Minha Preta (Karem 💍♥️)")
st.caption("Envie o PDF e receba o Excel pronto para o TOTVS")

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

            status.text(f"📄 Processando página {page_num+1} de {total_paginas}")
            progresso.progress((page_num + 1) / total_paginas)

            texto = page.extract_text() or page.extract_text(layout=True) or page.extract_text(x_tolerance=3)

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

        status.text("✅ Extração concluída!")

    df = pd.DataFrame(dados)

    if not df.empty:
        df["nome"] = df["nome"].str.replace(r"[:\s]+$", "", regex=True).str.strip()

    return df


# -------------------------------
# UPLOAD
# -------------------------------
uploaded_file = st.file_uploader("📤 Envie o PDF", type=["pdf"])

if uploaded_file:

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        pdf_path = tmp.name

    st.info("🔄 Iniciando processamento... (preta, tenha um pouco de paciência 😂♥️)")

    df = extrair_folha_analitica(pdf_path)

    if df.empty:
        st.error("❌ Nenhum dado encontrado")
        st.stop()

    # -------------------------------
    # PIVOT
    # -------------------------------
    pivot = df.pivot_table(
        values="valor",
        index=["nome", "matricula", "tipo"],
        columns="codigo",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # -------------------------------
    # ANÁLISE
    # -------------------------------
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

    # -------------------------------
    # SPLIT
    # -------------------------------
    linhas = []

    for _, row in analise.iterrows():

        med_tit = row["Assistência Médica Titular"]
        odo_tit = row["Assistência Odontológica Titular"]

        med_dep = row["Assistência Médica Dependente"]
        odo_dep = row["Assistência Odontológica Dependente"]

        qtd_med = int(round(med_dep / med_tit)) if med_tit > 0 else 0
        qtd_odo = int(round(odo_dep / odo_tit)) if odo_tit > 0 else 0

        qtd = max(qtd_med, qtd_odo)

        # TITULAR
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

        # DEPENDENTES
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

    # -------------------------------
    # EXCEL EM MEMÓRIA
    # -------------------------------
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Detalhamento", index=False)
        pivot.to_excel(writer, sheet_name="Pivot", index=False)
        analise.to_excel(writer, sheet_name="Analise", index=False)
        df_totvs.to_excel(writer, sheet_name="Base_TOTVS", index=False)

    buffer.seek(0)

    # -------------------------------
    # PÓS-PROCESSAMENTO TOTVS
    # -------------------------------
    wb = load_workbook(buffer)
    ws = wb["Base_TOTVS"]

    ultima = ws.max_row

    linha_tit = ultima + 2
    linha_dep = ultima + 3

    ws[f"D{linha_tit}"] = "TITULAR"
    ws[f"D{linha_dep}"] = "DEPENDENTE"

    for col in ["E", "F", "G", "H"]:
        ws[f"{col}{linha_tit}"] = f'=SUMIF(C:C,"TITULAR",{col}:{col})'
        ws[f"{col}{linha_dep}"] = f'=SUMIF(C:C,"DEPENDENTE",{col}:{col})'

    final = BytesIO()
    wb.save(final)
    final.seek(0)

    # -------------------------------
    # HISTÓRICO
    # -------------------------------
    st.session_state.historico.append({
        "arquivo": uploaded_file.name,
        "data": datetime.now().strftime("%d/%m %H:%M"),
        "linhas": len(df_totvs)
    })

    # -------------------------------
    # DOWNLOAD
    # -------------------------------
    st.success("✅ Pronto!")

    st.download_button(
        "⬇️ Baixar Excel",
        final,
        file_name="folha_processada.xlsx"
    )

# -------------------------------
# HISTÓRICO VISUAL
# -------------------------------
if st.session_state.historico:
    st.divider()
    st.subheader("📊 Histórico")
    st.dataframe(pd.DataFrame(st.session_state.historico), use_container_width=True)
