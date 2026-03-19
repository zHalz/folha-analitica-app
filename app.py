import streamlit as st
import pdfplumber
import pandas as pd
import re
import tempfile
from datetime import datetime

st.set_page_config(page_title="Folha Analítica", layout="centered")

# -------------------------------
# ESTILO (deixa bonito)
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
st.caption("Mor, envie o PDF e receba o Excel pronto para importação")

# -------------------------------
# HISTÓRICO (sessão)
# -------------------------------
if "historico" not in st.session_state:
    st.session_state.historico = []

# -------------------------------
# FUNÇÃO EXTRAÇÃO COM PROGRESSO
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

            for i, linha in enumerate(linhas):

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

                    nome_final = nome_atual if nome_atual else "SEM_NOME"
                    matricula_final = matricula_atual if matricula_atual else "SEM_MAT"

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
                                "nome": nome_final,
                                "matricula": matricula_final,
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

    nome_arquivo = uploaded_file.name

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(uploaded_file.read())
        pdf_path = tmp.name

    st.info("🔄 Iniciando processamento... (preta, tenha um pouco de paciência 😂♥️)")

    df_consolidado = extrair_folha_analitica(pdf_path)

    if df_consolidado.empty:
        st.error("❌ Nenhum dado encontrado")
        st.stop()

    # -------------------------------
    # PIVOT
    # -------------------------------
    pivot_completa = df_consolidado.pivot_table(
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
        if col not in pivot_completa.columns:
            pivot_completa[col] = 0

    analise = pivot_completa[
        ["nome", "matricula"] + list(mapa.keys())
    ].rename(columns=mapa)

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
    # GERAR EXCEL
    # -------------------------------
    output = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")

    with pd.ExcelWriter(output.name, engine="openpyxl") as writer:
        df_consolidado.to_excel(writer, sheet_name="Detalhamento", index=False)
        pivot_completa.to_excel(writer, sheet_name="Pivot", index=False)
        analise.to_excel(writer, sheet_name="Analise", index=False)
        df_totvs.to_excel(writer, sheet_name="TOTVS", index=False)

    # -------------------------------
    # HISTÓRICO
    # -------------------------------
    st.session_state.historico.append({
        "arquivo": nome_arquivo,
        "data": datetime.now().strftime("%d/%m %H:%M"),
        "linhas": len(df_totvs)
    })

    # -------------------------------
    # DOWNLOAD
    # -------------------------------
    with open(output.name, "rb") as f:
        st.success("✅ Concluído!")
        st.download_button("⬇️ Baixar Excel", f, file_name="folha_processada.xlsx")

# -------------------------------
# EXIBIR HISTÓRICO
# -------------------------------
if st.session_state.historico:

    st.divider()
    st.subheader("📊 Histórico de Processamentos")

    hist_df = pd.DataFrame(st.session_state.historico)
    st.dataframe(hist_df, use_container_width=True)
