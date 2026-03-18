import streamlit as st
import pdfplumber
import pandas as pd
import re
from tqdm import tqdm
import tempfile

st.set_page_config(page_title="Folha Analítica", layout="centered")

st.title("📄 Processador de Folha Analítica")
st.write("Envie o PDF e baixe o Excel pronto")

# -------------------------------
# FUNÇÃO EXTRAÇÃO (SEU CÓDIGO)
# -------------------------------

def extrair_folha_analitica(pdf_path):

    dados = []

    with pdfplumber.open(pdf_path) as pdf:

        for page_num, page in enumerate(pdf.pages):

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
                                "linha_original": i + 1,
                                "nome": nome_final,
                                "matricula": matricula_final,
                                "tipo": tipo,
                                "codigo": codigo,
                                "descricao": desc.strip(),
                                "referencia": ref,
                                "valor": valor
                            })

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

    st.info("🔄 Processando arquivo...")

    df_consolidado = extrair_folha_analitica(pdf_path)

    if df_consolidado.empty:
        st.error("❌ Nenhum dado encontrado no PDF")
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
    ).round(2).reset_index()

    # -------------------------------
    # ANÁLISE
    # -------------------------------

    mapa_codigos = {
        "455": "Assistência Médica Titular",
        "454": "Assistência Odontológica Titular",
        "458": "Coparticipação",
        "456": "Assistência Médica Dependente",
        "461": "Assistência Odontológica Dependente"
    }

    for col in mapa_codigos.keys():
        if col not in pivot_completa.columns:
            pivot_completa[col] = 0

    analise_plano = pivot_completa[
        ["nome", "matricula"] + list(mapa_codigos.keys())
    ].copy().rename(columns=mapa_codigos)

    analise_plano = analise_plano.groupby(
        ["nome", "matricula"]
    ).sum().reset_index()

    # -------------------------------
    # SPLIT CORRETO
    # -------------------------------

    linhas = []

    for _, row in analise_plano.iterrows():

        nome = row["nome"]
        mat = row["matricula"]

        med_tit = row["Assistência Médica Titular"]
        odo_tit = row["Assistência Odontológica Titular"]

        med_dep_total = row["Assistência Médica Dependente"]
        odo_dep_total = row["Assistência Odontológica Dependente"]

        cop = row["Coparticipação"]

        qtd_dep_med = int(round(med_dep_total / med_tit)) if med_tit > 0 else 0
        qtd_dep_odo = int(round(odo_dep_total / odo_tit)) if odo_tit > 0 else 0

        qtd_dependentes = max(qtd_dep_med, qtd_dep_odo)

        # TITULAR
        linhas.append({
            "nome": nome,
            "matricula": mat,
            "tipo_registro": "TITULAR",
            "dependente_id": 0,
            "vlr_medico": med_tit,
            "vlr_odonto": odo_tit,
            "vlr_copart": cop,
            "total": round(med_tit + odo_tit + cop, 2)
        })

        # DEPENDENTES
        for i in range(qtd_dependentes):

            vlr_med = med_dep_total / qtd_dep_med if i < qtd_dep_med and qtd_dep_med > 0 else 0
            vlr_odo = odo_dep_total / qtd_dep_odo if i < qtd_dep_odo and qtd_dep_odo > 0 else 0

            linhas.append({
                "nome": nome,
                "matricula": mat,
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
        analise_plano.to_excel(writer, sheet_name="Analise", index=False)
        df_totvs.to_excel(writer, sheet_name="TOTVS", index=False)

    # -------------------------------
    # DOWNLOAD
    # -------------------------------

    with open(output.name, "rb") as f:
        st.success("✅ Processamento concluído!")
        st.download_button(
            "⬇️ Baixar Excel",
            f,
            file_name="folha_processada.xlsx"
        )
