import os
import pandas as pd
import streamlit as st
import io
import time
from docx import Document

st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
st.title("📄 Emissor de Memorandos Individuais - Hospital Dr. Jackson Lago")

arquivo_excel = st.file_uploader("Suba a planilha contendo os incidentes (.xlsx)", type=["xlsx"])
caminho_modelo = "modelo_memorando.docx"

def substituir_texto_protegendo_logos(doc, dicionario_tags):
    """Substitui o texto alterando apenas os 'runs' para proteger imagens e cabeçalhos."""
    for paragrafo in doc.paragraphs:
        for tag, valor in dicionario_tags.items():
            if tag in paragrafo.text:
                for run in paragrafo.runs:
                    if tag in run.text:
                        run.text = run.text.replace(tag, valor)
                        
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    for tag, valor in dicionario_tags.items():
                        if tag in paragrafo.text:
                            for run in paragrafo.runs:
                                if tag in run.text:
                                    run.text = run.text.replace(tag, valor)

def formatar_data_br(valor):
    """Garante que as datas sejam exibidas no formato brasileiro DD/MM/AAAA"""
    try:
        if pd.isna(valor) or str(valor).strip() == "":
            return ""
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except:
        v_str = str(valor).strip()
        # Se vier com hora do Excel (Ex: 2026-06-25 00:00:00), limpa e mantém só a data
        if " " in v_str:
            v_str = v_str.split(" ")[0]
        if "-" in v_str:
            partes = v_str.split("-")
            if len(partes) == 3:
                return f"{partes[2]}/{partes[1]}/{partes[0]}"
        return v_str

def limpar_nan(valor):
    """Substitui qualquer valor vazio ou 'nan' por uma linha em branco limpa"""
    if pd.isna(valor) or str(valor).strip().lower() == "nan" or str(valor).strip() == "":
        return ""
    return str(valor).strip()

def limpar_numero_float(valor):
    """Remove o .0 de números inteiros vindos do Excel (ex: 886.0 vira 886)"""
    if pd.isna(valor):
        return ""
    try:
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        v_str = str(valor).strip()
        if v_str.endswith(".0"):
            return v_str[:-2]
        return v_str.strip()
    except:
        return str(valor).strip()

@st.fragment
def renderizar_linha_paciente_sob_demanda(index, linha, col_paciente, mapa_colunas):
    nome_do_paciente = str(linha[col_paciente]).strip()
    
    # Coleta inteligente do número do memorando nas duas colunas possíveis
    memo_01 = str(linha.get(mapa_colunas.get("MEMO01", ""), "")).strip()
    memo_02 = str(linha.get(mapa_colunas.get("MEMO02", ""), "")).strip()
    
    if memo_01 == "" or memo_01.lower() == "nan":
        num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" else "S-N"
    else:
        num_memo_cru = memo_01
        
    num_notif = limpar_numero_float(linha.get(mapa_colunas.get("NOTIF", ""), index + 1))
    if num_notif == "":
        num_notif = str(index + 1)
    
    # Padronização limpa do nome do arquivo final
    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
    nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
    
    # Lógica minuciosa para os turnos (Busca em maiúsculo na coluna certa)
    turno_planilha = str(linha.get(mapa_colunas.get("TURNO", ""), "")).strip().upper()
    marca_manha = "X" if "MANH" in turno_planilha else " "
    marca_tarde = "X" if "TARD" in turno_planilha else " "
    marca_noite = "X" if "NOIT" in turno_planilha else " "
    
    # Dicionário de tags aplicando o mapa inteligente de colunas
    dados_memorando = {
        "{{numero_memorando}}": num_memo_cru,
        "{{gestor}}": limpar_nan(linha.get(mapa_colunas.get("GESTOR", ""), "")),
        "{{setor}}": limpar_nan(linha.get(mapa_colunas.get("SETOR_NOTIF", ""), "")),
        "{{notificacao_n}}": num_notif,
        "{{data_notificacao}}": formatar_data_br(linha.get(mapa_colunas.get("DATA_NOTIF", ""), "")),
        "{{data_ocorrencia}}": formatar_data_br(linha.get(mapa_colunas.get("DATA_OCORR", ""), "")),
        "{{localizacao}}": limpar_nan(linha.get(mapa_colunas.get("ONDE", ""), "")), 
        "{{classificacao_incidente}}": limpar_nan(linha.get("CLASSIFICAÇÃO DO INCIDENTE", linha.get("CLASSIFICACAO DO INCIDENTE", ""))), 
        "{{setor_notificante}}": limpar_nan(linha.get("SETOR NOTIFICANTE", "")), 
        "{{tipo_incidente}}": limpar_nan(linha.get(mapa_colunas.get("TIPO", ""), "")),
        "{{descricao_notificacao}}": limpar_nan(linha.get(mapa_colunas.get("DESC", ""), "")),
        "{{nome_paciente}}": nome_do_paciente,
        "{{leito}}": limpar_numero_float(linha.get(mapa_colunas.get("LEITO", ""), "")),
        "{{sugestao}}": limpar_nan(linha.get(mapa_colunas.get("SUGESTAO", ""), "")),
        "{{m}}": marca_manha,
        "{{t}}": marca_tarde,
        "{{n}}": marca_noite
    }

    col_nome, col_word, col_pdf = st.columns(3)
    
    with col_nome:
        st.markdown(f"**🔹 {nome_do_paciente}**")
        
    with col_word:
        doc_word = Document(caminho_modelo)
        substituir_texto_protegendo_logos(doc_word, dados_memorando)
        word_io = io.BytesIO()
        doc_word.save(word_io)
        word_io.seek(0)
        
        st.download_button(
            label="📝 Baixar WORD",
            data=word_io.getvalue(),
            file_name=f"{nome_base_arquivo}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"w_{index}"
        )
        
    with col_pdf:
        doc_pdf = Document(caminho_modelo)
        substituir_texto_protegendo_logos(doc_pdf, dados_memorando)
        pdf_io = io.BytesIO()
        doc_pdf.save(pdf_io)
        pdf_io.seek(0)
        
        st.download_button(
            label="📕 Baixar PDF",
            data=pdf_io.getvalue(),
            file_name=f"{nome_base_arquivo}.pdf",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"p_{index}"
        )
    st.markdown("---")

if arquivo_excel:
    df = pd.read_excel(arquivo_excel)
    
    # 🔍 Dicionário de Tradução Dinâmica para ignorar maiúsculas/minúsculas e acentos
    mapa_colunas = {}
    for col in df.columns:
        c_limpa = str(col).strip().upper().replace("Á", "A").replace("ÇÃO", "CAO").replace("Õ", "O").replace("Í", "I")
        
        if col == "Nº": mapa_colunas["NOTIF"] = col
        elif "PACIENTE" in c_limpa or "NOME" in c_limpa: mapa_colunas["PACIENTE"] = col
        elif "MEMO 01" in c_limpa: mapa_colunas["MEMO01"] = col
        elif "MEMO 02" in c_limpa: mapa_colunas["MEMO02"] = col
        elif "GESTOR" in c_limpa: mapa_colunas["GESTOR"] = col
        elif "SETOR NOTIFICADO" in c_limpa: mapa_colunas["SETOR_NOTIF"] = col
        elif "DATA" in c_limpa and "NOTIF" in c_limpa: mapa_colunas["DATA_NOTIF"] = col
        elif "DATA" in c_limpa and "OCORR" in c_limpa: mapa_colunas["DATA_OCORR"] = col
        elif "TURNO" in c_limpa: mapa_colunas["TURNO"] = col
        elif "TIPO" in c_limpa: mapa_colunas["TIPO"] = col
        elif "DESC" in c_limpa or "RESUMO" in c_limpa: mapa_colunas["DESC"] = col
        elif "LEITO" in c_limpa: mapa_colunas["LEITO"] = col
        elif "SUGEST" in c_limpa: mapa_colunas["SUGESTAO"] = col
        elif "ONDE OCORREU" in c_limpa: mapa_colunas["ONDE"] = col

    coluna_paciente = mapa_colunas.get("PACIENTE", df.columns[1])
        
    df = df.dropna(subset=[coluna_paciente])
    df = df[df[coluna_paciente].astype(str).str.strip() != ""]
    
    st.success(f"📋 Lista de verificação pronta! {len(df)} registros encontrados.")
    
    for index, linha in df.iterrows():
        renderizar_linha_paciente_sob_demanda(index, linha, coluna_paciente, mapa_colunas)
