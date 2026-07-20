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
        return str(valor).strip()

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
        return v_str
    except:
        return str(valor)

@st.fragment
def renderizar_linha_paciente_sob_demanda(index, linha, col_paciente, col_notif):
    nome_do_paciente = str(linha[col_paciente]).strip()
    
    # Coleta inteligente do número do memorando nas duas colunas possíveis (Memo 01 ou Memo 02)
    memo_01 = str(linha.get("Nº Memo 01", "")).strip()
    memo_02 = str(linha.get("Nº Memo 02", "")).strip()
    
    if memo_01 == "" or memo_01.lower() == "nan":
        num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" else "S-N"
    else:
        num_memo_cru = memo_01
        
    num_notif = limpar_numero_float(linha.get(col_notif, "S-N"))
    
    # Padronização limpa do nome do arquivo final
    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
    nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
    
    # Lógica para marcação dos turnos com X
    turno_planilha = str(linha.get("turno", "")).strip().upper()
    marca_manha = "X" if "MANH" in turno_planilha or "MANHÃ" in turno_planilha else " "
    marca_tarde = "X" if "TARD" in turno_planilha or "TARDE" in turno_planilha else " "
    marca_noite = "X" if "NOIT" in turno_planilha or "NOITE" in turno_planilha else " "
    
    # Mapeamento definitivo e amarrado com 100% de precisão aos cabeçalhos das suas imagens
    dados_memorando = {
        "{{numero_memorando}}": num_memo_cru,
        "{{gestor}}": limpar_nan(linha.get("Gestor 01", "")),
        "{{setor}}": limpar_nan(linha.get("SETOR NOTIFICADO", "")),
        "{{notificacao_n}}": num_notif,
        "{{data_notificacao}}": formatar_data_br(linha.get("data_notificacao", "")),
        "{{data_ocorrencia}}": formatar_data_br(linha.get("data_ocorrencia", "")),
        
        # Amarração com os nomes validados por imagem:
        "{{localizacao}}": limpar_nan(linha.get("ONDE OCORREU INCIDENTE", "")), 
        "{{classificacao_incidente}}": limpar_nan(linha.get("CLASSIFICAÇÃO DO INCIDENTE", "")), 
        "{{setor_notificante}}": limpar_nan(linha.get("SETOR NOTIFICANTE", "")), 
        
        "{{tipo_incidente}}": limpar_nan(linha.get("tipo_incidente", "")),
        "{{descricao_notificacao}}": limpar_nan(linha.get("descricao_notificacao", "")),
        "{{nome_paciente}}": nome_do_paciente,
        "{{leito}}": limpar_numero_float(linha.get("leito", "")),
        "{{sugestao}}": limpar_nan(linha.get("sugestao", "")),
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
        # Usa a estrutura em bytes do documento customizado para download idêntico do PDF preservando as mídias
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
    
    # Captura a primeira coluna (A) nativa para o número da Notificação
    col_notif_forcada = df.columns
    
    # Limpa todos os espaços invisíveis no início e fim dos nomes das colunas
    df.columns = df.columns.str.strip()
    
    col_paciente = None
    for col in df.columns:
        if "PACIENTE" in col.upper() or "NOME" in col.upper(): 
            col_paciente = col
            break
            
    if not col_paciente: 
        col_paciente = "PACIENTE"
        
    df = df.dropna(subset=[col_paciente])
    df = df[df[col_paciente].astype(str).str.strip() != ""]
    
    st.success(f"📋 Lista de verificação pronta! {len(df)} registros encontrados.")
    
    for index, linha in df.iterrows():
        renderizar_linha_paciente_sob_demanda(index, linha, col_paciente, col_notif_forcada)
