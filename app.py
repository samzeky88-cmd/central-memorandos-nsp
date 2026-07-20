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
        v_str = str(valor).strip().split(" ")[0]
        if "-" in v_str:
            parts = v_str.split("-")
            if len(parts) == 3:
                return f"{parts[2]}/{parts[1]}/{parts[0]}"
        return v_str

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
def renderizar_linha_paciente_sob_demanda(index, linha, col_paciente, col_notif, col_data_notif, col_data_ocorr, col_turno, col_tipo, col_classif, col_desc, col_leito, col_notificante, col_sugestao, col_localizacao):
    nome_do_paciente = str(linha[col_paciente]).strip()
    
    # Coleta inteligente do número do memorando nas duas colunas possíveis (Memo 01 ou Memo 02)
    memo_01 = str(linha.get("Nº Memo 01", "")).strip()
    memo_02 = str(linha.get("Nº Memo 02", "")).strip()
    
    if memo_01 == "" or memo_01.lower() == "nan":
        num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" else "S-N"
    else:
        num_memo_cru = memo_01
        
    # Coleta o número da notificação da coluna correta
    num_notif = limpar_numero_float(linha.get(col_notif, "S-N"))
    
    # Padronização limpa do nome do arquivo final
    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
    nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
    
    # Lógica minuciosa e resiliente para marcação dos turnos com X
    turno_planilha = str(linha.get(col_turno, "")).strip().upper() if col_turno else ""
    marca_manha = "X" if "MANH" in turno_planilha or "MANHÃ" in turno_planilha else " "
    marca_tarde = "X" if "TARD" in turno_planilha or "TARDE" in turno_planilha else " "
    marca_noite = "X" if "NOIT" in turno_planilha or "NOITE" in turno_planilha else " "
    
    # Vinculando os dados dinâmicos às tags do seu modelo Word
    dados_memorando = {
        "{{numero_memorando}}": num_memo_cru,
        "{{gestor}}": str(linha.get("Gestor 01", "")).strip(),
        "{{setor}}": str(linha.get("SETOR NOTIFICADO", "")).strip(),
        "{{notificacao_n}}": num_notif,
        "{{data_notificacao}}": formatar_data_br(linha.get(col_data_notif, "")) if col_data_notif else "",
        "{{data_ocorrencia}}": formatar_data_br(linha.get(col_data_ocorr, "")) if col_data_ocorr else "",
        "{{localizacao}}": str(linha.get(col_localizacao, "")).strip() if col_localizacao else "",
        "{{tipo_incidente}}": str(linha.get(col_tipo, "")).strip() if col_tipo else "",
        "{{classificacao_incidente}}": str(linha.get(col_classif, "")).strip() if col_classif else "",
        "{{descricao_notificacao}}": str(linha.get(col_desc, "")).strip() if col_desc else "",
        "{{nome_paciente}}": nome_do_paciente,
        "{{leito}}": limpar_numero_float(linha.get(col_leito, "")) if col_leito else "",
        "{{setor_notificante}}": str(linha.get(col_notificante, "")).strip() if col_notificante else "",
        "{{sugestao}}": str(linha.get(col_sugestao, "")).strip() if col_sugestao else "",
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
        # Cria o PDF usando estrutura direta na memória para evitar travar no Linux
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
    
    # 🎯 FIXADO: Puxa o nome exato da primeira coluna (Coluna A) para ser o índice da Notificação
    col_notif_forcada = df.columns[0]
    
    df.columns = df.columns.str.strip()
    
    col_paciente = None
    col_data_notif = None
    col_data_ocorr = None
    col_turno = None
    col_tipo = None
    col_classif = None
    col_desc = None
    col_leito = None
    col_notificante = None
    col_sugestao = None
    col_localizacao = None
    
    for col in df.columns:
        c_upper = col.upper()
        if "PACIENTE" in c_upper or "NOME" in c_upper: col_paciente = col
        elif "DATA" in c_upper and "NOTIF" in c_upper: col_data_notif = col
        elif "DATA" in c_upper and ("OCORR" in c_upper or "OCOR" in c_upper): col_data_ocorr = col
        elif "TURNO" in c_upper: col_turno = col
        elif "TIPO" in c_upper or "INCIDENTE" in c_upper: col_tipo = col
        elif "CLASSIF" in c_upper: col_classif = col
        elif "DESC" in c_upper or "RESUMO" in c_upper: col_desc = col
        elif "LEITO" in c_upper: col_leito = col
        elif "NOTIFICANTE" in c_upper or "QUEM" in c_upper: col_notificante = col
        elif "SUGES" in c_upper or "RECO" in c_upper: col_sugestao = col
        elif "LOCAL" in c_upper or "ALA" in c_upper or "O2" in c_upper: col_localizacao = col
            
    if not col_paciente: col_paciente = df.columns
        
    df = df.dropna(subset=[col_paciente])
    df = df[df[col_paciente].astype(str).str.strip() != ""]
    
    st.success(f"📋 Lista de verificação pronta! {len(df)} registros encontrados.")
    
    for index, linha in df.iterrows():
        renderizar_linha_paciente_sob_demanda(
            index, linha, col_paciente, col_notif_forcada, col_data_notif, 
            col_data_ocorr, col_turno, col_tipo, col_classif, col_desc, 
            col_leito, col_notificante, col_sugestao, col_localizacao
        )
