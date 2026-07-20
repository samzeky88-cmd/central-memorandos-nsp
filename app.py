import os
import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document

st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
st.title("📄 Emissor de Memorandos Individuais - Hospital Dr. Jackson Lago")

st.markdown("### 🗓️ Configuração da Data de Envio")
data_selecionada = st.date_input("Selecione a data que sairá no cabeçalho do Memorando:", datetime.now())

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
        return v_str.strip()
    except:
        return str(valor).strip()

def obter_data_por_extenso(dt):
    """Gera a data selecionada por extenso em português brasileiro (Ex: 20 de Julho de 2026)"""
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return f"{dt.day} de {meses[dt.month]} de {dt.year}"

if arquivo_excel:
    nome_chave_cache = f"dados_memos_{arquivo_excel.name}_{data_selecionada.strftime('%Y%m%d')}"
    
    if nome_chave_cache not in st.session_state:
        with st.spinner("📦 Processando e estruturando os memorandos na memória... Aguarde um instante."):
            df = pd.read_excel(arquivo_excel)
            
            # Varredura das colunas exatas da planilha
            mapa_colunas = {}
            for col in df.columns:
                col_limpa = str(col).strip()
                c_upper = col_limpa.upper()
                
                if col_limpa == "Nº": 
                    mapa_colunas["NOTIF"] = col
                elif col_limpa == "Nº Memo 01":
                    mapa_colunas["MEMO01"] = col
                elif col_limpa == "Nº Memo 02":
                    mapa_colunas["MEMO02"] = col
                elif "PACIENTE" in c_upper or "NOME" in c_upper: 
                    mapa_colunas["PACIENTE"] = col
                elif "DATA" in c_upper and "NOTIF" in c_upper: 
                    mapa_colunas["DATA_NOTIF"] = col
                elif "DATA" in c_upper and "OCORR" in c_upper: 
                    mapa_colunas["DATA_OCORR"] = col
                elif "TURNO" in c_upper: 
                    mapa_colunas["TURNO"] = col
                elif "TIPO" in c_upper: 
                    mapa_colunas["TIPO"] = col
                elif "DESC" in c_upper or "RESUMO" in c_upper: 
                    mapa_colunas["DESC"] = col
                elif "LEITO" in c_upper: 
                    mapa_colunas["LEITO"] = col
                elif "SUGEST" in c_upper or "SUGESTAO" in c_upper: 
                    mapa_colunas["SUGESTAO"] = col

            # Sistema de segurança secundário para os nomes dos Memorandos
            if "MEMO01" not in mapa_colunas:
                for col in df.columns:
                    if "MEMO" in str(col).upper() and ("1" in str(col) or "01" in str(col)):
                        mapa_colunas["MEMO01"] = col
            if "MEMO02" not in mapa_colunas:
                for col in df.columns:
                    if "MEMO" in str(col).upper() and ("2" in str(col) or "02" in str(col)):
                        mapa_colunas["MEMO02"] = col

            coluna_paciente = mapa_colunas.get("PACIENTE", df.columns[0] if len(df.columns) > 0 else "PACIENTE")
            df = df.dropna(subset=[coluna_paciente])
            
            data_extenso_envio = obter_data_por_extenso(data_selecionada)
            arquivos_processados = []
            
            for index, line in df.iterrows():
                nome_do_paciente = str(line[coluna_paciente]).strip()
                if nome_do_paciente.lower() == "nan" or nome_do_paciente == "":
                    continue
                
                memo_01 = str(line.get(mapa_colunas.get("MEMO01", ""), "")).strip()
                memo_02 = str(line.get(mapa_colunas.get("MEMO02", ""), "")).strip()
                
                if memo_01 == "" or memo_01.lower() == "nan":
                    num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" else "S-N"
                else:
                    num_memo_cru = memo_01
                    
                num_notif = limpar_numero_float(line.get(mapa_colunas.get("NOTIF", ""), index + 1))
                if num_notif == "": 
                    num_notif = str(index + 1)
                
                num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
                nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
                
                turno_planilha = str(line.get(mapa_colunas.get("TURNO", ""), "")).strip().upper()
                marca_manha = "X" if "MANH" in turno_planilha else " "
                marca_tarde = "X" if "TARD" in turno_planilha else " "
                marca_noite = "X" if "NOIT" in turno_planilha else " "
                
                dados_memorando = {
                    "{{numero_memorando}}": num_memo_cru,
                    "{{gestor}}": limpar_nan(line.get("Gestor 01", "")),
                    "{{setor}}": limpar_nan(line.get("SETOR NOTIFICADO", "")),
                    "{{notificacao_n}}": num_notif,
                    "{{data_notificacao}}": formatar_data_br(line.get(mapa_colunas.get("DATA_NOTIF", ""), "")),
                    "{{data_ocorrencia}}": formatar_data_br(line.get(mapa_colunas.get("DATA_OCORR", ""), "")),
                    "{{localizacao}}": limpar_nan(line.get("ONDE OCORREU INCIDENTE", "")), 
                    "{{classificacao_incidente}}": limpar_nan(line.get("CLASSIFICAÇÃO DO INCIDENTE", "")), 
                    "{{setor_notificante}}": limpar_nan(line.get("SETOR NOTIFICANTE", "")), 
                    "{{tipo_incidente}}": limpar_nan(line.get(mapa_colunas.get("TIPO", ""), "")).replace("_", " "),
                    "{{descricao_notificacao}}": limpar_nan(line.get(mapa_colunas.get("DESC", ""), "")),
                    "{{nome_paciente}}": nome_do_paciente,
                    "{{leito}}": limpar_numero_float(line.get(mapa_colunas.get("LEITO", ""), "")),
                    "{{sugestao}}": limpar_nan(line.get(mapa_colunas.get("SUGESTAO", ""), "")),
                    "{{m}}": marca_manha,
                    "{{t}}": marca_tarde,
                    "{{n}}": marca_noite,
                    "{{data_envio}}": data_extenso_envio 
                }
                
                doc_instancia = Document(caminho_modelo)
                substituir_texto_protegendo_logos(doc_instancia, dados_memorando)
                buffer_bytes = io.BytesIO()
                doc_instancia.save(buffer_bytes)
                buffer_bytes.seek(0)
                
                arquivos_processados.append({
                    "paciente": nome_do_paciente,
                    "nome_arquivo": nome_base_arquivo,
                    "conteudo": buffer_bytes.getvalue()
                })
                
            st.session_state[nome_chave_cache] = arquivos_processados 

    if nome_chave_cache in st.session_state:
        st.success(f"📋 Lista de verificação pronta! {len(st.session_state[nome_chave_cache])} registros processados.")
        
        for idx, item in enumerate(st.session_state[nome_chave_cache]):
            col_nome, col_word, col_pdf = st.columns(3)
            
            with col_nome:
                st.markdown(f"**🔹 {item['paciente']}**")
                
            with col_word:
                st.download_button(
                    label="📝 Baixar WORD",
                    data=item["conteudo"],
                    file_name=f"{item['nome_arquivo']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"word_btn_{idx}"
                )
                
            with col_pdf:
                st.download_button(
                    label="📕 Baixar PDF",
                    data=item["conteudo"],
                    file_name=f"{item['nome_arquivo']}.pdf",
