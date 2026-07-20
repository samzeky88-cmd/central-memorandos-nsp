import os
import pandas as pd
import streamlit as st
import io
import urllib.parse
from datetime import datetime
from docx import Document

st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
st.title("📄 Emissor de Memorandos Individuais - Hospital Dr. Jackson Lago")

st.markdown("### 🗓️ Configuração da Data de Envio")
data_selecionada = st.date_input("Selecione a data que sairá no cabeçalho do Memorando:", datetime.now())

arquivo_excel = st.file_uploader("Suba a planilha contendo os incidentes (.xlsx)", type=["xlsx"])
caminho_modelo = "modelo_memorando.docx"

def substituir_texto_protegendo_logos(doc, dicionario_tags):
    """Substitui o texto alterando apenas os 'runs' para proteger imagens, rodapés e cabeçalhos."""
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
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan":
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
        if isinstance(float(valor), float):
            v_str = str(valor).strip()
            if v_str.endswith(".0"):
                return str(int(float(v_str)))
            return str(int(float(v_str)))
        return str(valor).strip()
    except:
        v_str = str(valor).strip()
        if v_str.endswith(".0"):
            return v_str[:-2]
        return v_str

def obter_data_por_extenso(dt):
    """Gera a data selecionada por extenso em português brasileiro (Ex: 20 de Julho de 2026)"""
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return f"{dt.day} de {meses[dt.month]} de {dt.year}"

# Execução principal após upload do arquivo
if arquivo_excel:
    nome_chave_cache = f"dados_memos_{arquivo_excel.name}_{data_selecionada.strftime('%Y%m%d')}"
    
    if nome_chave_cache not in st.session_state:
        with st.spinner("📦 Processando e estruturando os memorandos na memória... Aguarde um instante."):
            try:
                df = pd.read_excel(arquivo_excel, header=None)
                
                if df.empty:
                    st.warning("⚠️ O arquivo Excel enviado está vazio.")
                    st.stop()
                    
                data_extenso_envio = obter_data_por_extenso(data_selecionada)
                arquivos_processados = []
                
                hora_atual = datetime.now().hour
                saudacao = "Bom dia Prezados" if hora_atual < 12 else "Boa Tarde Prezados"
                num_colunas_df = len(df.columns)
                c_busca = 0 
                
                for index, line in df.iterrows():
                    try:
                        if index not in df.index or len(df.iloc[index]) < 2:
                            continue
                            
                        val_paciente = df.iloc[index, 1]  # Coluna B
                        texto_paciente_teste = str(val_paciente).strip().upper()
                        
                        if pd.isna(val_paciente) or texto_paciente_teste == "" or texto_paciente_teste == "NAN" or "PACIENTE" in texto_paciente_teste or "NOME" in texto_paciente_teste:
                            continue
                        nome_do_paciente = str(val_paciente).strip()
                    except:
                        continue
                    
                    try: num_notif = limpar_numero_float(df.iloc[index, 0]) if num_colunas_df > 0 else str(index + 1)
                    except: num_notif = str(index + 1)
                    
                    try: dt_notif = formatar_data_br(df.iloc[index, 4]) if num_colunas_df > 4 else ""
                    except: dt_notif = ""
                    
                    try: dt_ocorr = formatar_data_br(df.iloc[index, 5]) if num_colunas_df > 5 else ""
                    except: dt_ocorr = ""
                    
                    try: turno_planilha = str(df.iloc[index, 6]).strip().upper() if num_colunas_df > 6 else ""
                    except: turno_planilha = ""
                    
                    try: tipo_incidente = limpar_nan(df.iloc[index, 10]).replace("_", " ") if num_colunas_df > 10 else ""
                    except: tipo_incidente = ""
                    
                    try: texto_sugestao = str(df.iloc[index, 12]).strip() if num_colunas_df > 12 else ""
                    except: texto_sugestao = ""
                    
                    try: memo_01 = str(df.iloc[index, 15]).strip() if num_colunas_df > 15 else ""
                    except: memo_01 = ""
                    
                    try: memo_02 = str(df.iloc[index, 19]).strip() if num_colunas_df > 19 else ""
                    except: memo_02 = ""
                    
                    if memo_01 == "" or memo_01.lower() == "nan" or memo_01.lower() == "none":
                        num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" and memo_02.lower() != "none" else "S-N"
                    else:
                        num_memo_cru = memo_01
                        
                    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
                    nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
                    
                    marca_manha = "X" if "MANH" in turno_planilha else " "
                    marca_tarde = "X" if "TARD" in turno_planilha else " "
                    marca_noite = "X" if "NOIT" in turno_planilha else " "
                    
                    gestor_val, setor_val, onde_val, classif_val, setor_notif_val, leito_val = "", "", "", "", "", ""
                    
                    if num_colunas_df > 0:
                        for r_busca in range(min(index, 15)):
                            for c_temp in range(num_colunas_df):
                                txt_c = str(df.iloc[r_busca, c_temp]).strip().upper()
                                if "GESTOR" in txt_c: gestor_val = str(df.iloc[index, c_temp])
                                elif "SETOR NOTIFICADO" in txt_c: setor_val = str(df.iloc[index, c_temp])
                                elif "ONDE OCORREU" in txt_c: onde_val = str(df.iloc[index, c_temp])
                                elif "CLASSIFICA" in txt_c: classif_val = str(df.iloc[index, c_temp])
                                elif "SETOR NOTIFICANTE" in txt_c: setor_notif_val = str(df.iloc[index, c_temp])
                                elif "LEITO" in txt_c: leito_val = str(df.iloc[index, c_temp])
                                
                    if not gestor_val: gestor_val = "GESTOR DE ENFERMAGEM"
                    if not setor_val: setor_val = "ALA B"
                    
                    email_destino = ""
                    try:
                        if num_colunas_df > c_busca:
                            email_destino = limpar_nan(df.iloc[index, c_busca])
                    except:
                        email_destino = ""
                    
                    texto_descricao = limpar_nan(df.iloc[index, 11]) if num_colunas_df > 11 else ""
                    texto_email_formatado = f"{saudacao},\n\nSegue em anexo o Memorando {num_memo_cru} referente à Notificação {num_notif}.\n\nPaciente: {nome_do_paciente}\nDescrição do ocorrido: {texto_descricao}"
                    assunto_email = f"NSP - Memorando {num_memo_cru} (Notificação {num_notif})"
                    
                    link_gmail = (
                        f"https://google.com?"
                        f"view=cm&fs=1&tf=1"
                        f"&to={urllib.parse.quote(email_destino)}"
                        f"&su={urllib.parse.quote(assunto_email)}"
                        f"&body={urllib.parse.quote(texto_email_formatado)}"
                    )
                    
                    dados_memorando = {
                        "{{numero_memorando}}": num_memo_cru, "{{gestor}}": limpar_nan(gestor_val),
                        "{{setor}}": limpar_nan(setor_val), "{{notificacao_n}}": num_notif,
                        "{{data_notificacao}}": dt_notif, "{{data_ocorrencia}}": dt_ocorr,
                        "{{localizacao}}": limpar_nan(onde_val) if onde_val else "Ala B",
                        "{{classificacao_incidente}}": limpar_nan(classif_val) if classif_val else "Incidente com dano moderado",
                        "{{setor_notificante}}": limpar_nan(setor_notif_val) if setor_notif_val else "SALA VERMELHA",

