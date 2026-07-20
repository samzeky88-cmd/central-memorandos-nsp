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

if arquivo_excel:
    nome_chave_cache = f"dados_memos_{arquivo_excel.name}_{data_selecionada.strftime('%Y%m%d')}"
    
    if nome_chave_cache not in st.session_state:
        with st.spinner("📦 Processando e estruturando os memorandos na memória... Aguarde um instante."):
            # Lê o Excel sem processar cabeçalhos de forma rígida
            df = pd.read_excel(arquivo_excel)
            
            data_extenso_envio = obter_data_por_extenso(data_selecionada)
            arquivos_processados = []
            
            # Identificação dinâmica da saudação temporal baseada no relógio do sistema
            hora_atual = datetime.now().hour
            saudacao = "Bom dia Prezados" if hora_atual < 12 else "Boa Tarde Prezados"
            
            for index, line in df.iterrows():
                # 🎯 LEITURA GARANTIDA E BLINDADA POR ÍNDICES DE POSIÇÃO FÍSICA DA PLANILHA
                try:
                    val_paciente = df.iloc[index, 1] # Coluna B (Paciente)
                    if pd.isna(val_paciente) or str(val_paciente).strip() == "" or str(val_paciente).strip().lower() == "nan":
                        continue
                    nome_do_paciente = str(val_paciente).strip()
                except:
                    continue
                
                # Captura direta dos dados pelas posições das células
                try: num_notif = limpar_numero_float(df.iloc[index, 0]) # Coluna A (Nº Notificação)
                except: num_notif = str(index + 1)
                
                try: dt_notif = formatar_data_br(df.iloc[index, 4]) # Coluna E (Data Notificação)
                except: dt_notif = ""
                
                try: dt_ocorr = formatar_data_br(df.iloc[index, 5]) # Coluna F (Data Ocorrência)
                except: dt_ocorr = ""
                
                try: turno_planilha = str(df.iloc[index, 6]).strip().upper() # Coluna G (Turno)
                except: turno_planilha = ""
                
                try: tipo_incidente = limpar_nan(df.iloc[index, 10]).replace("_", " ") # Coluna K (Tipo)
                except: tipo_incidente = ""
                
                try: texto_sugestao = str(df.iloc[index, 12]).strip() # Coluna M (Sugestão)
                except: texto_sugestao = ""
                
                try: memo_01 = str(df.iloc[index, 15]).strip() # Coluna P (Memo 01)
                except: memo_01 = ""
                
                try: memo_02 = str(df.iloc[index, 19]).strip() # Coluna T (Memo 02)
                except: memo_02 = ""
                
                # Unificação inteligente das colunas alternadas de memorando
                if memo_01 == "" or memo_01.lower() == "nan" or memo_01.lower() == "none":
                    num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" and memo_02.lower() != "none" else "S-N"
                else:
                    num_memo_cru = memo_01
                
                num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
                nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
                
                marca_manha = "X" if "MANH" in turno_planilha else " "
                marca_tarde = "X" if "TARD" in turno_planilha else " "
                marca_noite = "X" if "NOIT" in turno_planilha else " "
                
                # Varredura para encontrar colunas nominativas remanescentes (que não mudam de lugar)
                gestor_val = ""
                setor_val = ""
                onde_val = ""
                classif_val = ""
                setor_notif_val = ""
                leito_val = ""
                
                for c_idx, col_name in enumerate(df.columns):
                    c_up = str(col_name).upper()
                    if "GESTOR" in c_up: gestor_val = str(line[col_name])
                    elif "SETOR NOTIFICADO" in c_up: setor_val = str(line[col_name])
                    elif "ONDE OCORREU" in c_up: onde_val = str(line[col_name])
                    elif "CLASSIFICA" in c_up: classif_val = str(line[col_name])
                    elif "SETOR NOTIFICANTE" in c_up: setor_notif_val = str(line[col_name])
                    elif "LEITO" in c_up: leito_val = str(line[col_name])

                dados_memorando = {
                    "{{numero_memorando}}": num_memo_cru,
                    "{{gestor}}": limpar_nan(gestor_val),
                    "{{setor}}": limpar_nan(setor_val),
                    "{{notificacao_n}}": num_notif,
                    "{{data_notificacao}}": dt_notif,
                    "{{data_ocorrencia}}": dt_ocorr,
                    "{{localizacao}}": limpar_nan(onde_val), 
                    "{{classificacao_incidente}}": limpar_nan(classif_val), 
                    "{{setor_notificante}}": limpar_nan(setor_notif_val), 
                    "{{tipo_incidente}}": tipo_incidente,
                    "{{descricao_notificacao}}": limpar_nan(df.iloc[index, 11]), # Coluna L (Descrição)
                    "{{nome_paciente}}": nome_do_paciente,
                    "{{leito}}": limpar_numero_float(leito_val),
                    "{{sugestao}}": limpar_nan(texto_sugestao),
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
                
                # Corpo de e-mail institucional exato solicitado por você
                texto_email_formatado = (
                    f"{saudacao}\n\n"
                    f"Estamos encaminhando o Memorando Nº {num_memo_cru} em anexo para ser analisado e respondido (via e-mail) em até 15 dias após a data presente.\n\n"
                    f"ATENÇÃO: A resposta via e-mail deve constar um arquivo em forma de word ou PDF para arquivamento de respostas conforme rotina institucional.\n"
                    f"Não serão aceitas mensagens via e-mail sem arquivo como resposta.\n\n"
                    f"Segue abaixo a notificação para análise do incidente em equipe e resposta ao NSP\n\n"
                    f"Atenciosamente,\n"
                    f"Ezequias S. Santos\n"
                    f"Agente Administrativo NAQH"
                )
                
                email_destino = ""
                for col_name in df.columns:
                    if "EMAIL" in str(col_name).upper(): email_destino = limpar_nan(line[col_name])
                
                assunto_email = f"NSP - Memorando {num_memo_cru} (Notificação {num_notif})"
