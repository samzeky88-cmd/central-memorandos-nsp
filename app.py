import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="NSP - Central de Memorandos", page_icon="🏥", layout="wide")
st.title("🏥 Central de Notificações e Memorandos - NSP")
st.write("Hospital da Cidade Dr. Jackson Lago - São Luís")

CAMINHO_MODELO = "modelo_memorando.docx"
CAMINHO_ROTEIRO = "roteiro_tratativa.docx"

def formatar_data_br(valor_data):
    if pd.isna(valor_data) or str(valor_data).strip() == "":
        return ""
    try:
        dt = pd.to_datetime(valor_data)
        return dt.strftime("%d/%m/%Y")
    except:
        return str(valor_data)

def formatar_data_extenso(valor_data):
    if pd.isna(valor_data) or str(valor_data).strip() == "":
        return ""
    meses = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
        5: "maio", 6: "junho", 7: "julho", 8: "agosto",
        9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    try:
        dt = pd.to_datetime(valor_data)
        return f"{dt.day} de {meses[dt.month]} de {dt.year}"
    except:
        return str(valor_data)

# Substituição segura mantendo imagens e estilos
def substituir_seguro(doc, dados_dinamicos, dt_extenso):
    for p in doc.paragraphs:
        if "{{data_notificacao}}" in p.text and "São Luís" in p.text:
            p.text = f"São Luís, {dt_extenso}"
        else:
            for tag, valor in dados_dinamicos.items():
                if tag in p.text:
                    for run in p.runs:
                        if tag in run.text:
                            run.text = run.text.replace(tag, valor)
                            
    for tabela in doc.tables:
        for linha_tab in tabela.rows:
            for celula in linha_tab.cells:
                for p in celula.paragraphs:
                    if "{{data_notificacao}}" in p.text and "São Luís" in p.text:
                        p.text = f"São Luís, {dt_extenso}"
                    else:
                        for tag, valor in dados_dinamicos.items():
                            if tag in p.text:
                                for run in p.runs:
                                    if tag in run.text:
                                        run.text = run.text.replace(tag, valor)

arquivo_excel = st.file_uploader("1. Faça o upload da sua planilha de notificações (.xlsx)", type=["xlsx"])

if arquivo_excel:
    if not os.path.exists(CAMINHO_MODELO) or not os.path.exists(CAMINHO_ROTEIRO):
        st.error("❌ Arquivos de modelo não encontrados no GitHub.")
    else:
        df = pd.read_excel(arquivo_excel)
        df.columns = df.columns.astype(str).str.strip()
        df = df.dropna(subset=["Nº"])
        
        df_pendentes = df[df["STATUS"].astype(str).str.strip().str.upper() == "PENDENTE"] if "STATUS" in df.columns else df
            
        if len(df_pendentes) == 0:
            st.success("✨ Nenhum memorando pendente encontrado!")
        else:
            st.subheader(f"📋 Painel de Conferência ({len(df_pendentes)} itens)")
            st.dataframe(df_pendentes[[c for c in ["Nº", "STATUS", "NOME"] if c in df_pendentes.columns]])
            
            for index, linha in df_pendentes.iterrows():
                num_notif = str(int(linha.get("Nº", 0)))
                raw_memo = linha.get("Nº Memo 02", linha.get("Nº MEMO", "0000"))
                num_memo = str(raw_memo).split('.')[0] if pd.notna(raw_memo) else "0000"
                memo_formatado = f"{num_memo}/2026" if "/" not in num_memo else num_memo
                
                paciente = str(linha.get("NOME", "Paciente"))
                setor_destino = str(linha.get("GESTOR DE QUEM VAI RECEBER O GMAIL", linha.get("SETOR NOTIFICADO 02", "")))
                email_destino = str(linha.get("EMAIL_SETOR", "")).strip()
                
                dt_ocorrencia_br = formatar_data_br(linha.get("DATA DA OCORRÊNCIA", ""))
                dt_notificacao_br = formatar_data_br(linha.get("DATA DA NOTIFICAÇÃO", ""))
                dt_extenso_br = formatar_data_extenso(linha.get("DATA DA NOTIFICAÇÃO", datetime.now()))
                
                st.markdown(f"### 📄 Notificação Nº {num_notif}")
                
                doc = Document(CAMINHO_MODELO)
                
                dados_dinamicos = {
                    "{{numero_memorando}}": memo_formatado,
                    "{{notificacao_n}}": num_notif,
                    "{{data_ocorrencia}}": dt_ocorrencia_br,
                    "{{data_notificacao}}": dt_notificacao_br,
                    "{{localizacao}}": str(linha.get("ONDE OCORREU INCIDENTE", linha.get("LOCALIZAÇÃO", ""))),
                    "{{tipo_incidente}}": str(linha.get("TIPO DE INCIDENTE", "")),
                    "{{classificacao_incidente}}": str(linha.get("CLASSIFICAÇÃO DO INCIDENTE", "")),
                    "{{descricao_notificacao}}": str(linha.get("DESCRIÇÃO DA NOTIFICAÇÃO", "")),
                    "{{setor_notificante}}": str(linha.get("SETOR NOTIFICANTE", "")),
                    "{{setor_destino}}": setor_destino,
                    "{{nome_paciente}}": paciente,
                    "{{leito}}": str(linha.get("LEITO", "")),
                    "{{sugestao}}": str(linha.get("SUGESTÃO", ""))
                }
                
                # Marcação de turnos precisa do espaço idêntico ao modelo do seu Word
                turno = str(linha.get("TURNO", "")).strip().upper()
                dados_dinamicos["{{m}}"] = "X" if "MANH" in turno else " "
                dados_dinamicos["{{t}}"] = "X" if "TARD" in turno else " "
                dados_dinamicos["{{n}}"] = "X" if "NOIT" in turno else " "
                
                substituir_seguro(doc, dados_dinamicos, dt_extenso_br)
                                        
                nome_word = f"MEMORANDO_Nº_{num_memo}_NOTIFICAÇÃO_Nº_{num_notif}_I_NSP.docx"
                doc.save(nome_word)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(label="📥 Baixar em Word", data=open(nome_word, "rb").read(), file_name=nome_word, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"d_{index}")
                with col2:
                    if st.button("🚀 Enviar E-mail", key=f"b_{index}"):
                        msg = EmailMessage()
                        msg["Subject"] = f"MEMORANDO Nº {num_memo} - NOTIFICAÇÃO Nº {num_notif} _I_NSP"
                        msg["From"] = st.secrets["EMAIL_USER"]
                        msg["To"] = email_destino
                        msg.set_content(f"Prezados,\n\nSeguem em anexo os documentos do Memorando Nº {memo_formatado}.\n\nAtenciosamente,\nNSP.")
                        msg.add_attachment(open(nome_word, "rb").read(), maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename=nome_word)
                        msg.add_attachment(open(CAMINHO_ROTEIRO, "rb").read(), maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename="Roteiro_Para_Tratativa_NSP.docx")
                        
                        with smtplib.SMTP_SSL("://gmail.com", 465) as smtp:
                            smtp.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASSWORD"])
                            smtp.send_message(msg)
                        st.success("✅ Enviado com sucesso!")
                        
                if os.path.exists(nome_word):
                    os.remove(nome_word)
                st.markdown("---")
