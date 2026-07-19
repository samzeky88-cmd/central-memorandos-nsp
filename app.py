import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
import pandas as pd
import streamlit as st
from docx import Document
from docx2pdf import convert  # Motor de conversão para PDF na nuvem

# Configuração da página web do Streamlit
st.set_page_config(page_title="NSP - Central de Memorandos", page_icon="🏥", layout="wide")
st.title("🏥 Central de Notificações e Memorandos - NSP")
st.write("Hospital da Cidade Dr. Jackson Lago - São Luís")

# Arquivos fixos que devem estar guardados na raiz do seu repositório do GitHub
CAMINHO_MODELO = "modelo_memorando.docx"
CAMINHO_ROTEIRO = "roteiro_tratativa.docx"

# Função auxiliar para formatar datas vindo do Excel
def formatar_data_br(valor_data):
    if pd.isna(valor_data) or str(valor_data).strip() == "":
        return ""
    try:
        dt = pd.to_datetime(valor_data)
        return dt.strftime("%d/%m/%Y")
    except:
        return str(valor_data)

# Função auxiliar para formatar a data por extenso para a linha de São Luís
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

# Entrada da Planilha compactada via Upload
arquivo_excel = st.file_uploader("1. Faça o upload da sua planilha de notificações (.xlsx)", type=["xlsx"])

if arquivo_excel:
    if not os.path.exists(CAMINHO_MODELO):
        st.error(f"❌ Erro: O arquivo '{CAMINHO_MODELO}' não foi encontrado no repositório do GitHub.")
    elif not os.path.exists(CAMINHO_ROTEIRO):
        st.error(f"❌ Erro: O arquivo '{CAMINHO_ROTEIRO}' não foi encontrado no repositório do GitHub.")
    else:
        df = pd.read_excel(arquivo_excel)
        df.columns = df.columns.astype(str).str.strip()
        df = df.dropna(subset=["Nº"])
        
        if "STATUS" in df.columns:
            df_pendentes = df[df["STATUS"].astype(str).str.strip().str.upper() == "PENDENTE"]
        else:
            df_pendentes = df
            
        if len(df_pendentes) == 0:
            st.success("✨ Nenhum memorando pendente encontrado na planilha!")
        else:
            st.subheader(f"📋 Painel de Conferência ({len(df_pendentes)} memorandos identificados)")
            colunas_existentes = [c for c in ["Nº", "STATUS", "NOME"] if c in df_pendentes.columns]
            st.dataframe(df_pendentes[colunas_existentes])
            st.write("---")
            
            for index, linha in df_pendentes.iterrows():
                num_notif = str(int(linha.get("Nº", 0)))
                
                # Resgata o número do memo e limpa decimais
                raw_memo = linha.get("Nº Memo 02", linha.get("Nº MEMO", "0000"))
                if pd.notna(raw_memo) and str(raw_memo).strip() != "":
                    num_memo = str(raw_memo).split('.')
                else:
                    num_memo = "0000"
                
                memo_formatado = f"{num_memo}/2026" if "/" not in num_memo else num_memo
                
                paciente = str(linha.get("NOME", "Paciente"))
                setor_destino = str(linha.get("GESTOR DE QUEM VAI RECEBER O GMAIL", linha.get("SETOR NOTIFICADO 02", "")))
                email_destino = str(linha.get("EMAIL_SETOR", "")).strip()
                
                leito_paciente = str(linha.get("LEITO", ""))
                setor_notif = str(linha.get("SETOR NOTIFICANTE", ""))
                sugestao_nsp = str(linha.get("SUGESTÃO", ""))
                
                # Tratamento das datas
                dt_ocorrencia_br = formatar_data_br(linha.get("DATA DA OCORRÊNCIA", ""))
                dt_notificacao_br = formatar_data_br(linha.get("DATA DA NOTIFICAÇÃO", ""))
                dt_extenso_br = formatar_data_extenso(linha.get("DATA DA NOTIFICAÇÃO", datetime.now()))
                
                hora_atual = datetime.now().hour
                saudacao = "Bom dia" if hora_atual < 12 else "Boa tarde"
                
                st.markdown(f"### 📄 Bloco de Envio Isolado: Notificação Nº {num_notif}")
                st.caption(f"**Setor de Destino (PARA):** {setor_destino} | **E-mail Alvo:** {email_destino}")
                
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
                    "{{setor_notificante}}": setor_notif,
                    "{{setor_destino}}": setor_destino,
                    "{{nome_paciente}}": paciente,
                    "{{leito}}": leito_paciente,
                    "{{sugestao}}": sugestao_nsp
                }
                
                # Limpeza robusta para a marcação de turnos
                turno = str(linha.get("TURNO", "")).strip().upper()
                dados_dinamicos["{{m}}"] = "X" if "MANH" in turno else " "
                dados_dinamicos["{{t}}"] = "X" if "TARD" in turno else " "
                dados_dinamicos["{{n}}"] = "X" if "NOIT" in turno else " "
                
                # Substituição preservando o Brasão original intacto
                for p in doc.paragraphs:
                    if "{{data_notificacao}}" in p.text and "São Luís" in p.text:
                        p.text = f"São Luís, {dt_extenso_br}"
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
                                    p.text = f"São Luís, {dt_extenso_br}"
                                else:
                                    for tag, valor in dados_dinamicos.items():
                                        if tag in p.text:
                                            for run in p.runs:
                                                if tag in run.text:
                                                    run.text = run.text.replace(tag, valor)
                                        
                # Salva o arquivo em formato Word temporário
                nome_word = f"MEMORANDO_Nº_{num_memo}_NOTIFICAÇÃO_Nº_{num_notif}_I_NSP.docx"
                nome_pdf = f"MEMORANDO_Nº_{num_memo}_NOTIFICAÇÃO_Nº_{num_notif}_I_NSP.pdf"
                doc.save(nome_word)
                
                # 🌟 Cria a versão em PDF na nuvem
                try:
                    convert(nome_word, nome_pdf)
                    pdf_pronto = True
                except:
                    pdf_pronto = False
                
                # Desenha os botões de ação divididos na tela
                c1, c2, c3 = st.columns(3)
                
                with c1:
                    with open(nome_word, "rb") as f_word:
                        st.download_button(
                            label=f"📥 Baixar em Word",
                            data=f_word,
                            file_name=nome_word,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"doc_{index}"
                        )
                        
                with c2:
                    if pdf_pronto:
                        with open(nome_pdf, "rb") as f_pdf:
                            st.download_button(
                                label=f"📄 Baixar em PDF",
                                data=f_pdf,
                                file_name=nome_pdf,
                                mime="application/pdf",
                                key=f"pdf_{index}"
                            )
                    else:
                        st.caption("⚠️ PDF indisponível no momento.")
                        
                with c3:
                    if st.button(f"🚀 Confirmar e Enviar E-mail", key=f"btn_{index}"):
                        if not email_destino or "@" not in email_destino:
                            st.error("❌ Erro: O e-mail da coluna 'EMAIL_SETOR' está incorreto.")
                        else:
                            try:
                                msg = EmailMessage()
                                msg["Subject"] = f"MEMORANDO Nº {num_memo} - NOTIFICAÇÃO Nº {num_notif} _I_NSP"
                                msg["From"] = st.secrets["EMAIL_USER"]
                                msg["To"] = email_destino
                                
                                msg.set_content(
                                    f"{saudacao} Prezados,\n\n"
                                    f"Seguem em anexo os documentos validados e emitidos pelo Núcleo de Segurança do Paciente (NSP) "
