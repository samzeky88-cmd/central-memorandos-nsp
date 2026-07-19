import os
import smtplib
from datetime import datetime
from email.message import EmailMessage
import pandas as pd
import streamlit as st
from docx import Document

# Configuração da página web do Streamlit
st.set_page_config(page_title="NSP - Central de Memorandos", page_icon="🏥", layout="wide")
st.title("🏥 Central de Notificações e Memorandos - NSP")
st.write("Hospital da Cidade Dr. Jackson Lago - São Luís")

# Arquivos fixos que devem estar guardados na raiz do seu repositório do GitHub
CAMINHO_MODELO = "modelo_memorando.docx"
CAMINHO_ROTEIRO = "roteiro_tratativa.docx"  # O modelo limpo que o outro setor vai preencher

# Entrada da Planilha compactada via Upload
arquivo_excel = st.file_uploader("1. Faça o upload da sua planilha de notificações (.xlsx)", type=["xlsx"])

if arquivo_excel:
    if not os.path.exists(CAMINHO_MODELO):
        st.error(f"❌ Erro: O arquivo '{CAMINHO_MODELO}' não foi encontrado no repositório do GitHub.")
    elif not os.path.exists(CAMINHO_ROTEIRO):
        st.error(f"❌ Erro: O arquivo '{CAMINHO_ROTEIRO}' não foi encontrado no repositório do GitHub.")
    else:
        # CORREÇÃO APLICADA: Lê automaticamente a primeira aba da planilha, independente do nome
        df = pd.read_excel(arquivo_excel)
        
        # LIMPEZA: Remove linhas vazias baseando-se na coluna Nº (Notificação)
        df = df.dropna(subset=["Nº"])
        
        # FILTRO: Filtra e exibe apenas o que está marcado estritamente como 'Pendente'
        if "STATUS" in df.columns:
            df_pendentes = df[df["STATUS"].astype(str).str.strip().str.upper() == "PENDENTE"]
        else:
            df_pendentes = df
            
        if len(df_pendentes) == 0:
            st.success("✨ Nenhum memorando pendente encontrado na planilha!")
        else:
            st.subheader(f"📋 Painel de Conferência ({len(df_pendentes)} memorandos identificados)")
            st.dataframe(df_pendentes[["Nº", "STATUS", "SETOR NOTIFICADO 02", "NOME"]])
            st.write("---")
            
            # Varre a planilha linha por linha tratando cada envio de forma isolada e separada
            for index, linha in df_pendentes.iterrows():
                num_notif = str(int(linha.get("Nº", 0)))
                num_memo = str(int(linha.get("Nº Memo 02", 0))) if pd.notna(linha.get("Nº Memo 02")) else "0000"
                paciente = str(linha.get("NOME", ""))
                setor_destino = str(linha.get("GESTOR DE QUEM VAI RECEBER O GMAIL", ""))
                email_destino = str(linha.get("EMAIL_SETOR", "")).strip()
                
                # Identifica dinamicamente se o horário atual pede Bom dia ou Boa tarde
                hora_atual = datetime.now().hour
                saudacao = "Bom dia" if hora_atual < 12 else "Boa tarde"
                
                st.markdown(f"### 📄 Bloco de Envio Isolado: Notificação Nº {num_notif}")
                st.caption(f"**Setor de Destino (PARA):** {setor_destino} | **E-mail Alvo:** {email_destino}")
                
                # Carrega o modelo do Word limpo para evitar contaminação de dados antigos
                doc = Document(CAMINHO_MODELO)
                
                # Mapeamento exato das colunas reais da planilha com as chaves do Word
                dados_dinamicos = {
                    "{{numero_memorando}}": num_memo,
                    "{{notificacao_n}}": num_notif,
                    "{{data_analise}}": str(linha.get("DATA DA ANÁLISE", "")).split(' ')[0],
                    "{{data_notificacao}}": str(linha.get("DATA DA NOTIFICAÇÃO", "")).split(' ')[0],
                    "{{setor_notificante}}": str(linha.get("SETOR NOTIFICANTE", "")),
                    "{{setor_destino}}": setor_destino,
                    "{{nome_paciente}}": paciente,
                }
                
                # Trata a marcação automática dos turnos
                turno = str(linha.get("TURNO", "")).strip().upper()
                dados_dinamicos["{{m}}"] = "X" if turno == "MANHÃ" else " "
                dados_dinamicos["{{t}}"] = "X" if turno == "TARDE" else " "
                dados_dinamicos["{{n}}"] = "X" if turno == "NOITE" else " "
                
                # Substitui os textos nos parágrafos comuns do Word
                for p in doc.paragraphs:
                    for tag, valor in dados_dinamicos.items():
                        if tag in p.text:
                            p.text = p.text.replace(tag, valor)
                            
                # Substitui os textos caso suas tags estejam inseridas dentro de tabelas do Word
                for tabela in doc.tables:
                    for linha_tab in tabela.rows:
                        for celula in linha_tab.cells:
                            for p in celula.paragraphs:
                                for tag, valor in dados_dinamicos.items():
                                    if tag in p.text:
                                        p.text = p.text.replace(tag, valor)
                                        
                # Define o nome do arquivo exatamente no formato padronizado
                nome_arquivo_padrao = f"MEMORANDO_Nº_{num_memo}_NOTIFICAÇÃO_Nº_{num_notif}_I_NSP.docx"
                doc.save(nome_arquivo_padrao)
                
                col1, col2 = st.columns(2)
                
                # BOTÃO DE CHECKAGEM MANUAL ANTES DE ENVIAR
                with col1:
                    with open(nome_arquivo_padrao, "rb") as f_word:
                        st.download_button(
                            label=f"📥 1º Baixar e Revisar Word (MEMO {num_memo})",
                            data=f_word,
                            file_name=nome_arquivo_padrao,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_{index}"
                        )
                        
                # BOTÃO DE DISPARO INDIVIDUAL E ISOLADO VIA GMAIL
                with col2:
                    if st.button(f"🚀 2º Confirmar e Enviar E-mail (MEMO {num_memo})", key=f"btn_{index}"):
                        if not email_destino or "@" not in email_destino:
                            st.error(f"❌ Erro: O e-mail da coluna 'EMAIL_SETOR' está em branco ou incorreto.")
                        else:
                            try:
                                msg = EmailMessage()
                                msg["Subject"] = f"MEMORANDO Nº {num_memo} - NOTIFICAÇÃO Nº {num_notif} _I_NSP"
                                msg["From"] = st.secrets["EMAIL_USER"]
                                msg["To"] = email_destino
                                
                                msg.set_content(
                                    f"{saudacao} Prezados,\n\n"
                                    f"Seguem em anexo os documentos validados e emitidos pelo Núcleo de Segurança do Paciente (NSP) "
                                    f"referentes ao Memorando Nº {num_memo} da Notificação de Incidente Hospitalar Nº {num_notif}.\n\n"
                                    f"Solicitamos que o setor responsável analise e proceda com as tratativas necessárias utilizando a ficha limpa em anexo.\n\n"
                                    f"Atenciosamente,\n"
                                    f"NSP - Hospital da Cidade Dr. Jackson Lago."
                                )
                                
                                with open(nome_arquivo_padrao, "rb") as f_anexo1:
                                    msg.add_attachment(f_anexo1.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename=nome_arquivo_padrao)
                                    
                                with open(CAMINHO_ROTEIRO, "rb") as f_anexo2:
                                    msg.add_attachment(f_anexo2.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.wordprocessingml.document", filename="Roteiro_Para_Tratativa_NSP.docx")
                                    
                                with smtplib.SMTP_SSL("://gmail.com", 465) as smtp:
                                    smtp.login(st.secrets["EMAIL_USER"], st.secrets["EMAIL_PASSWORD"])
                                    smtp.send_message(msg)
                                    
                                st.success(f"✅ Memorando Nº {num_memo} enviado com sucesso para: {email_destino}!")
                            except Exception as e:
                                st.error(f"❌ Erro técnico ao tentar enviar este e-mail: {e}")
                                
                if os.path.exists(nome_arquivo_padrao):
                    os.remove(nome_arquivo_padrao)
                st.markdown("---")
