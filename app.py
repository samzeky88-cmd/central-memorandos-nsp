# =================================================================
# GERADOR DE MEMORANDOS — HOSPITAL DA CIDADE DR. JACKSON LAGO
# =================================================================

import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document

TOTAL_GERAL_MEMORANDOS = 964

# =================================================================
# FUNÇÕES AUXILIARES
# =================================================================

def tratar_texto(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip()

def limpar_numero(valor):
    if pd.isna(valor):
        return ""
    try:
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor).strip()
    except:
        return str(valor)

def formatar_data_sem_hora(valor):
    if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().upper() == "NAN":
        return ""
    try:
        dt = pd.to_datetime(valor)
        return dt.strftime("%d/%m/%Y")
    except:
        return str(valor).strip().split(" ")[0]

def data_por_extenso(dt):
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    return f"{dt.day} de {meses[dt.month-1]} de {dt.year}"

def buscar_email_setor(nome_setor):
    if not nome_setor or str(nome_setor).strip() == "":
        return "nspsoc2@gmail.com"
    
    nome = str(nome_setor).strip().upper()
    
    if "CENTRAL DE TRANSPORTE" in nome or "TRANSPORTE EXTERNO" in nome:
        return "centraltransporte.hcid@gmail.com"
    
    return "nspsoc2@gmail.com"

def substituir_texto_no_docx(doc, dados):
    for p in doc.paragraphs:
        for chave, valor in dados.items():
            if chave in p.text:
                p.text = p.text.replace(chave, str(valor))
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                substituir_texto_no_docx(celula, dados)

def gerar_docx(dados):
    try:
        with open("modelo_memorando.docx", "rb") as f:
            doc = Document(f)
        substituir_texto_no_docx(doc, dados)
        saida = io.BytesIO()
        doc.save(saida)
        saida.seek(0)
        return saida
    except FileNotFoundError:
        st.error("ERRO: Coloque o arquivo modelo_memorando.docx na mesma pasta do app.py")
        return None
    except Exception as e:
        st.error(f"Erro ao gerar arquivo: {str(e)}")
        return None

def contar_enviados(df):
    qtd = 0
    for _, linha in df.iterrows():
        if len(linha) <= 1:
            continue
        status = tratar_texto(linha.iloc[1]).upper()
        if status == "ENVIADO":
            qtd += 1
    return qtd

# =================================================================
# INTERFACE PRINCIPAL
# =================================================================

st.set_page_config(page_title="Emissor de Memorandos", page_icon="📄", layout="wide")
st.title("📄 Emissor de Memorandos — Hospital da Cidade Dr. Jackson Lago")
st.divider()

arquivo_excel = st.file_uploader("Carregue a Planilha (.xlsx)", type=["xlsx"])

if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel, engine="openpyxl")
        
        enviados = contar_enviados(df)
        total = TOTAL_GERAL_MEMORANDOS
        pendentes = total - enviados
        
        st.subheader("RESUMO GERAL")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Memorandos", total)
        col2.metric("Já Enviados", enviados)
        col3.metric("Pendentes", pendentes)
        st.info(f"{total} - {enviados} = {pendentes} memorandos pendentes")
        st.divider()
        
        data_atual = datetime.now()
        ano_atual = data_atual.year
        data_extenso = data_por_extenso(data_atual)
        
        qtd_processados = 0
        
        for idx, linha in df.iterrows():
            num_notif = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
            if not num_notif or num_notif.upper() == "STATUS" or "NOTIF" in str(num_notif).upper():
                continue
            
            nome_paciente = tratar_texto(linha.iloc[9]) if len(linha) > 9 else ""
            if not nome_paciente:
                continue
            
            data_ocorrencia = formatar_data_sem_hora(linha.iloc[2]) if len(linha) > 2 else ""
            data_notificacao = formatar_data_sem_hora(linha.iloc[3]) if len(linha) > 3 else ""
            turno = tratar_texto(linha.iloc[4]) if len(linha) > 4 else ""
            localizacao = tratar_texto(linha.iloc[5]) if len(linha) > 5 else ""
            tipo_incidente = tratar_texto(linha.iloc[6]) if len(linha) > 6 else ""
            classificacao_incidente = tratar_texto(linha.iloc[7]) if len(linha) > 7 else ""
            descricao_notificacao = tratar_texto(linha.iloc[8]) if len(linha) > 8 else ""
            leito = tratar_texto(linha.iloc[10]) if len(linha) > 10 else ""
            setor_notificante = tratar_texto(linha.iloc[11]) if len(linha) > 11 else ""
            sugestao = tratar_texto(linha.iloc[12]) if len(linha) > 12 else ""
            
            manha = "☐"
            tarde = "☐"
            noite = "☐"
            if turno:
                if "MANHÃ" in turno or "MANHA" in turno:
                    manha = "☑"
                elif "TARDE" in turno:
                    tarde = "☑"
                elif "NOITE" in turno:
                    noite = "☑"
            
            gestor = tratar_texto(linha.iloc[13]) if len(linha) > 13 else ""
            setor_destino = tratar_texto(linha.iloc[14]) if len(linha) > 14 else ""
            
            email_destino = buscar_email_setor(setor_notificante)
            
            qtd_processados += 1
            
            # Saudação automática conforme horário
            hora_atual = datetime.now().hour
            if 6 <= hora_atual < 12:
                saudacao = "Bom dia,"
            elif 12 <= hora_atual < 18:
                saudacao = "Boa tarde,"
            else:
                saudacao = "Boa noite,"
            
            with st.expander(f"Notificação {num_notif} — {nome_paciente}", expanded=True):
                st.write(f"📅 Data de Emissão: {data_extenso}")
                st.write(f"📅 Data da Ocorrência: {data_ocorrencia} | Notificação: {data_notificacao}")
                st.write(f"📍 Local: {localizacao} | Setor Notificante: {setor_notificante}")
                st.write(f"⚠️ Tipo: {tipo_incidente} | Classificação: {classificacao_incidente}")
                st.write(f"👤 Paciente: {nome_paciente} | Leito: {leito}")
                st.write(f"🔄 Sugestão: {sugestao}")
                st.write(f"📧 E-mail: {email_destino}")
                
                corpo_email = f"""
{saudacao} Prezado(a),

Segue em anexo o Memorando Nº {num_notif}/{ano_atual} referente à notificação Nº {num_notif}, 
para ser analisado e respondido (via e-mail) em até 15 dias após a data presente.

ATENÇÃO: A resposta via e-mail deve constar um arquivo em forma de Word ou PDF para arquivamento 
de respostas conforme rotina institucional. Não serão aceitas mensagens via e-mail sem arquivo como resposta.

Atenciosamente,

Ezequias S. Santos
Agente Administrativo — NAQH / NSP
Núcleo de Segurança do Paciente
nspsoc2@gmail.com
                """
                
                with st.expander("Texto do E-mail"):
                    st.text(corpo_email.strip())
                
                dados_docx = {
                    "{{numero_memorando}}": f"NSP Nº {num_notif}/{ano_atual}",
                    "{{notificacao_n}}": num_notif,
                    "{{gestor}}": gestor,
                    "{{setor}}": setor_destino,
                    "{{data_ocorrencia}}": data_ocorrencia,
                    "{{data_notificacao}}": data_notificacao,
                    "{{data_extenso}}": data_extenso,
                    "{{m_manha}}": manha,
                    "{{t_tarde}}": tarde,
                    "{{n_noite}}": noite,
                    "{{localizacao}}": localizacao,
                    "{{tipo_incidente}}": tipo_incidente,
                    "{{classificacao_incidente}}": classificacao_incidente,
                    "{{descricao_notificacao}}": descricao_notificacao,
                    "{{nome_paciente}}": nome_paciente,
                    "{{leito}}": leito,
                    "{{setor_not}}": setor_notificante,
                    "{{sugestao}}": sugestao,
                    "{{email_destino}}": email_destino,
                }
                
                arquivo_pronto = gerar_docx(dados_docx)
                if arquivo_pronto:
                    st.download_button(
                        label="📄 Baixar Memorando",
                        data=arquivo_pronto,
                        file_name=f"NSP_Nº {num_notif}_{ano_atual}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"botao_{num_notif}"
                    )
    except Exception as e:
        st.error(f"Erro ao processar planilha: {str(e)}")
