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

# ✅ E-MAIL — Baseado EXCLUSIVAMENTE no SETOR NOTIFICADO
def buscar_email_setor(setor_notificado):
    if not setor_notificado or str(setor_notificado).strip() == "":
        return "nspsoc2@gmail.com"
    
    nome = str(setor_notificado).strip().upper()
    
    emails = {
        "ALA D": "ortopediaalad.hcid@gmail.com",
        "AGÊNCIA TRANSFUSIONAL": "agenciatransfusional.hcid@gmail.com",
        "ALA A": "alaaclinicamedica@gmail.com",
        "ALA F (UTI B)": "utihcid@gmail.com",
        "ALA I": "ortopediaalai.hcid@gmail.com",
        "ALA F (UTI C)": "utihcid@gmail.com",
        "ALA G": "aladgh.hcid@gmail.com",
        "SALA VERMELHA": "salavermelhasdm@gmail.com",
        "ALA J": "ortopediaalaj.hcid@gmail.com",
        "FARMÁCIA": "farmacia.hcid@gmail.com",
        "CENTRAL INTERNA DE TRANSPORTE": "centraltransporte.hcid@gmail.com",
        "HOTELARIA": "hotelaria.hcid@gmail.com",
        "ALA F (UTI A)": "utihcid@gmail.com",
        "ECP - EQUIPE DE CUIDADOS COM A PELE": "carlosemilioecp@gmail.com",
        "AUXILIAR DE TRANSPORTE": "centraltransporte.hcid@gmail.com",
        "FISIOTERAPIA": "fisioreabsoc2@gmail.com",
        "TERAPIA OCUPACIONAL": "terapiaocupacional.hcid@gmail.com",
        "SERVIÇO SOCIAL": "servicosocialhcid@gmail.com",
        "NUTRIÇÃO": "nutricaohospitaldacidade@gmail.com",
    }
    
    for chave, email in emails.items():
        if chave == nome or chave in nome or nome in chave:
            return email
    
    return "nspsoc2@gmail.com"

# ✅ GESTOR — Baseado EXCLUSIVAMENTE no SETOR NOTIFICADO
def buscar_gestor(setor_notificado):
    if not setor_notificado or str(setor_notificado).strip() == "":
        return ""
    
    nome = str(setor_notificado).strip().upper()
    
    gestores = {
        "ALA D": "GESTOR DE ENFERMAGEM",
        "AGÊNCIA TRANSFUSIONAL": "GESTOR DE AGÊNCIA",
        "ALA A": "GESTOR DE ENFERMAGEM",
        "ALA F (UTI B)": "GESTOR DE ENFERMAGEM",
        "ALA I": "GESTOR DE ENFERMAGEM",
        "ALA F (UTI C)": "GESTOR DE ENFERMAGEM",
        "ALA G": "GESTOR DE ENFERMAGEM",
        "SALA VERMELHA": "GESTOR DE ENFERMAGEM",
        "ALA J": "GESTOR DE ENFERMAGEM",
        "FARMÁCIA": "GESTOR DE FARMÁCIA",
        "CENTRAL INTERNA DE TRANSPORTE": "GESTOR DA CENTRAL INTERNA DE TRANSPORTE",
        "HOTELARIA": "GESTOR DE HOTELARIA",
        "ALA F (UTI A)": "GESTOR DE ENFERMAGEM",
        "ECP - EQUIPE DE CUIDADOS COM A PELE": "GESTOR DA ECP",
        "AUXILIAR DE TRANSPORTE": "GESTOR DA CENTRAL INTERNA DE TRANSPORTE",
        "FISIOTERAPIA": "GESTOR DE FISIOTERAPIA",
        "TERAPIA OCUPACIONAL": "GESTOR DE TERAPIA",
        "SERVIÇO SOCIAL": "GESTOR DE SERVIÇO SOCIAL",
        "NUTRIÇÃO": "GESTOR DE NUTRIÇÃO",
    }
    
    for chave, gestor in gestores.items():
        if chave == nome or chave in nome or nome in chave:
            return gestor
    
    return ""

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

def contar_enviados(df, coluna_status=1):
    qtd = 0
    for _, linha in df.iterrows():
        if len(linha) <= coluna_status:
            continue
        status = tratar_texto(linha.iloc[coluna_status]).upper()
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
        col2.metric("✅ Já Enviados", enviados)
        col3.metric("⏳ Pendentes", pendentes)
        st.info(f"{total} - {enviados} = {pendentes} memorandos pendentes")
        st.divider()
        
        data_atual = datetime.now()
        ano_atual = data_atual.year
        data_extenso = data_por_extenso(data_atual)
        
        qtd_processados = 0
        
        for idx, linha in df.iterrows():
            # ✅ IGNORAR LINHA SE SETOR NOTIFICADO ESTIVER VAZIO
            setor_notificado = tratar_texto(linha.iloc[14]) if len(linha) > 14 else ""
            if not setor_notificado:
                continue
            
            num_notif = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
            if not num_notif or num_notif.upper() == "STATUS" or "NOTIF" in str(num_notif).upper():
                continue
            
            nome_paciente = tratar_texto(linha.iloc[9]) if len(linha) > 9 else ""
            if not nome_paciente:
                continue
            
            status = tratar_texto(linha.iloc[1]).upper() if len(linha) > 1 else ""
            enviado = (status == "ENVIADO")
            
            data_ocorrencia = formatar_data_sem_hora(linha.iloc[2]) if len(linha) > 2 else ""
            data_notificacao = formatar_data_sem_hora(linha.iloc[3]) if len(linha) > 3 else ""
            turno = tratar_texto(linha.iloc[4]) if len(linha) > 4 else ""
            localizacao = tratar_texto(linha.iloc[5]) if len(linha) > 5 else ""
            tipo_incidente = tratar_texto(linha.iloc[6]) if len(linha) > 6 else ""
            classificacao_incidente = tratar_texto(linha.iloc[7]) if len(linha) > 7 else ""
            descricao_notificacao = tratar_texto(linha.iloc[8]) if len(linha) > 8 else ""
            leito = tratar_texto(linha.iloc[10]) if len(linha) > 10 else ""
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
            
            # ✅ SÓ SETOR NOTIFICADO → GESTOR E E-MAIL
            gestor = buscar_gestor(setor_notificado)
            email_destino = buscar_email_setor(setor_notificado)
            
            qtd_processados += 1
            
            hora_atual = datetime.now().hour
            if 6 <= hora_atual < 12:
                saudacao = "Bom dia,"
            elif 12 <= hora_atual < 18:
                saudacao = "Boa tarde,"
            else:
                saudacao = "Boa noite,"
            
            if enviado:
                titulo_memo = f"~~✅ Notificação {num_notif} — {nome_paciente}~~  **ENVIADO**"
            else:
                titulo_memo = f"📋 Notificação {num_notif} — {nome_paciente}"
            
            with st.expander(titulo_memo, expanded=not enviado):
                if enviado:
                    st.success("✅ Este memorando já foi enviado!")
                    st.divider()
                
                st.write(f"📅 Data de Emissão: {data_extenso}")
                st.write(f"📅 Data da Ocorrência: {data_ocorrencia} | Notificação: {data_notificacao}")
                st.write(f"📍 Local: {localizacao}")
                st.write(f"📤 **Setor Notificado:** {setor_notificado}")
                st.write(f"👤 **Gestor:** {gestor}")
                st.write(f"⚠️ Tipo: {tipo_incidente} | Classificação: {classificacao_incidente}")
                st.write(f"👤 Paciente: {nome_paciente} | Leito: {leito}")
                st.write(f"🔄 Sugestão: {sugestao}")
                st.write(f"📧 **E-mail para envio:** {email_destino}")
                
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
                    "{{setor_notificado}}": setor_notificado,
                    "{{data_envio}}": data_extenso,
                    "{{data_ocorrencia}}": data_ocorrencia,
                    "{{data_notificacao}}": data_notificacao,
                    "{{m}}": manha,
                    "{{t}}": tarde,
                    "{{n}}": noite,
                    "{{localizacao}}": localizacao,
                    "{{tipo_incidente}}": tipo_incidente,
                    "{{classificacao_incidente}}": classificacao_incidente,
                    "{{descricao_notificacao}}": descricao_notificacao,
                    "{{nome_paciente}}": nome_paciente,
                    "{{leito}}": leito,
                    "{{sugestao}}": sugestao,
                }
                
                arquivo_pronto = gerar_docx(dados_docx)
                if arquivo_pronto:
                    st.download_button(
                        label="📄 Baixar Memorando",
                        data=arquivo_pronto,
                        file_name=f"NSP_Nº {num_notif}_{ano_atual}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"botao_{num_notif}_{idx}",
                        disabled=enviado
                    )
    except Exception as e:
        st.error(f"Erro ao processar planilha: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
