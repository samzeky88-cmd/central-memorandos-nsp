# =================================================================
# GERADOR DE MEMORANDOS — HOSPITAL DA CIDADE DR. JACKSON LAGO
# =================================================================

import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document

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

# =================================================================
# E-MAILS
# =================================================================
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
        "GESTOR DE ENFERMAGEM": "enfermagem.hcid@gmail.com",
        "GESTOR DE GERÊNCIA DE ENFERMAGEM": "gerenciaenfermagem.hcid@gmail.com",
        "NIR (NÚCLEO INTERNO DE REGULAÇÃO)": "coordenanirs2@gmail.com",
        "CENTRO CIRÚRGICO": "centrocirurgicosoc2@gmail.com",
        "ENDOSCOPIA": "endoscopiasoc2@gmail.com",
        "DIREÇÃO ADMINISTRATIVA": "direcao.hospitaldacidade@gmail.com",
    }
    
    for chave, email in emails.items():
        if chave == nome or chave in nome or nome in chave:
            return email
    return "nspsoc2@gmail.com"

# =================================================================
# GESTORES
# =================================================================
def buscar_gestor(setor_notificado):
    if not setor_notificado or str(setor_notificado).strip() == "":
        return ""
    
    nome = str(setor_notificado).strip().upper()
    
    gestores = {
        "ALA D": "GESTOR DE ENFERMAGEM",
        "AGÊNCIA TRANSFUSIONAL": "GESTOR DA AGÊNCIA TRANSFUSIONAL",
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
        "TERAPIA OCUPACIONAL": "GESTOR DE TERAPIA OCUPACIONAL",
        "SERVIÇO SOCIAL": "GESTOR DE SERVIÇO SOCIAL",
        "NUTRIÇÃO": "GESTOR DE NUTRIÇÃO",
        "GESTOR DE ENFERMAGEM": "GESTOR DE ENFERMAGEM",
        "GESTOR DE GERÊNCIA DE ENFERMAGEM": "GESTOR DE GERÊNCIA DE ENFERMAGEM",
        "NIR (NÚCLEO INTERNO DE REGULAÇÃO)": "GESTOR DO NIR",
        "CENTRO CIRÚRGICO": "GESTOR DE ENFERMAGEM",
        "ENDOSCOPIA": "GESTOR DE ENDOSCOPIA",
        "DIREÇÃO ADMINISTRATIVA": "GESTOR DE DIREÇÃO",
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
        
        if 'Status' not in df.columns:
            df.insert(1, 'Status', 'PENDENTE')
        
        total = len(df)
        enviados = df['Status'].apply(lambda x: str(x).strip().upper() == "ENVIADO").sum()
        pendentes = total - enviados
        
        try:
            cols = list(df.columns)
            an_idx = cols.index('AN') if 'AN' in cols else -1
            ao_idx = cols.index('AO') if 'AO' in cols else -1
            ap_idx = cols.index('AP') if 'AP' in cols else -1
            if an_idx >= 0 and ao_idx >= 0 and ap_idx >= 0:
                an_val = pd.to_numeric(df.iloc[0, an_idx], errors='coerce')
                ao_val = pd.to_numeric(df.iloc[0, ao_idx], errors='coerce')
                ap_val = pd.to_numeric(df.iloc[0, ap_idx], errors='coerce')
                if pd.notna(an_val) and pd.notna(ao_val) and pd.notna(ap_val):
                    total = int(an_val)
                    enviados = int(ao_val)
                    pendentes = int(ap_val)
        except:
            pass
        
        st.subheader("RESUMO GERAL")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Memorandos", total)
        col2.metric("✅ Já Enviados", enviados)
        col3.metric("⏳ Pendentes", pendentes)
        st.info(f"{total} - {enviados} = {pendentes} memorandos pendentes")
        st.divider()
        
        data_atual = datetime.now()
        data_extenso = data_por_extenso(data_atual)
        output_planilha = io.BytesIO()
        
        aba_pendentes, aba_enviados = st.tabs(["⏳ PENDENTES", "✅ ENVIADOS"])
        todos_itens = []
        
        for idx, linha in df.iterrows():
            num_notificacao = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
            if not num_notificacao or num_notificacao.upper() == "Nº" or "NOTIF" in str(num_notificacao).upper():
                continue
            num_memorando = tratar_texto(linha.iloc[15]) if len(linha) > 15 else ""
            if not num_memorando:
                continue
            nome_paciente = tratar_texto(linha.iloc[9]) if len(linha) > 9 else ""
            status_atual = tratar_texto(linha.get('Status', '')).upper()
            enviado_bool = (status_atual == "ENVIADO")
            todos_itens.append({
                'idx': idx,
                'num_notificacao': num_notificacao,
                'num_memorando': num_memorando,
                'nome_paciente': nome_paciente,
                'enviado': enviado_bool,
                'linha': linha
            })
        
        # ABA PENDENTES
        with aba_pendentes:
            st.markdown("### 📋 Memorandos Pendentes")
            tem_pendentes = False
            for item in todos_itens:
                if item['enviado']:
                    continue
                tem_pendentes = True
                idx = item['idx']
                linha = item['linha']
                num_notificacao = item['num_notificacao']
                num_memorando = item['num_memorando']
                nome_paciente = item['nome_paciente']
                
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
                setor_notificado = tratar_texto(linha.iloc[18]) if len(linha) > 18 else ""
                
                m = "☑" if "MANHÃ" in turno.upper() else "☐"
                t = "☑" if "TARDE" in turno.upper() else "☐"
                n = "☑" if "NOITE" in turno.upper() else "☐"
                
                gestor = buscar_gestor(setor_notificado)
                email_destino = buscar_email_setor(setor_notificado)
                
                hora_atual = datetime.now().hour
                if 6 <= hora_atual < 12:
                    saudacao = "Bom dia,"
                elif 12 <= hora_atual < 18:
                    saudacao = "Boa tarde,"
                else:
                    saudacao = "Boa noite,"
                
                titulo_memo = f"📋 Notificação Nº {num_notificacao} — {nome_paciente}"
                with st.expander(titulo_memo, expanded=True):
                    marcado = st.checkbox(
                        "✅ Já enviei este memorando",
                        value=False,
                        key=f"check_pend_{num_notificacao}_{num_memorando}_{idx}"
                    )
                    if marcado:
                        df.at[idx, 'Status'] = 'ENVIADO'
                        st.success("✅ Movido para a aba ENVIADOS! Salve a planilha para manter a alteração.")
                    st.divider()
                    st.write(f"📅 Data de Emissão: {data_extenso}")
                    st.write(f"📅 Data da Ocorrência: {data_ocorrencia} | Notificação: {data_notificacao}")
                    st.write(f"📍 Local: {localizacao}")
                    st.write(f"🏛️ Setor Notificante: {setor_notificante}")
                    st.write(f"📤 Setor Notificado: {setor_notificado}")
                    st.write(f"👤 Gestor: {gestor}")
                    st.write(f"⚠️ Tipo: {tipo_incidente} | Classificação: {classificacao_incidente}")
                    st.write(f"👤 Paciente: {nome_paciente} | Leito: {leito}")
                    st.write(f"💡 Sugestão: {sugestao}")
                    st.write(f"📧 E-mail para envio: {email_destino}")
                    
                    corpo_email = f"""
{saudacao} Prezado(a),

Segue em anexo a Notificação Nº {num_notificacao} referente ao Memorando {num_memorando},
para ser analisada e respondida (via e-mail) em até 15 dias após a data presente.

ATENÇÃO: A resposta via e-mail deve conter um arquivo em formato Word ou PDF para arquivamento
das respostas conforme rotina institucional. Não serão aceitas mensagens sem arquivo anexo.

Atenciosamente,

Ezequias S. Santos
Agente Administrativo — NAQH / NSP
Núcleo de Segurança do Paciente
nspsoc2@gmail.com
                    """
                    with st.expander("Texto do E-mail"):
                        st.text(corpo_email.strip())
                    
                    dados_docx = {
                        "{{numero_memorando}}": num_memorando,
                        "{{notificacao_n}}": num_notificacao,
                        "{{gestor}}": gestor,
                        "{{setor_notificado}}": setor_notificado,
                        "{{setor_notificante}}": setor_notificante,
                        "{{data_envio}}": data_extenso,
                        "{{data_ocorrencia}}": data_ocorrencia,
                        "{{data_notificacao}}": data_notificacao,
                        "{{m}}": m,
                        "{{t}}": t,
                        "{{n}}": n,
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
                            file_name=f"NSP_{num_memorando.replace('/', '_')}.docx",
                            mime="application/vnd
