# =================================================================
# GERADOR DE MEMORANDOS — HOSPITAL DA CIDADE DR. JACKSON LAGO
# =================================================================

# PASSO 1 — IMPORTAR PRIMEIRO
import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document

# PASSO 2 — CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Emissor de Memorandos", page_icon="📄", layout="wide")

# PASSO 3 — CSS DE MELHORIAS DE NAVEGAÇÃO
st.markdown("""
<style>
/* ===== BARRA DE ROLAGEM MAIS VISÍVEL ===== */
::-webkit-scrollbar {
    width: 14px !important;
    height: 14px !important;
}
::-webkit-scrollbar-track {
    background: #2c3e50 !important;
    border-radius: 7px !important;
}
::-webkit-scrollbar-thumb {
    background: #3498db !important;
    border-radius: 7px !important;
    border: 2px solid #2c3e50 !important;
}
::-webkit-scrollbar-thumb:hover {
    background: #2980b9 !important;
}
/* ===== DESTAQUE NOS EXPANDERS ===== */
div[data-testid="stExpander"] {
    border: 1px solid #4a5568 !important;
    border-radius: 8px !important;
    margin-bottom: 12px !important;
    background: #1e293b !important;
}
div[data-testid="stExpander"]:hover {
    border-color: #3498db !important;
    box-shadow: 0 2px 8px rgba(52, 152, 219, 0.15) !important;
}
/* ===== TARJA ENVIADO ===== */
.tarja-enviado {
    background: linear-gradient(90deg, #00b894, #00a884) !important;
    color: white !important;
    padding: 10px 16px !important;
    border-radius: 6px !important;
    font-weight: bold !important;
    text-align: center !important;
    margin-bottom: 16px !important;
    font-size: 16px !important;
}
/* ===== BOTÕES ===== */
button[kind="primary"] {
    background: #3498db !important;
    border: none !important;
    font-weight: 500 !important;
}
button[kind="primary"]:hover {
    background: #2980b9 !important;
}
/* ===== LEITURA GERAL ===== */
.block-container {
    max-width: 100% !important;
    padding-top: 1.5rem !important;
    padding-bottom: 3rem !important;
}
hr {
    border: none !important;
    border-top: 1px solid #4a5568 !important;
    margin: 1.5rem 0 !important;
}
/* ===== SEÇÕES ===== */
.secao-enviados h2 {
    color: #00b894 !important;
    border-bottom: 2px solid #00b894;
    padding-bottom: 8px;
}
.secao-pendentes h2 {
    color: #f39c12 !important;
    border-bottom: 2px solid #f39c12;
    padding-bottom: 8px;
}
</style>
""", unsafe_allow_html=True)

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
# DADOS FIXOS
# =================================================================
TOTAL_GERAL = 964
JA_ENVIADOS_INICIAL = 433
PENDENTES_INICIAL = 531

# =================================================================
# INTERFACE PRINCIPAL
# =================================================================

st.title("📄 Emissor de Memorandos — Hospital da Cidade Dr. Jackson Lago")
st.divider()

arquivo_excel = st.file_uploader("Carregue a Planilha (.xlsx)", type=["xlsx"])

if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel, engine="openpyxl")
        
        if 'Status' not in df.columns:
            df.insert(1, 'Status', '')
        
        # Separar os itens
        itens_enviados = []
        itens_pendentes = []
        
        for idx, linha in df.iterrows():
            num_notificacao = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
            if not num_notificacao or num_notificacao.upper() == "Nº" or "NOTIF" in str(num_notificacao).upper():
                continue
            num_memorando = tratar_texto(linha.iloc[15]) if len(linha) > 15 else ""
            if not num_memorando:
                continue
            nome_paciente = tratar_texto(linha.iloc[9]) if len(linha) > 9 else ""
            status_atual = tratar_texto(linha.get('Status', '')).upper()
            
            item = {
                'idx': idx,
                'linha': linha,
                'num_notificacao': num_notificacao,
                'num_memorando': num_memorando,
                'nome_paciente': nome_paciente,
                'enviado': (status_atual == "ENVIADO")
            }
            
            if item['enviado']:
                itens_enviados.append(item)
            else:
                itens_pendentes.append(item)
        
        # Calcular totais
        marcados_agora = len(itens_enviados) - JA_ENVIADOS_INICIAL
        if marcados_agora < 0:
            marcados_agora = 0
        total_enviados = JA_ENVIADOS_INICIAL + marcados_agora
        total_pendentes = PENDENTES_INICIAL - marcados_agora
        if total_pendentes < 0:
            total_pendentes = 0
        
        # RESUMO
        st.subheader("📊 RESUMO GERAL")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Memorandos", TOTAL_GERAL)
        col2.metric("✅ Já Enviados", total_enviados)
        col3.metric("⏳ Pendentes", total_pendentes)
        st.info(f"{TOTAL_GERAL} - {total_enviados} = {total_pendentes} memorandos pendentes")
        st.divider()
        
        data_atual = datetime.now()
        data_extenso = data_por_extenso(data_atual)
        output_planilha = io.BytesIO()
        
        # FUNÇÃO PARA RENDERIZAR UM ITEM
        def exibir_item(item, eh_pendente=True):
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
            
            with st.expander(titulo_memo, expanded=eh_pendente):
                if not eh_pendente:
                    st.markdown("""
<div style="background: linear-gradient(90deg, #00b894, #00a884); color:white; padding:10px 16px; border-radius:6px; font-weight:bold; text-align:center; margin-bottom:16px; font-size:16px;">
✅ JÁ FOI ENVIADO
</div>
                    """, unsafe_allow_html=True)
                else:
                    marcado = st.checkbox(
                        "✅ Marcar como enviado",
                        value=False,
                        key=f"check_{num_notificacao}_{num_memorando}_{idx}"
                    )
                    if marcado:
                        df.at[idx, 'Status'] = 'ENVIADO'
                        st.rerun()
                
                st.divider()
                st.write(f"📤 **Setor Notificado:** {setor_notificado if setor_notificado else 'Não informado'}")
                st.write(f"📧 **E-mail para envio:** {email_destino}")
                st.divider()
                
                corpo_email = f"""{saudacao} Prezado(a),

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
                with st.expander("📋 Copiar texto do e-mail"):
                    st.code(corpo_email, language=None)
                
                dados_docx = {
                    "{{numero_memorando}}": num_memorando,
                    "{{notificacao_n}}": num_notificacao,
                    "{{gestor}}":
