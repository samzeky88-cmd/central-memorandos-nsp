# =================================================================
# GERADOR DE MEMORANDOS — HOSPITAL DA CIDADE DR. JACKSON LAGO
# =================================================================

import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document

# =================================================================
# CONFIGURAÇÕES FIXAS
# =================================================================
TOTAL_GERAL_MEMORANDOS = 964

# =================================================================
# LISTA COMPLETA DE E-MAILS
# =================================================================
LISTA_SETORES = [
    ("DIREÇÃO", "direcao.hospitaldacidade@gmail.com"),
    ("GABINETE DA DIREÇÃO", "gabinete.hospitaldacidade@gmail.com"),
    ("AGÊNCIA TRANSFUSIONAL", "agenciatransfusional.hcid@gmail.com"),
    ("ALA A - ENFERMAGEM", "alaaclinicamedica@gmail.com"),
    ("ALA B - ENFERMAGEM", "alab.hcid@gmail.com"),
    ("ALA C - ENFERMAGEM", "cirurgica.hcid@gmail.com"),
    ("ALA D - ENFERMAGEM", "ortopediaalad.hcid@gmail.com"),
    ("ALA G/H", "aladgh.hcid@gmail.com"),
    ("ALA F - ENFERMAGEM", "utihcid@gmail.com"),
    ("ALA I - ENFERMAGEM", "ortopediaalai.hcid@gmail.com"),
    ("ALA J - ENFERMAGEM", "ortopediaalaj.hcid@gmail.com"),
    ("ALA L - ENFERMAGEM", "alalortopedia.hcid@gmail.com"),
    ("ALMOXARIFADO", "almoxarifadosocorrao2@gmail.com"),
    ("AMBULATÓRIO", "ambulatorio.hcid@gmail.com"),
    ("ASSESSORIA JURÍDICA", "fernandoacmoraes.advogado@gmail.com"),
    ("CAC", "cac.socorrao2@gmail.com"),
    ("CENTRAL DE CÓPIAS", "centraldecopias.hcid@gmail.com"),
    ("CCIH", "ccih.hcid@gmail.com"),
    ("CENTRO CIRÚRGICO - ENFERMAGEM", "centrocirurgicosoc2@gmail.com"),
    ("COMISSÃO DE ÓBITOS", "comissaodeobito.hcid@gmail.com"),
    ("COMISSÃO DE ÉTICA EM ENFERMAGEM", "comissaodeetica.soc2@gmail.com"),
    ("COMISSÃO DE ÉTICA MÉDICA", "cmeesterilizacao@gmail.com"),
    ("CME", "cmeesterilizacao@gmail.com"),
    ("CENTRAL DE TRANSPORTE INTRA HOSPITALAR", "centraltransporte.hcid@gmail.com"),
    ("DOCUMENTOS INSTITUCIONAIS", "documentos.soc2@gmail.com"),
    ("NAQH", "naqhsoc2@gmail.com"),
    ("ECP", "carlosemilioecp@gmail.com"),
    ("ENDOSCOPIA", "endoscopiasoc2@gmail.com"),
    ("EDOT", "edot.hcid@gmail.com"),
    ("EMTN", "emtnhospitaldacidade@gmail.com"),
    ("ESCALAS", "escalas.hcid25@gmail.com"),
    ("FARMÁCIA", "farmacia.hcid@gmail.com"),
    ("FARMACOVIGILÂNCIA", "farmacovigilanciahcid@gmail.com"),
    ("FATURAMENTO", "contasmedicassocorrao2@gmail.com"),
    ("FISIOTERAPIA / PS - ENFERMARIAS", "fisioreabsoc2@gmail.com"),
    ("FISIOTERAPIA - UTI", "fisioterapiasocorrao@gmail.com"),
    ("FONOAUDIOLOGIA", "fonosocorrao2@yahoo.com"),
    ("GERÊNCIA DE ENFERMAGEM", "gerenciadeenf.hcid@gmail.com"),
    ("GERÊNCIA DA REABILITAÇÃO", "gerenciareab.hcid@gmail.com"),
    ("GERENCIAMENTO DE RESÍDUOS", "residuos.hcid@gmail.com"),
    ("HOTELARIA", "hotelaria.hcid@gmail.com"),
    ("LABORATÓRIO", "hcid.laboratorio@gmail.com"),
    ("MANUTENÇÃO", "manutencao.hcid@gmail.com"),
    ("MAQUEIROS", "centraltransporte.hcid@gmail.com"),
    ("NCE", "nce.socorrao2@gmail.com"),
    ("NEP", "nep.hcid@gmail.com"),
    ("NEPEX", "nepexhcid@gmail.com"),
    ("NIR", "coordenanirs2@gmail.com"),
    ("NSP", "nspsoc2@gmail.com"),
    ("NUTRIÇÃO", "nutricaohospitaldacidade@gmail.com"),
    ("NVE", "nhe.hcid@gmail.com"),
    ("NHE", "nhe.hcid@gmail.com"),
    ("ODONTOLOGIA", "robertapalacioferes@gmail.com"),
    ("OUVIDORIA", "ouvidoriahcid@gmail.com"),
    ("PARECER DE EXAMES", "coodsaet2026@gmail.com"),
    ("PATRIMÔNIO", "patrimonio.hcid@gmail.com"),
    ("PORTARIA", "portaria.hcid@gmail.com"),
    ("PSICOLOGIA", "psicologia.hcid@gmail.com"),
    ("RECURSOS HUMANOS", "setorpessoal.hcid@gmail.com"),
    ("ROUPARIA", "rouparia.hcid@gmail.com"),
    ("SAEEM", "saeem.hcid@gmail.com"),
    ("SAME", "same25.hcid@gmail.com"),
    ("SAET", "coodsaet2026@gmail.com"),
    ("SADT", "sadthosp@gmail.com"),
    ("SETOR DE IMAGEM", "sadt.hosp@gmail.com"),
    ("SALA VERMELHA", "salavermelhasdm@gmail.com"),
    ("SDM/VERMELHA - ENFERMAGEM", "salavermelhasdm@gmail.com"),
    ("SERVIÇO SOCIAL", "servicosocialhcid@gmail.com"),
    ("TECNOLOGIA DA INFORMAÇÃO", "suporteti.hcid@gmail.com"),
    ("TERAPIA OCUPACIONAL", "terapiaocupacional.hcid@gmail.com"),
    ("TRANSPORTE EXTERNO", "centraltransporte.hcid@gmail.com"),
    ("UNIDADE SENTINELA", "unidadesentinela@gmail.com"),
    ("UTI A - MÉDICOS", "pedrohfrota@gmail.com"),
    ("UTI B - MÉDICOS", "livia.goreth@ufma.br"),
    ("UTI C - MÉDICOS", "hugogama19@hotmail.com"),
    ("ORTOPEDIA - MÉDICOS", "abraaojunior@hotmail.com"),
    ("PRONTO SOCORRO - MÉDICOS", "gustavo1422@hotmail.com"),
    ("ALA A - MÉDICOS", "kledsonmarques10@gmail.com"),
    ("ALA B - MÉDICOS", "aliomarcamaradecarvalho@gmail.com"),
    ("CENTRO CIRÚRGICO - MÉDICOS", "felipeabreu805@gmail.com"),
    ("CARDIOLOGIA - MÉDICOS", "jordach.xerez@gmail.com"),
    ("ANESTESISTAS - MÉDICO", "rafaelcbarrosneto@yahoo.com.br"),
    ("XEROX/REPROGRAFIA", "centraldecopias.hcid@gmail.com"),
]

df_setores = pd.DataFrame(LISTA_SETORES, columns=["SETOR", "E-MAIL"])
df_setores = df_setores[df_setores["E-MAIL"].str.strip() != ""].reset_index(drop=True)

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
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    return f"{dt.day} de {meses[dt.month-1]} de {dt.year}"

def buscar_email_setor(nome_setor):
    if not nome_setor or str(nome_setor).strip() == "":
        return "nspsoc2@gmail.com"
    nome_upper = str(nome_setor).strip().upper()
    for _, linha in df_setores.iterrows():
        setor_upper = linha["SETOR"].upper()
        if nome_upper in setor_upper or setor_upper in nome_upper:
            return linha["E-MAIL"]
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
        st.error("Arquivo modelo_memorando.docx não encontrado!")
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
            for m_col in:
                if m_col < len(linha):
                    num = limpar_numero(linha.iloc[m_col])
                    if num and num != "" and "Nº MEMO" not in num:
                        qtd += 1
    return qtd

# =================================================================
# INTERFACE PRINCIPAL
# =================================================================
st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
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

        data_selecionada = st.date_input("Data de Envio", datetime.now())
        ano_atual = data_selecionada.year
        data_extenso = data_por_extenso(data_selecionada)

        qtd_processados = 0

        for idx, linha in df.iterrows():
            num_notif = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
            if not num_notif or num_notif.upper() == "STATUS" or "NOTIF" in str(num_notif).upper():
                continue

            nome_paciente = tratar_texto(linha.iloc[9]) if len(linha) > 9 else ""
            if not nome_paciente:
                continue

            status_registro = tratar_texto(linha.iloc[1]).upper() if len(linha) > 1 else "PENDENTE"
            data_ocorrencia = formatar_data_sem_hora(linha.iloc[2]) if len(linha) > 2 else ""
            data_notificacao = formatar_data_sem_hora(linha.iloc[3]) if len(linha) > 3 else ""
            turno = tratar_texto(linha.iloc[4]).upper() if len(linha) > 4 else ""
