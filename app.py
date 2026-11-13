# =================================================================
# 📄 GERADOR DE MEMORANDOS — HOSPITAL DA CIDADE
# Versão Atualizada: Visualiza • Converte PDF • Envia por E-mail
# =================================================================

import os
import pandas as pd
import streamlit as st
import io
import urllib.parse
from datetime import datetime, timedelta, timezone
from docx import Document
import tempfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# =================================================================
# 📚 LISTA DE E-MAILS DOS SETORES
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
    ("COMISSÃO DE ÓBITOS", "comissaodeobito.hdacidade@gmail.com"),
    ("COMISSÃO DE ÉTICA EM ENFERMAGEM", "comissaodeetica.soc2@gmail.com"),
    ("CME", "cmeesterilizacao@gmail.com"),
    ("DOCUMENTOS INSTITUCIONAIS (NAQH)", "documentos.soc2@gmail.com"),
    ("ECP", "carlosemilioecp@gmail.com"),
    ("ENDOSCOPIA", "endoscopiasoc2@gmail.com"),
    ("EDOT", "edot.hcid@gmail.com"),
    ("EMTN", "emtnhospitaldacidade@gmail.com"),
    ("ESCALAS", "escalahcid25@gmail.com"),
    ("FARMÁCIA", "farmacia.hcid@gmail.com"),
    ("FARMACOVIGILÂNCIA", "farmacovigilanciahcid@gmail.com"),
    ("FATURAMENTO", "contasmedicassocorrao2@gmail.com"),
    ("FISIOTERAPIA /PS - ENFERMARIAS", "fisioreabsoc2@gmail.com"),
    ("FISIOTERAPIA - UTI", "fisioterapiasocorrao@gmail.com"),
    ("FONOAUDIOLOGIA", "fonosocorrao2@yahoo.com"),
    ("GERÊNCIA DE ENFERMAGEM", "gerenciadeenf.hcid@gmail.com"),
    ("GERÊNCIA DA REABILITAÇÃO", "gerenciareab.hcid@gmail.com"),
    ("GERENCIAMENTO DE RESÍDUOS", "residuos.hcid@gmail.com"),
    ("HOTELARIA", "hotelaria.hcid@gmail.com"),
    ("LABORATÓRIO", "hcid.laboratorio@gmail.com"),
    ("MANUTENÇÃO", "manutencao.hcid@gmail.com"),
    ("MAQUEIROS", "centraltransporte.hcid@gmail.com"),
    ("NAQH", "naqhsoc2@gmail.com"),
    ("NCE", "nce.socorrao2@gmail.com"),
    ("NEP", "nep.hcid@gmail.com"),
    ("NEPEX", "nepexhcid@gmail.com"),
    ("NSP", "nspsoc2@gmail.com"),
    ("NIR", "coordenanirs2@gmail.com"),
    ("NUTRIÇÃO", "nutricaohospitaldacidade@gmail.com"),
    ("NVE", "nhe.hcid@gmail.com"),
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
    ("SDM/VERMELHA - ENFERMAGEM", "salavermelhasdm@gmail.com"),
    ("SERVIÇO SOCIAL", "servicosocialhcid@gmail.com"),
    ("SETOR DE IMAGEM", "sadt.hosp@gmail.com"),
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
    ("XEROX/REPROGRAFIA", "centraldecopias.hcid@gmail.com")
]

df_setores = pd.DataFrame(LISTA_SETORES, columns=["SETOR", "E-MAIL"])
df_setores = df_setores[df_setores["E-MAIL"].str.strip() != ""].reset_index(drop=True)

# =================================================================
# 🔐 CONFIGURAÇÕES DE E-MAIL — Usamos Secrets do Streamlit
# =================================================================
CONFIG_EMAIL = {
    "remetente": st.secrets.get("EMAIL_REMETENTE", os.getenv("EMAIL_REMETENTE", "")),
    "senha_app": st.secrets.get("EMAIL_SENHA_APP", os.getenv("EMAIL_SENHA_APP", "")),
    "smtp_servidor": "smtp.gmail.com",
    "smtp_porta": 587,
    "assinatura": "Ezequias S. Santos\nAgente Administrativo - NAQH & NSP",
    "caminho_roteiro_fixo": "ROTEIRO_COORDENADORES.pdf"
}

st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
st.title("📝 Emissor de Memorandos Individuais - Hospital Dr. Jackson Lago")
st.markdown("### 📅 Configuração da Data de Envio")
data_selecionada = st.date_input("Selecione a data que sairá no cabeçalho do Memorando:", datetime.now())
arquivo_excel = st.file_uploader("Suba a planilha contendo os incidentes (.xlsx)", type=["xlsx"])
caminho_modelo = "modelo_memorando.docx"

# =================================================================
# 🧠 FUNÇÕES AUXILIARES
# =================================================================
def substituir_texto_protegendo_logos(doc, dicionario_tags):
    """Substitui o texto alterando apenas os 'runs' para proteger imagens e cabeçalhos."""
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
    try:
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan":
            return ""
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except:
        return str(valor).strip()

def limpar_numero_float(valor):
    if pd.isna(valor) or str(valor).strip().lower() == "nan":
        return ""
    try:
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        v_str = str(valor).strip()
        if v_str.endswith(".0"):
            return v_str[:-2]
        return v_str
    except:
        return str(valor)

def tratar_str_limpa(valor):
    if pd.isna(valor) or str(valor).strip().lower() == "nan" or str(valor).strip().lower() == "none":
        return ""
    return str(valor).strip()

def obter_data_por_extenso(dt):
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return f"{dt.day} de {meses[dt.month]} de {dt.year}"

def obter_saudacao():
    fuso_brasilia = timezone(timedelta(hours=-3))
    hora = datetime.now(fuso_brasilia).hour
    if hora < 12:
        return "Bom Dia Prezados"
    elif hora < 18:
        return "Boa Tarde Prezados"
    else:
        return "Boa Noite Prezados"

def converter_docx_para_pdf(docx_bytes):
    """Converte DOCX para PDF usando docx2pdf"""
    try:
        from docx2pdf import convert
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
            tmp_docx.write(docx_bytes)
            tmp_docx_path = tmp_docx.name
        tmp_pdf_path = tmp_docx_path.replace(".docx", ".pdf")
        convert(tmp_docx_path, tmp_pdf_path)
        with open(tmp_pdf_path, "rb") as f:
            pdf_bytes = f.read()
        os.unlink(tmp_docx_path)
        os.unlink(tmp_pdf_path)
        return pdf_bytes, None
    except Exception as e:
        return None, str(e)

def carregar_roteiro_fixo():
    """Carrega o arquivo fixo do Roteiro dos Coordenadores"""
    caminho = CONFIG_EMAIL["caminho_roteiro_fixo"]
    if os.path.exists(caminho):
        with open(caminho, "rb") as f:
            return f.read()
    return None

def enviar_email(destinatario, assunto, corpo, anexos):
    """
    Envia e-mail com anexos
    anexos = [(nome_arquivo, bytes, tipo_mime), ...]
    """
    if not CONFIG_EMAIL["remetente"] or not CONFIG_EMAIL["senha_app"]:
        return False, "⚠️ Credenciais de e-mail não configuradas nas Secrets do Streamlit!"
    try:
        msg = MIMEMultipart()
        msg["From"] = CONFIG_EMAIL["remetente"]
        msg["To"] = destinatario
        msg["Subject"] = assunto
        msg.attach(MIMEText(corpo, "plain", "utf-8"))
        for nome_arquivo, conteudo, tipo in anexos:
            part = MIMEBase(*tipo.split("/"))
            part.set_payload(conteudo)
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename=\"{nome_arquivo}\"")
            msg.attach(part)
        servidor = smtplib.SMTP(CONFIG_EMAIL["smtp_servidor"], CONFIG_EMAIL["smtp_porta"])
        servidor.starttls()
        servidor.login(CONFIG_EMAIL["remetente"], CONFIG_EMAIL["senha_app"])
        servidor.sendmail(CONFIG_EMAIL["remetente"], destinatario, msg.as_string())
        servidor.quit()
        return True, "✅ E-mail enviado com sucesso!"
    except Exception as e:
        return False, f"❌ Erro ao enviar: {str(e)}"

# =================================================================
# 📄 RENDERIZAÇÃO DE CADA LINHA/MEMORANDO
# =================================================================
@st.fragment
def renderizar_linha_paciente_sob_demanda(index, linha, num_colunas, data_extenso_envio):
    num_notif = limpar_numero_float(linha.iloc[0]) if num_colunas > 0 else "S-N"
    if num_notif.upper() == "STATUS" or "NOTIF" in num_notif.upper() or num_notif == "" or num_notif == "1":
        return
    status_envio = str(linha.iloc[1]).strip().upper() if num_colunas > 1 else ""
    dt_ocorr = formatar_data_br(linha.iloc[2]) if num_colunas > 2 else ""
    dt_notif = formatar_data_br(linha.iloc[3]) if num_colunas > 3 else ""
    turno_planilha = str(linha.iloc[4]).strip().upper() if num_colunas > 4 else ""
    onde_ocorreu = tratar_str_limpa(linha.iloc[5]) if num_colunas > 5 else "Ala B"
    tipo_incidente = tratar_str_limpa(linha.iloc[6]) if num_colunas > 6 else ""
    classificacao_incidente = tratar_str_limpa(linha.iloc[7]) if num_colunas > 7 else ""
    descricao_notificacao = tratar_str_limpa(linha.iloc[8]) if num_colunas > 8 else ""
    nome_do_paciente = tratar_str_limpa(linha.iloc[9]) if num_colunas > 9 else "Paciente Não Identificado"
    if nome_do_paciente.upper() == "PACIENTE":
        return
    leito_paciente = limpar_numero_float(linha.iloc[10]) if num_colunas > 10 else ""
    setor_notificante_bruto = tratar_str_limpa(linha.iloc[11]) if num_colunas > 11 else ""
    sugestao_nsp = tratar_str_limpa(linha.iloc[12]) if num_colunas > 12 else ""
    gestor_destinatario = tratar_str_limpa(linha.iloc[13]) if num_colunas > 13 else "GESTOR DE ENFERMAGEM"
    setor_notificado_final = tratar_str_limpa(linha.iloc[14]) if num_colunas > 14 else onde_ocorreu
    num_memo_cru = tratar_str_limpa(linha.iloc[15]) if num_colunas > 15 else ""
    if num_memo_cru == "" or num_memo_cru.upper() == "N° MEMO 01":
        num_memo_cru = "S-N"
    email_destino_planilha = tratar_str_limpa(linha.iloc[21]) if num_colunas > 21 else ""
    if email_destino_planilha.upper() == "EMAIL_SETOR":
        return
    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("No", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
    nome_base_arquivo = f"MEMORANDO N° {num_memo_limpo}_NOTIFICAÇÃO_N° {num_notif}_I_NSP"
    marca_manha = "X" if "MANH" in turno_planilha else ""
    marca_tarde = "X" if "TARD" in turno_planilha else ""
    marca_noite = "X" if "NOIT" in turno_planilha else ""
    if setor_notificante_bruto == "":
        setor_notificante_bruto = "NSP - NÚCLEO DE SEGURANÇA DO PACIENTE"
    
    dados_memorando = {
        "{{numero_memorando}}": num_memo_cru,
        "{{gestor}}": gestor_destinatario,
        "{{setor}}": setor_notificado_final,
        "{{notificacao_n}}": num_notif,
        "{{data_notificacao}}": dt_notif,
        "{{data_ocorrencia}}": dt_ocorr,
        "{{localizacao}}": onde_ocorreu,
        "{{tipo_incidente}}": tipo_inc
