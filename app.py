import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
import io
import pandas as pd

# ------------------- CONFIGURAÇÕES -------------------
EMAIL_REMETENTE = "nspsoc2@gmail.com"
ANO = "2026"

# ------------------- LISTA DE E-MAILS DOS SETORES -------------------
lista_setores = {
    "DIREÇÃO": "direcao.hospitaldacidade@gmail.com",
    "GABINETE DA DIREÇÃO": "gabinete.hospitaldacidade@gmail.com",
    "ALA A - ENFERMAGEM": "alaaclinicamedica@gmail.com",
    "ALA B - ENFERMAGEM": "alab.hcid@gmail.com",
    "ALA C - ENFERMAGEM": "cirurgica.hcid@gmail.com",
    "ALA D - ENFERMAGEM": "ortopediaalad.hcid@gmail.com",
    "ALA L - ENFERMAGEM": "alalortopedia.hcid@gmail.com",
    "CENTRO CIRÚRGICO": "centrocirurgicosoc2@gmail.com",
    "FISIOTERAPIA - ENFERMARIAS": "fisioreabsoc2@gmail.com",
    "FISIOTERAPIA - UTI": "fisioterapiasocorrao@gmail.com",
    "GERÊNCIA DE ENFERMAGEM": "gerenciadeenf.hcid@gmail.com",
    "ECP": "carlosemilioecp@gmail.com",
    "NSP": "nspsoc2@gmail.com",
    "FARMÁCIA": "farmacia.hcid@gmail.com",
    "SAET": "coodsaet2026@gmail.com",
    "SDM / SALA VERMELHA": "salavermelhasdm@gmail.com",
    "SERVIÇO SOCIAL": "servicosocialhcid@gmail.com",
    "NIR - NÚCLEO INTERNO DE REGULAÇÃO DE LEITOS": "coordenanirs2@gmail.com"
}

# ------------------- CONTADORES -------------------
if "contador_memo" not in st.session_state:
    st.session_state.contador_memo = 1195  # ← PRÓXIMO NÚMERO DE MEMORANDO

# ------------------- GERAR MEMORANDO WORD — EXATAMENTE IGUAL AO MODELO -------------------
def gerar_memorando_word(dados):
    doc = Document()

    cab = doc.add_paragraph()
    cab.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cab.add_run("PREFEITURA DE SÃO LUÍS\nSECRETARIA MUNICIPAL DE SAÚDE\nHOSPITAL DA CIDADE DR. JACKSON LAGO")
    run.bold = True
    run.font.size = Pt(12)

    doc.add_paragraph()

    p_memo = doc.add_paragraph()
    run = p_memo.add_run(f'MEMO: Nº NSP {dados["memo_num"]} / {ANO}')
    run.bold = True

    doc.add_paragraph(f'DE: Coordenação do Núcleo de Segurança do Paciente do Hospital da Cidade Dr. Jackson Lago')
    doc.add_paragraph(f'PARA: {dados["destinatario"]}')
    p_assun = doc.add_paragraph()
    run = p_assun.add_run(f'ASSUNTO: Nº {dados["notif_num"]}')
    run.bold = True

    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_data.add_run(f'São Luís, {dados["data_envio"]}')

    doc.add_paragraph()

    doc.add_paragraph("Prezado (a), vimos através deste comunicar que recebemos uma notificação de incidente ocorrida neste setor. Segue abaixo as informações encaminhadas ao NSP:")

    if dados['turno'] == "MANHÃ":
        turno_texto = "( X ) MANHÃ (   ) TARDE (   ) NOITE"
    elif dados['turno'] == "TARDE":
        turno_texto = "(   ) MANHÃ ( X ) TARDE (   ) NOITE"
    else:
        turno_texto = "(   ) MANHÃ (   ) TARDE ( X ) NOITE"

    p = doc.add_paragraph()
    p.add_run(f"• DATA DA OCORRÊNCIA: {dados['data_ocorrencia']}\n")
    p.add_run(f"• DATA DA NOTIFICAÇÃO: {dados['data_notif']}\n")
    p.add_run(f"• TURNO QUE OCORREU INCIDENTE: {turno_texto}\n")
    p.add_run(f"• ONDE OCORREU INCIDENTE: {dados['local']}\n")
    p.add_run(f"• TIPO DE INCIDENTE: {dados['tipo']}\n")
    p.add_run(f"• CLASSIFICAÇÃO DO INCIDENTE: {dados['classificacao']}\n")
    p.add_run(f"• DESCRIÇÃO DA NOTIFICAÇÃO: {dados['descricao']}\n")
    p.add_run(f"• PACIENTE: {dados['paciente']}\n")
    p.add_run(f"• LEITO: {dados['leito']}\n")
    p.add_run(f"• SETOR NOTIFICANTE: {dados['setor_origem']}\n\n")
    p.add_run(f"SUGESTÃO: {dados['sugestao']}\n")

    doc.add_paragraph("Conforme rotina institucional, o gestor tem o prazo de 15 dias para realizar comunicação do incidente com sua equipe e discutir barreiras para evitar a ocorrência de novos eventos.")

    doc.add_paragraph("\nAtenciosamente,\n\nFABRÍCIA ROCHA\nCoordenadora")

    doc.add_paragraph("\nRua Tancredo Neves S/N – Santa Efigênia – CEP 65010-000, São Luís – MA")
    doc.add_paragraph("E-mail: nspsoc2@gmail.com")
    doc.add_paragraph("CNPJ: 02.930.277/0001-49")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------- GERAR ROTEIRO OFICIAL -------------------
def gerar_roteiro_word(dados):
    doc = Document()
    cab = doc.add_paragraph()
    cab.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cab.add_run("ANÁLISE DE INCIDENTE LEVE OU MODERADO\nGESTOR DE ÁREA\n(Preencher após análise em equipe)")
    run.bold = True
    run.font.size = Pt(12)

    doc.add_paragraph()
    doc.add_paragraph(f"NÚMERO DA NOTIFICAÇÃO: {dados['notif_num']}")
    doc.add_paragraph(f"PACIENTE: {dados['paciente']}")
    doc.add_paragraph(f"DATA DA ANÁLISE: {dados['data_envio']}")
    doc.add_paragraph("PRONTUÁRIO: ________________________")
    doc.add_paragraph("Equipe responsável pela investigação do evento: ____________________________________________________")
    doc.add_paragraph()

    p1 = doc.add_paragraph()
    p1.add_run("1ª ETAPA: DEFINIÇÃO DO INCIDENTE\n").bold = True
    doc.add_paragraph("Como ocorreu o incidente notificado? Existem registros e/ou informações relacionados ao incidente em prontuários, relatórios ou outras fontes? Citar as fontes onde estão os registros.")
    doc.add_paragraph("\n" + "_" * 80 + "\n" + "_" * 80 + "\n" + "_" * 80 + "\n")
    doc.add_paragraph("Quais áreas, serviços e equipamentos estão envolvidos no incidente?")
    doc.add_paragraph("\n" + "_" * 80 + "\n" + "_" * 80 + "\n" + "_" * 80 + "\n")
    doc.add_paragraph("Quais consequências do incidente?")
    doc.add_paragraph("\n" + "_" * 80 + "\n" + "_" * 80 + "\n" + "_" * 80 + "\n")

    p2 = doc.add_paragraph()
    p2.add_run("2ª ETAPA: ANÁLISE DO INCIDENTE\n").bold = True
    doc.add_paragraph("Existe um processo de trabalho definido em POP, protocolo, norma, etc? Qual? As pessoas envolvidas conhecem? O protocolo foi seguido?")
    doc.add_paragraph("\n" + "_" * 80 + "\n" + "_" * 80 + "\n")
    doc.add_paragraph("Existe monitoramento/gerenciamento da norma/protocolo? Como é gerenciado e quais resultados?")
    doc.add_paragraph("\n" + "_" * 80 + "\n" + "_" * 80 + "\n" + "_" * 80 + "\n")

    p3 = doc.add_paragraph()
    p3.add_run("3ª ETAPA: IDENTIFICAÇÃO DAS CAUSAS\n").bold = True
    doc.add_paragraph("Quais causas foram identificadas?")
    doc.add_paragraph("\n" + "_" * 80 + "\n" + "_" * 80 + "\n" + "_" * 80 + "\n" + "_" * 80 + "\n")
    doc.add_paragraph("Qual(is) medida(s) será(ão) adotada(s) para evitar repetição? Colocar ação, responsável e prazo.")
    doc.add_paragraph()

    tabela = doc.add_table(rows=4, cols=4)
    tabela.style = "Table Grid"
    hdr = tabela.rows[0].cells
    hdr[0].text = "Ação"
    hdr[1].text = "Responsável"
    hdr[2].text = "Prazo"
    hdr[3].text = "Assinatura"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# ------------------- ENVIAR E-MAIL -------------------
def enviar_email(dest_email, assunto, corpo, arq_memo, arq_roteiro, num_memo, num_notif):
    msg = MIMEMultipart()
    msg["From"] = EMAIL_REMETENTE
    msg["To"] = dest_email
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "plain", "utf-8"))

    anexo1 = MIMEApplication(arq_memo.read())
    anexo1.add_header("Content-Disposition", "attachment", filename=f"Memorando_NSP_{num_memo}_Notif_{num_notif}.docx")
    msg.attach(anexo1)

    anexo2 = MIMEApplication(arq_roteiro.read())
    anexo2.add_header("Content-Disposition", "attachment", filename=f"Roteiro_Tratativa_Notif_{num_notif}.docx")
    msg.attach(anexo2)

    arq_memo.seek(0)
    arq_roteiro.seek(0)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as serv:
        senha = st.secrets.get("senha_gmail", "")
        if not senha:
            return False, "Senha de app não configurada nos Segredos"
        serv.login(EMAIL_REMETENTE, senha)
        serv.send_message(msg)
    return True, "✅ ENVIADO COM SUCESSO!"

# ------------------- INTERFACE PRINCIPAL — BONEQUINHO GARANTIDO -------------------
st.set_page_config(page_title="Emissor de Memorandos Individuais — NSP", layout="wide")

# ========== CABEÇALHO COM TÍTULO + BONEQUINHO E FRASE DO LADO ==========
cab_esq, cab_dir = st.columns([3, 2])

with cab_esq:
    st.markdown("""
    <h1 style="margin-top: 10px; margin-bottom: 0; font-size: 26px;">📄 Emissor de Memorandos Individuais — Hospital Dr. Jackson Lago</h1>
    """, unsafe_allow_html=True)

with cab_dir:
    st.markdown("""
    <div style="text-align: center; padding-top: 5px;">
        <img src="https://cdn-icons-png.flaticon.com/512/1005/1005141.png" width="45">
        <p style="font-size: 16px; font-weight: 600; margin: 6px 0 3px 0; color: #f0f0f0;">Criando Soluções Automatizadas</p>
        <p style="font-size: 14px; font-style: italic; margin: 0; color: #d0d0d0;">Ezequias S. Santos</p>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# ========== CONFIGURAÇÃO DE DATA ==========
st.subheader("📅 Configuração da Data de Envio")
data_envio = st.date_input("Selecione a data que sairá no cabeçalho do Memorando:", value=datetime.now())
data_formatada = data_envio.strftime("%d/%m/%Y")

st.divider()

# ========== UPLOAD DA PLANILHA EXCEL ==========
st.subheader("📊 Suba a planilha contendo os incidentes (.xlsx)")
arquivo_excel = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"], label_visibility="collapsed")

if arquivo_excel:
    df = pd.read_excel(arquivo_excel)
    st.success(f"✅ Planilha carregada com {len(df)} registro(s)!")
    st.divider()

    st.subheader("📋 Pré-visualização dos dados")
    st.dataframe(df, use_container_width=True)
    st.divider()

    if st.button("✅ GERAR E ENVIAR TODOS OS MEMORANDOS", type="primary"):
        for idx, linha in df.iterrows():
            st.session_state.contador_memo += 1
            num_memo_atual = st.session_state.contador_memo

            # Identificar e-mail
            setor_destino = str(linha.get("Setor_Destino", "")).strip()
            email_planilha = str(linha.get("Email_Destino", "")).strip()

            if email_planilha and "@" in email_planilha:
                email_final = email_planilha
            elif setor_destino in lista_setores:
                email_final = lista_setores[setor_destino]
            else:
                st.warning(f"⚠️ Linha {idx+1}: E-mail não encontrado — pulando...")
                continue

            destinatario_nome = setor_destino if setor_destino else "Destinatário"

            dados = {
                "memo_num": num_memo_atual,
                "notif_num": str(linha.get("Numero_Notificacao", "")),
                "paciente": str(linha.get("Paciente", "")),
                "data_ocorrencia": str(linha.get("Data_Ocorrencia", "")),
                "data_notif": str(linha.get("Data_Notificacao", "")),
                "data_envio": data_formatada,
                "turno": str(linha.get("Turno", "")).upper(),
                "local": str(linha.get("Local", "")),
                "tipo": str(linha.get("Tipo_Incidente", "")),
                "classificacao": str(linha.get("Classificacao", "")),
                "descricao": str(linha.get("Descricao", "")),
                "leito": str(linha.get("Leito", "")),
                "setor_origem": str(linha.get("Setor_Origem", "NSP")),
                "sugestao": str(linha.get("Sugestao", "Sugerimos analisar o incidente juntamente com a equipe assistencial e discutir propostas de cuidados e prevenção conforme protocolo.")),
                "destinatario": destinatario_nome
            }

            st.subheader(f"➡️ Memorando Nº {num_memo_atual}/{ANO} — {destinatario_nome}")

            arq_memo = gerar_memorando_word(dados)
            arq_roteiro = gerar_roteiro_word(dados)

            corpo = f"""Boa Tarde/Pela Manhã,

Segue em anexo o Memorando Nº {num_memo_atual}/{ANO} referente à Notificação Nº {dados['notif_num']}, acompanhado do Roteiro de Tratativa.

Favor preencher e devolver no prazo de 15 dias.

Atenciosamente,
Ezequias S. Santos
Agente Administrativo — NAQH & NSP"""

            assunto = f"Notificação Nº {dados['notif_num']} | Memorando Nº {num_memo_atual}/{ANO}"

            dl1, dl2 = st.columns([1, 1])
            with dl1:
                st.download_button(f"📄 Baixar Memorando Nº {num_memo_atual}", arq_memo,
                    file_name=f"Memorando_NSP_{num_memo_atual}_Notif_{dados['notif_num']}.docx", key=f"dw_{idx}")
            with dl2:
                st.download_button(f"📋 Baixar Roteiro Nº {num_memo_atual}", arq_roteiro,
                    file_name=f"Roteiro_Tratativa_Notif_{dados['notif_num']}.docx", key=f"dr_{idx}")

            arq_memo.seek(0)
            arq_roteiro.seek(0)

            if email_final and "@" in email_final:
                ok, msg = enviar_email(email_final, assunto, corpo, arq_memo, arq_roteiro, num_memo_atual, dados['notif_num'])
                if ok:
                    st.success(f"✅ ENVIADO para {email_final}")
                else:
                    st.error(f"❌ {msg}")
            else:
                st.info(f"ℹ️ E-mail vazio — baixe o arquivo e envie manualmente")

            st.divider()

st.caption("👨‍💻 Criando Soluções Automatizadas — Ezequias S. Santos | Data à DIREITA | PARA e ASSUNTO em linhas separadas | Lê planilha Excel | E-mails da lista ou digitados na planilha")
