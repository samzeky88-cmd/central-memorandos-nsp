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

# ------------------- CONFIGURAÇÕES -------------------
EMAIL_REMETENTE = "nspsoc2@gmail.com"
ANO = "2026"

# ------------------- LISTA DE E-MAILS DOS SETORES -------------------
lista_setores = [
    ("DIREÇÃO", "direcao.hospitaldacidade@gmail.com"),
    ("GABINETE DA DIREÇÃO", "gabinete.hospitaldacidade@gmail.com"),
    ("ALA A - ENFERMAGEM", "alaaclinicamedica@gmail.com"),
    ("ALA B - ENFERMAGEM", "alab.hcid@gmail.com"),
    ("ALA C - ENFERMAGEM", "cirurgica.hcid@gmail.com"),
    ("ALA D - ENFERMAGEM", "ortopediaalad.hcid@gmail.com"),
    ("ALA L - ENFERMAGEM", "alalortopedia.hcid@gmail.com"),
    ("CENTRO CIRÚRGICO", "centrocirurgicosoc2@gmail.com"),
    ("FISIOTERAPIA - ENFERMARIAS", "fisioreabsoc2@gmail.com"),
    ("FISIOTERAPIA - UTI", "fisioterapiasocorrao@gmail.com"),
    ("GERÊNCIA DE ENFERMAGEM", "gerenciadeenf.hcid@gmail.com"),
    ("ECP", "carlosemilioecp@gmail.com"),
    ("NSP", "nspsoc2@gmail.com"),
    ("FARMÁCIA", "farmacia.hcid@gmail.com"),
    ("SAET", "coodsaet2026@gmail.com"),
    ("SDM / SALA VERMELHA", "salavermelhasdm@gmail.com"),
    ("SERVIÇO SOCIAL", "servicosocialhcid@gmail.com"),
    ("NIR - NÚCLEO INTERNO DE REGULAÇÃO DE LEITOS", "coordenanirs2@gmail.com"),
    ("OUTRO", "")
]

# ------------------- CONTROLE DE PACIENTE E NÚMEROS -------------------
if "paciente_anterior" not in st.session_state:
    st.session_state.paciente_anterior = ""
if "ultimo_num_notif" not in st.session_state:
    st.session_state.ultimo_num_notif = ""
if "contador_memo" not in st.session_state:
    st.session_state.contador_memo = 1195  # ← PRÓXIMO NÚMERO DE MEMORANDO

# ------------------- PRÉ-VISUALIZAÇÃO PARA CONFERIR -------------------
def pre_visualizar_memorando(dados):
    if dados['turno'] == "MANHÃ":
        turno_texto = "( X ) MANHÃ (   ) TARDE (   ) NOITE"
    elif dados['turno'] == "TARDE":
        turno_texto = "(   ) MANHÃ ( X ) TARDE (   ) NOITE"
    else:
        turno_texto = "(   ) MANHÃ (   ) TARDE ( X ) NOITE"

    visual = f"""
📄 **PRÉ-VISUALIZAÇÃO DO MEMORANDO — CONFIRA ANTES DE ENVIAR**
---

**PREFEITURA DE SÃO LUÍS**
**SECRETARIA MUNICIPAL DE SAÚDE**
**HOSPITAL DA CIDADE DR. JACKSON LAGO**

**MEMO: Nº NSP {dados['memo_num']} / {ANO}**

**DE:** Coordenação do Núcleo de Segurança do Paciente do Hospital da Cidade Dr. Jackson Lago
**PARA:** {dados['destinatario']}
**ASSUNTO:** Nº {dados['notif_num']}

&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;**São Luís, {dados['data_envio']}**

Prezado (a), vimos através deste comunicar que recebemos uma notificação de incidente ocorrida neste setor. Segue abaixo as informações encaminhadas ao NSP:

• **DATA DA OCORRÊNCIA:** {dados['data_ocorrencia']}
• **DATA DA NOTIFICAÇÃO:** {dados['data_notif']}
• **TURNO QUE OCORREU INCIDENTE:** {turno_texto}
• **ONDE OCORREU INCIDENTE:** {dados['local']}
• **TIPO DE INCIDENTE:** {dados['tipo']}
• **CLASSIFICAÇÃO DO INCIDENTE:** {dados['classificacao']}
• **DESCRIÇÃO DA NOTIFICAÇÃO:** {dados['descricao']}
• **PACIENTE:** {dados['paciente']}
• **LEITO:** {dados['leito']}
• **SETOR NOTIFICANTE:** {dados['setor_origem']}

**SUGESTÃO:** {dados['sugestao']}

Conforme rotina institucional, o gestor tem o prazo de 15 dias para realizar comunicação do incidente com sua equipe e discutir barreiras para evitar a ocorrência de novos eventos.

Atenciosamente,

**FABRÍCIA ROCHA**
Coordenadora

Rua Tancredo Neves S/N – Santa Efigênia – CEP 65010-000, São Luís – MA
E-mail: nspsoc2@gmail.com | CNPJ: 02.930.277/0001-49

---
✅ **CONFERE TUDO?** Se estiver CERTO, clique em **GERAR E ENVIAR**!
"""
    return visual

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
    doc.add_paragraph("Existe um processo de trabalho relacionado ao incidente definido em POP, protocolo, norma, etc? Qual? As pessoas envolvidas no processo conhecem? O protocolo foi seguido durante o evento?")
    doc.add_paragraph("\n" + "_" * 80 + "\n" + "_" * 80 + "\n")
    doc.add_paragraph("Existe monitoramento/gerenciamento da utilização do POP, protocolo, norma, etc? Como é gerenciado e quais resultados?")
    doc.add_paragraph("\n" + "_" * 80 + "\n" + "_" * 80 + "\n" + "_" * 80 + "\n")

    p3 = doc.add_paragraph()
    p3.add_run("3ª ETAPA: IDENTIFICAÇÃO DAS CAUSAS\n").bold = True
    doc.add_paragraph("Quais causas foram identificadas?")
    doc.add_paragraph("\n" + "_" * 80 + "\n" + "_" * 80 + "\n" + "_" * 80 + "\n" + "_" * 80 + "\n")
    doc.add_paragraph("Qual(is) medida(s) será(ão) tomada(s) para evitar que o incidente ocorra novamente? Colocar ação, responsável pela execução e data planejada para realizar.")
    doc.add_paragraph()

    tabela = doc.add_table(rows=4, cols=4)
    tabela.style = "Table Grid"
    hdr = tabela.rows[0].cells
    hdr[0].text = "Ação"
    hdr[1].text = "Responsável"
    hdr[2].text = "Data de realização"
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

# ------------------- INTERFACE PRINCIPAL COM ASSINATURA -------------------
st.set_page_config(page_title="Emissor de Memorandos — NSP", layout="wide")

# ========== CABEÇALHO COM TÍTULO + BONEQUINHO E FRASE DO LADO ==========
cab_esq, cab_dir = st.columns([3, 2])
with cab_esq:
    st.title("📝 Emissor de Memorandos — NSP")
with cab_dir:
    st.markdown("""
    <div style="text-align: center; padding-top: 15px;">
        <span style="font-size: 38px;">👨‍💻</span>
        <p style="font-size: 16px; font-weight: 600; margin: 5px 0; color: #2c3e50;">Criando Soluções Automatizadas</p>
        <p style="font-size: 15px; font-style: italic; margin: 0; color: #34495e;">Ezequias S. Santos</p>
    </div>
    """, unsafe_allow_html=True)

st.divider()

# ==================== RESTO DO FORMULÁRIO ====================
st.header("📋 Dados da Notificação — Preencha")

col1, col2, col3 = st.columns(3)
with col1:
    paciente = st.text_input("Nome do Paciente")
    data_ocorrencia = st.text_input("Data da Ocorrência", value="25/06/2026")
    data_notif = st.text_input("Data da Notificação", value="25/06/2026")

with col2:
    if paciente and paciente.strip().upper() == st.session_state.paciente_anterior.strip().upper():
        st.info(f"🔄 **Mesmo Paciente → Mesmo Nº Notificação: {st.session_state.ultimo_num_notif}**")
        num_notif = st.session_state.ultimo_num_notif
    else:
        num_notif = st.text_input("Nº da Notificação", value="")
        if paciente and num_notif:
            st.session_state.paciente_anterior = paciente.strip().upper()
            st.session_state.ultimo_num_notif = num_notif

    data_envio_br = st.text_input("Data do Memorando (aparece à DIREITA)", value=datetime.now().strftime("%d/%m/%Y"))

with col3:
    local = st.text_input("Local / Setor da Ocorrência", value="Ala B")
    tipo = st.text_input("Tipo de Incidente", value="Lesão_por_pressão")
    classificacao = st.text_input("Classificação", value="Incidente com dano leve")

col_a, col_b = st.columns(2)
with col_a:
    turno = st.selectbox("Turno", ["MANHÃ", "TARDE", "NOITE"])
with col_b:
    leito = st.text_input("Leito", value="")

setor_origem = st.text_input("Setor Notificante", value="NSP")
descricao = st.text_area("Descrição da Notificação", value="")
sugestao = st.text_area("Sugestão / Orientação", value="Sugerimos analisar o incidente juntamente com a equipe assistencial e discutir propostas de cuidados e prevenção conforme Protocolo institucional de Prevenção de LP")

st.divider()
st.header("👤 Destinatários — Escolha na Lista")

qtd_dest = st.number_input("Quantos destinatários?", min_value=1, max_value=5, value=1)
destinatarios = []

for i in range(qtd_dest):
    st.subheader(f"Destinatário {i+1}")
    c1, c2, c3 = st.columns([2, 2, 3])

    with c1:
        nomes_setores = [s[0] for s in lista_setores]
        escolha = st.selectbox(f"Setor — Dest {i+1}", nomes_setores, key=f"set_{i}")

    email_dest = ""
    for nome, em in lista_setores:
        if nome == escolha:
            email_dest = em
            break

    with c2:
        email_final = st.text_input(f"E-mail", value=email_dest, key=f"em_{i}")

    with c3:
        st.session_state.contador_memo += 1
        num_memo_atual = st.session_state.contador_memo
        st.info(f"📄 **Memorando Nº: {num_memo_atual}/{ANO}** | Notificação: **{num_notif}**")

    destinatarios.append({
        "nome": escolha,
        "email": email_final,
        "num_memo": num_memo_atual
    })

st.divider()

# 👀 BOTÃO DE PRÉ-VISUALIZAÇÃO
if st.button("👀 MOSTRAR PRÉ-VISUALIZAÇÃO PARA CONFERIR", type="secondary"):
    if not paciente or not num_notif or not descricao:
        st.error("❌ Preencha Nome do Paciente, Nº Notificação e Descrição!")
    else:
        for dest in destinatarios:
            st.subheader(f"📋 Pré-visualização para: {dest['nome']}")
            dados_pre = {
                "memo_num": dest["num_memo"],
                "notif_num": num_notif,
                "paciente": paciente,
                "data_ocorrencia": data_ocorrencia,
                "data_notif": data_notif,
                "data_envio": data_envio_br,
                "turno": turno,
                "local": local,
                "tipo": tipo,
                "classificacao": classificacao,
                "descricao": descricao,
                "leito": leito,
                "setor_origem": setor_origem,
                "sugestao": sugestao,
                "destinatario": dest["nome"]
            }
            st.markdown(pre_visualizar_memorando(dados_pre))
            st.divider()

st.divider()

# ✅ BOTÃO FINAL — GERAR + BAIXAR + ENVIAR
if st.button("✅ TUDO CERTO — GERAR + BAIXAR + ENVIAR", type="primary"):
    if not paciente or not num_notif or not descricao:
        st.error("❌ Preencha Nome do Paciente, Nº Notificação e Descrição!")
        st.stop()

    for dest in destinatarios:
        st.subheader(f"➡️ {dest['nome']} — Memorando Nº {dest['num_memo']}/{ANO}")

        dados = {
            "memo_num": dest["num_memo"],
            "notif_num": num_notif,
            "paciente": paciente,
            "data_ocorrencia": data_ocorrencia,
            "data_notif": data_notif,
            "data_envio": data_envio_br,
            "turno": turno,
            "local": local,
            "tipo": tipo,
            "classificacao": classificacao,
            "descricao": descricao,
            "leito": leito,
            "setor_origem": setor_origem,
            "sugestao": sugestao,
            "destinatario": dest["nome"]
        }

        arq_memo = gerar_memorando_word(dados)
        arq_roteiro = gerar_roteiro_word(dados)

        corpo = f"""Boa Tarde/Pela Manhã,

Segue em anexo o Memorando Nº {dest['num_memo']}/{ANO} referente à Notificação Nº {num_notif}, acompanhado do Roteiro de Tratativa.

Favor preencher e devolver no prazo de 15 dias.

Atenciosamente,
Ezequias S. Santos
Agente Administrativo — NAQH & NSP"""

        assunto = f"Notificação Nº {num_notif} | Memorando Nº {dest['num_memo']}/{ANO}"

        dl1, dl2 = st.columns([1, 1])
        with dl1:
            st.download_button("📄 Baixar Memorando (WORD)", arq_memo,
                file_name=f"Memorando_NSP_{dest['num_memo']}_Notif_{num_notif}.docx", key=f"dw_{dest['num_memo']}")
        with dl2:
            st.download_button("📋 Baixar Roteiro (OFICIAL)", arq_roteiro,
                file_name=f"Roteiro_Tratativa_Notif_{num_notif}.docx", key=f"dr_{dest['num_memo']}")

        arq_memo.seek(0)
        arq_roteiro.seek(0)
        ok, msg = enviar_email(dest["email"], assunto, corpo, arq_memo, arq_roteiro, dest["num_memo"], num_notif)
        if ok:
            st.success(f"✅ **ENVIADO COM SUCESSO!** → {dest['nome']}")
        else:
            st.error(f"❌ Erro: {msg}")
        st.divider()

st.caption("👨‍💻 Criando Soluções Automatizadas — Ezequias S. Santos | Data à DIREITA | PARA e ASSUNTO em linhas separadas | Sem repetição | Conferir antes de enviar")
