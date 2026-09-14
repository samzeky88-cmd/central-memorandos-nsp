import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import io
import pandas as pd

# ------------------- CONFIGURAÇÕES -------------------
ANO = "2026"

# ------------------- LISTA DE E-MAILS DOS SETORES -------------------
lista_setores = {
    "DIREÇÃO": "direcao.hospitaldacidade@gmail.com",
    "GABINETE DA DIREÇÃO": "gabinete.hospitaldacidade@gmail.com",
    "ALA A": "alaaclinicamedica@gmail.com",
    "ALA B": "alab.hcid@gmail.com",
    "ALA C": "cirurgica.hcid@gmail.com",
    "ALA D": "ortopediaalad.hcid@gmail.com",
    "ALA L": "alalortopedia.hcid@gmail.com",
    "ALA J": "alaj.hcid@gmail.com",
    "ALA H": "alah.hcid@gmail.com",
    "ALA G": "alag.hcid@gmail.com",
    "ALA I": "alai.hcid@gmail.com",
    "UTI B": "uti.b.hcid@gmail.com",
    "ALA F (UTI B)": "utib.hcid@gmail.com",
    "ALA F (UTI C)": "utic.hcid@gmail.com",
    "CENTRO CIRÚRGICO": "centrocirurgicosoc2@gmail.com",
    "FISIOTERAPIA": "fisioterapiasocorrao@gmail.com",
    "FISIOTERAPIA - ENFERMARIAS": "fisioreabsoc2@gmail.com",
    "FISIOTERAPIA - UTI": "fisioterapiasocorrao@gmail.com",
    "GERÊNCIA DE ENFERMAGEM": "gerenciadeenf.hcid@gmail.com",
    "GESTOR DE ENFERMAGEM": "gerenciadeenf.hcid@gmail.com",
    "ECP": "carlosemilioecp@gmail.com",
    "NSP": "nspsoc2@gmail.com",
    "FARMÁCIA": "farmacia.hcid@gmail.com",
    "SAET": "coodsaet2026@gmail.com",
    "SDM": "salavermelhasdm@gmail.com",
    "SALA VERMELHA": "salavermelhasdm@gmail.com",
    "SERVIÇO SOCIAL": "servicosocialhcid@gmail.com",
    "NIR": "coordenanirs2@gmail.com",
    "NÚCLEO INTERNO DE REGULAÇÃO DE LEITOS": "coordenanirs2@gmail.com",
    "GESTOR DE NIR": "coordenanirs2@gmail.com",
    "HOTELARIA": "hotelaria.hcid@gmail.com",
    "GESTOR DE HOTELARIA": "hotelaria.hcid@gmail.com",
    "DIREÇÃO TÉCNICA": "direcaotecnica.hcid@gmail.com",
    "GESTOR DE DIREÇÃO TÉCNICA": "direcaotecnica.hcid@gmail.com",
    "DIREÇÃO ADMINISTRATIVA": "direcaoadministrativa.hcid@gmail.com",
    "ENDOSCOPIA": "endoscopia.hcid@gmail.com",
    "IMAGEM": "imagem.hcid@gmail.com",
    "AGÊNCIA TRANSFUSIONAL": "transfusional.hcid@gmail.com",
    "HEMODIÁLISE": "hemodialise.hcid@gmail.com",
    "OUVIDORIA": "ouvidoria.hcid@gmail.com",
    "EQUIPE MULTIPROFISSIONAL": "emtn.hcid@gmail.com"
}

# ------------------- BUSCA INTELIGENTE -------------------
def encontrar_email(setor_nome):
    if pd.isna(setor_nome) or not str(setor_nome).strip():
        return ""
    nome_limpo = str(setor_nome).strip().upper()
    if nome_limpo in lista_setores:
        return lista_setores[nome_limpo]
    for chave, email in lista_setores.items():
        chave_limpa = chave.upper()
        if nome_limpo in chave_limpa or chave_limpa in nome_limpa:
            return email
    return ""

# ------------------- CONTADOR -------------------
if "contador_memo" not in st.session_state:
    st.session_state.contador_memo = 1195

# ------------------- GERAR MEMORANDO -------------------
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

# ------------------- GERAR ROTEIRO -------------------
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

# ------------------- INTERFACE PRINCIPAL -------------------
st.set_page_config(page_title="Gerador de Memorandos — NSP", layout="wide")

# CABEÇALHO
cab_esq, cab_dir = st.columns([3, 2])
with cab_esq:
    st.markdown("""
    <h1 style="margin-top: 10px; margin-bottom: 0; font-size: 26px;">📄 Gerador de Memorandos — Hospital Dr. Jackson Lago</h1>
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

# AVISO IMPORTANTE
st.info("ℹ️ O sistema irá GERAR e BAIXAR os Memorandos e Roteiros em Word. Você anexa e envia manualmente pelo e-mail. ✅")
st.divider()

# DATA DE ENVIO
st.subheader("📅 Configuração da Data de Envio")
data_envio = st.date_input("Selecione a data que sairá no cabeçalho do Memorando:", value=datetime.now())
data_formatada = data_envio.strftime("%d/%m/%Y")
st.divider()

# UPLOAD DA PLANILHA
st.subheader("📊 Suba a planilha contendo os incidentes (.xlsx)")
arquivo_excel = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"], label_visibility="collapsed")

if arquivo_excel:
    df = pd.read_excel(arquivo_excel)
    st.success(f"✅ Planilha carregada com {len(df)} registro(s)!")
    st.divider()

    # PRÉ-VISUALIZAÇÃO
    st.subheader("📋 Pré-visualização dos dados")
    st.dataframe(df, use_container_width=True)
    st.divider()

    # TABELA DE CONFERÊNCIA (SÓ PARA VOCÊ VER)
    st.subheader("🔍 Conferência de E-mails (apenas consulta)")
    conferencia = []

    for idx, linha in df.iterrows():
        setor_nome = str(linha.get("SETOR NOTIFICADO", linha.get("Setor Notificado", ""))).strip()
        email_da_planilha = str(linha.get("EMAIL_SETOR", linha.get("Email Setor", linha.get("Email Destino", "")))).strip()

        if email_da_planilha and "@" in email_da_planilha:
            email_final = email_da_planilha
            status = "✅ Usado da planilha"
        else:
            email_final = encontrar_email(setor_nome)
            status = f"✅ Encontrado" if email_final else "⚠️ Sem cadastro"

        conferencia.append({
            "Linha": idx + 1,
            "Setor": setor_nome,
            "E-mail": email_final if email_final else "⚠️ Preencher manual",
            "Status": status
        })

    st.dataframe(pd.DataFrame(conferencia), use_container_width=True)
    st.info("💡 Use a coluna 'E-mail' acima para saber para qual endereço enviar manualmente.")
    st.divider()

    # BOTÃO DE GERAÇÃO
    if st.button("✅ GERAR TODOS OS MEMORANDOS E ROTEIROS", type="primary"):
        st.success("🔄 Gerando arquivos... Pode demorar um pouco com muitos registros!")
        
        for idx, linha in df.iterrows():
            st.session_state.contador_memo += 1
            num_memo_atual = linha.get("Nº Memo", st.session_state.contador_memo)
            setor_nome = str(linha.get("SETOR NOTIFICADO", linha.get("Setor Notificado", ""))).strip()
            
            # Descobre o e-mail só pra te mostrar
            email_da_planilha = str(linha.get("EMAIL_SETOR", linha.get("Email Setor", ""))).strip()
            if email_da_planilha and "@" in email_da_planilha:
                email_final = email_da_planilha
            else:
                email_final = encontrar_email(setor_nome)

            dados = {
                "memo_num": num_memo_atual,
                "notif_num": str(linha.get("Nº", linha.get("Nº Notificação", ""))),
                "paciente": str(linha.get("PACIENTE", "")),
                "data_ocorrencia": str(linha.get("DATA DA OCORRÊNCIA", "")),
                "data_notif": str(linha.get("DATA DA NOTIFICAÇÃO", "")),
                "data_envio": data_formatada,
                "turno": str(linha.get("TURNO QUE OCORREU INCIDENTE", "")).upper().strip(),
                "local": str(linha.get("ONDE OCORREU INCIDENTE", "")),
                "tipo": str(linha.get("TIPO DE INCIDENTE", "")),
                "classificacao": str(linha.get("CLASSIFICAÇÃO DO INCIDENTE", "")),
                "descricao": str(linha.get("DESCRIÇÃO DA NOTIFICAÇÃO", "")),
                "leito": str(linha.get("LEITO", "")),
                "setor_origem": str(linha.get("SETOR NOTIFICANTE", "NSP")),
                "sugestao": str(linha.get("SUGESTÃO", "Sugerimos analisar o incidente juntamente com a equipe assistencial e discutir propostas de cuidados e prevenção conforme protocolo.")),
                "destinatario": setor_nome
            }

            st.subheader(f"📄 Linha {idx+1} — Memo Nº {num_memo_atual} → {setor_nome}")
            if email_final:
                st.write(f"📧 Enviar para: **{email_final}**")
            else:
                st.warning("⚠️ E-mail não encontrado — preencher manualmente")

            # GERA OS DOIS ARQUIVOS
            arq_memo = gerar_memorando_word(dados)
            arq_roteiro = gerar_roteiro_word(dados)

            # BOTÕES DE DOWNLOAD
            dl1, dl2 = st.columns([1, 1])
            with dl1:
                st.download_button(f"📄 Baixar Memorando", arq_memo,
                    file_name=f"Memorando_NSP_{num_memo_atual}_Notif_{dados['notif_num']}.docx", key=f"dw_{idx}")
            with dl2:
                st.download_button(f"📋 Baixar Roteiro", arq_roteiro,
                    file_name=f"Roteiro_Tratativa_Notif_{dados['notif_num']}.docx", key=f"dr_{idx}")
            
            st.divider()

st.caption("👨‍💻 Criando Soluções Automatizadas — Ezequias S. Santos | Gera Word → Você envia manualmente pelo e-mail ✅")
