import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import io
import pandas as pd
import re

# ------------------- CONFIGURAÇÕES -------------------
ANO = "2026"
COLUNA_STATUS = "STATUS"

# ------------------- LISTA DE E-MAILS DOS SETORES -------------------
lista_setores = {
    "DIREÇÃO": "direcao.hospitaldacidade@gmail.com",
    "GABINETE DA DIREÇÃO": "gabinete.hospitaldacidade@gmail.com",
    "ALA A": "alaaclinicamedica@gmail.com",
    "ALA B": "alab.hcid@gmail.com",
    "ALA C": "alaclinica.hcid@gmail.com",
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
    "NÚCLEO DE SEGURANÇA DO PACIENTE": "nspsoc2@gmail.com",
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

# ------------------- BUSCA INTELIGENTE DE E-MAIL -------------------
def encontrar_email(setor_nome):
    if pd.isna(setor_nome) or not str(setor_nome).strip():
        return ""
    nome_limpo = str(setor_nome).strip().upper()
    if nome_limpo in lista_setores:
        return lista_setores[nome_limpo]
    for chave in lista_setores:
        chave_upper = chave.upper()
        if nome_limpo in chave_upper or chave_upper in nome_limpo:
            return lista_setores[chave]
    return ""

# ------------------- LIMPAR NÚMEROS — MAIS ROBUSTO -------------------
def limpar_numero_memo(texto):
    """Extrai o número mesmo com espaços ou caracteres estranhos"""
    if pd.isna(texto):
        return ""
    t = str(texto).strip()
    if not t or t.upper() == "NAN":
        return ""
    # Pega QUALQUER sequência de dígitos
    digitos = re.findall(r'\d+', t)
    if digitos:
        return digitos[0]  # Pega o PRIMEIRO número
    return ""

def limpar_numero_notif(valor):
    if pd.isna(valor):
        return ""
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    t = str(valor).strip()
    t = re.sub(r'\.0$', '', t)
    digitos = re.findall(r'\d+', t)
    if digitos:
        return digitos[0]
    return ""

# ------------------- FORMATAR DATA SEM HORÁRIO -------------------
def formatar_data(valor):
    if pd.isna(valor):
        return ""
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    t = str(valor).strip()
    if " " in t:
        t = t.split(" ")[0]
    m = re.match(r'(\d{4})-(\d{1,2})-(\d{1,2})', t)
    if m:
        return f"{m.group(3)}/{m.group(2)}/{m.group(1)}"
    return t

# ==================================================
# 📋 COLUNAS DA PLANILHA — ATÉ 5 MEMORANDOS
# ==================================================
MEMORANDOS = [
    {"memo": "Nº Memo 01", "setor": "SETOR NOTIFICADO", "resposta": "Resposta"},
    {"memo": "Nº Memo 02", "setor": "SETOR NOTIFICADO 02", "resposta": "Resposta MEMO 02"},
    {"memo": "Nº Memo 03", "setor": "SETOR NOTIFICADO 03", "resposta": "Resposta MEMO 03"},
    {"memo": "Nº Memo 04", "setor": "SETOR NOTIFICADO 04", "resposta": "Resposta MEMO 04"},
    {"memo": "Nº Memo 05", "setor": "SETOR NOTIFICADO 05", "resposta": "Resposta MEMO 05"}
]

# ✅ BASTA COLOCAR "-" NA COLUNA RESPOSTA = JÁ ENVIADO!
def tem_resposta(texto):
    if pd.isna(texto):
        return False
    t = str(texto).strip().upper()
    if t == "-" or t == "ENVIADO" or t == "SIM":
        return True
    if not t or t == "NAN" or t == "NÃO PREENCHER" or t == "NAO PREENCHER":
        return False
    return True

# ------------------- GERAR MEMORANDO WORD -------------------
def gerar_memorando_word(dados):
    doc = Document()
    cab = doc.add_paragraph()
    cab.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = cab.add_run("PREFEITURA DE SÃO LUÍS SECRETARIA MUNICIPAL DE SAÚDE HOSPITAL DA CIDADE DR. JACKSON LAGO")
    run.bold = True
    run.font.size = Pt(12)

    p_memo = doc.add_paragraph()
    p_memo.space_before = Pt(12)
    run = p_memo.add_run(f'MEMO: Nº NSP {dados["memo_num"]} / {ANO}')
    run.bold = True
    run.font.size = Pt(12)

    doc.add_paragraph(f'DE: Coordenação do Núcleo de Segurança do Paciente do Hospital da Cidade Dr. Jackson Lago')
    doc.add_paragraph(f'PARA: {dados["destinatario"]}')

    p_assun = doc.add_paragraph()
    run = p_assun.add_run(f'ASSUNTO: Nº {dados["notif_num"]}')
    run.bold = True

    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_data.space_before = Pt(-16)
    p_data.add_run(f'São Luís, {dados["data_envio"]}')

    doc.add_paragraph()
    doc.add_paragraph("Prezado (a), vimos através deste comunicar que recebemos uma notificação de incidente ocorrida neste setor. Segue abaixo as informações encaminhadas ao NSP:")
    doc.add_paragraph()

    turno = dados['turno']
    if turno == "MANHÃ":
        turno_texto = "( X ) MANHÃ (   ) TARDE (   ) NOITE"
    elif turno == "TARDE":
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
    p.add_run(f"• SETOR NOTIFICANTE: {dados['setor_origem']}")

    doc.add_paragraph()
    doc.add_paragraph(f"SUGESTÃO: {dados['sugestao']}")
    doc.add_paragraph()
    doc.add_paragraph("Conforme rotina institucional, o gestor tem o prazo de 15 dias para realizar comunicação do incidente com sua equipe e discutir barreiras para evitar a ocorrência de novos eventos.")
    doc.add_paragraph()
    doc.add_paragraph("Atenciosamente,")
    doc.add_paragraph()
    doc.add_paragraph("FABRÍCIA ROCHA")
    doc.add_paragraph("Coordenadora")
    doc.add_paragraph()
    doc.add_paragraph("Rua Tancredo Neves S/N – Santa Efigênia – CEP 65010-000, São Luís – MA")
    doc.add_paragraph("E-mail: nspsoc2@gmail.com")
    doc.add_paragraph("CNPJ: 02.930.277/0001-49")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

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

# ✅ INSTRUÇÃO IMPORTANTE
st.info("ℹ️ ✅ DIAGNÓSTICO ATIVADO! Vai mostrar cada linha lida e o motivo se não aparecer!")
st.divider()

st.subheader("📅 Configuração da Data de Envio")
data_envio = st.date_input("Selecione a data que sairá no cabeçalho do Memorando:", value=datetime.now())
data_formatada = data_envio.strftime("%d/%m/%Y")
st.divider()

st.subheader("📊 Suba a planilha contendo os incidentes (.xlsx)")
arquivo_excel = st.file_uploader("Selecione o arquivo Excel", type=["xlsx"], label_visibility="collapsed")

if arquivo_excel:
    df = pd.read_excel(arquivo_excel)
    st.success(f"✅ Planilha carregada com {len(df)} linha(s)!")
    st.info(f"📋 Colunas encontradas: {', '.join(list(df.columns))}")
    st.divider()

    st.subheader("👀 Pré-visualização da planilha completa")
    st.dataframe(df, use_container_width=True)
    st.divider()

    if st.button("✅ GERAR TODOS COM DIAGNÓSTICO", type="primary"):
        st.success("🔄 Lendo TODAS as linhas...")
        qtd_gerados = 0
        qtd_enviados = 0
        qtd_nao_enviar = 0
        linhas_nao_processadas = []

        for idx, linha in df.iterrows():
            status_linha = str(linha.get(COLUNA_STATUS, "")).strip().upper()
            eh_nao_enviar = "NÃO ENVIAR" in status_linha or "NAO ENVIAR" in status_linha
            paciente = str(linha.get("PACIENTE", linha.get("Paciente", "Não informado"))).strip()

            val1 = linha.get("Nº")
            val2 = linha.get("Nº Notificação")
            num_notif = limpar_numero_notif(val1 if pd.notna(val1) else val2)

            encontrou_memo_na_linha = False

            for cfg in MEMORANDOS:
                memo_texto = str(linha.get(cfg["memo"], "")).strip()
                setor_nome = str(linha.get(cfg["setor"], "")).strip()
                resposta_coluna = linha.get(cfg["resposta"])

                # ✅ DIAGNÓSTICO — mostra o que está lendo
                num_memo_atual = limpar_numero_memo(memo_texto)

                if not num_memo_atual:
                    if memo_texto and memo_texto.upper() != "NAN":
                        linhas_nao_processadas.append(f"Linha {idx+2}: '{memo_texto}' → número não extraído")
                    continue

                encontrou_memo_na_linha = True
                memo_ja_enviado = tem_resposta(resposta_coluna)

                # ⛔ NÃO ENVIAR
                if eh_nao_enviar:
                    qtd_nao_enviar += 1
                    st.markdown(f"""
                    <div style="opacity:0.6; padding:12px; border:2px solid #ccc; border-radius:8px; background:#f8f8f8;">
                        <h3 style="color:#888; margin:0;">📄 Nº {num_memo_atual} | {paciente} → {setor_nome}</h3>
                        <p style="color:#888; font-weight:bold; margin:8px 0 0 0;">━━━━━━━━━━━━━━━ ⛔ NÃO ENVIAR ━━━━━━━━━━━━━━━━</p>
                    </div>
                    """, unsafe_allow_html=True)
                    st.divider()
                    continue

                # ✅ JÁ ENVIADO → mostra com tarja
                if memo_ja_enviado:
                    qtd_enviados += 1
                    st.markdown(f"""
                    <div style="opacity:0.5; padding:12px; border:2px solid #d4af37; border-radius:8px; background:#fff9e6;">
                        <h3 style="text-decoration: line-through; color:#999; margin:0;">📄 Nº {num_memo_atual} | {paciente} → {setor_nome}</h3>
                        <p style="color:#b8860b; font-weight:bold; margin:8px 0 0 0;">━━━━━━━━━━━━━━━━ ✅ JÁ ENVIADO ━━━━━━━━━━━━━━━━</p>
                    </div>
                    """, unsafe_allow_html=True)
                    st.divider()
                    continue

                # ✅ PENDENTE → gera com botão
                email_final = encontrar_email(setor_nome)

                dados = {
                    "memo_num": num_memo_atual,
                    "notif_num": num_notif,
                    "paciente": paciente,
                    "data_ocorrencia": formatar_data(linha.get("DATA DA OCORRÊNCIA", "")),
                    "data_notif": formatar_data(linha.get("DATA DA NOTIFICAÇÃO", "")),
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

                st.subheader(f"📄 Nº {num_memo_atual} / {ANO} | 👤 {paciente} → {setor_nome}")

                if email_final:
                    st.success(f"📧 **E-MAIL PRA COPIAR:** `{email_final}`")
                else:
                    st.warning("⚠️ E-mail não encontrado — preencher manualmente")

                st.markdown("### 📋 TEXTO DO CORPO DO E-MAIL — COPIE ABAIXO:")
                texto_email = f"""Boa Tarde, Prezados, ou Bom dia!

Segue em Anexo o Memorando Nº {num_memo_atual}/ {ANO} para ser analisado e respondido (via e-mail) em até 15 dias após a data presente.

**ATENÇÃO:** A resposta via e-mail deve constar um arquivo em forma de Word ou PDF para arquivamento de respostas conforme rotina institucional. Não serão aceitas mensagens via e-mail sem arquivo como resposta.

Segue abaixo a notificação para análise do incidente em equipe e resposta ao NSP:

• Memorando: Nº {num_memo_atual}/ {ANO}
• Notificação: Nº {num_notif}

Atenciosamente,

**Ezequias S. Santos**
Agente Administrativo - NAQH & NSP
"""
                st.code(texto_email, language=None)

                arq_memo = gerar_memorando_word(dados)
                arq_roteiro = gerar_roteiro_word(dados)

                dl1, dl2 = st.columns([1, 1])
                with dl1:
                    st.download_button(f"📄 Baixar Memorando", arq_memo,
                        file_name=f"Memorando_NSP_{num_memo_atual}_Notif_{num_notif}.docx", key=f"dw_{idx}_{num_memo_atual}")
                with dl2:
                    st.download_button(f"📋 Baixar Roteiro", arq_roteiro,
                        file_name=f"Roteiro_Tratativa_Notif_{num_notif}.docx", key=f"dr_{idx}_{num_memo_atual}")

                st.divider()
                qtd_gerados += 1

            if not encontrou_memo_na_linha:
                linhas_nao_processadas.append(f"Linha {idx+2}: SEM número de memorando detectado")

        # ✅ RESUMO FINAL
        st.success(f"✅ **PROCESSO CONCLUÍDO!**")
        st.markdown(f"""
        - 📋 **{qtd_gerados} memorandos gerados (PENDENTES)**
        - ✅ **{qtd_enviados} já enviados** — tarja riscada
        - ⛔ **{qtd_nao_enviar} marcados como NÃO ENVIAR**
        """)

        if linhas_nao_processadas:
            st.warning("⚠️ **LINHAS NÃO PROCESSADAS — VERIFIQUE ABAIXO:**")
            for aviso in linhas_nao_processadas:
                st.write(f"  → {aviso}")

st.caption("👨‍💻 Diagnóstico ativado: se faltar algum número, aparece o motivo! ✅")
