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
    
    emails = {
        "DIREÇÃO": "direcao.hospitaldacidade@gmail.com",
        "GABINETE DA DIREÇÃO": "gabinete.hospitaldacidade@gmail.com",
        "AGÊNCIA TRANSFUSIONAL": "agenciatransfusional.hcid@gmail.com",
        "ALA A - ENFERMAGEM": "alaaclinicamedica@gmail.com",
        "ALA B - ENFERMAGEM": "alab.hcid@gmail.com",
        "ALA C - ENFERMAGEM": "cirurgica.hcid@gmail.com",
        "ALA D - ENFERMAGEM": "ortopediaalad.hcid@gmail.com",
        "ALA G/H": "aladgh.hcid@gmail.com",
        "ALA F - ENFERMAGEM": "utihcid@gmail.com",
        "ALA I - ENFERMAGEM": "ortopediaalai.hcid@gmail.com",
        "ALA J - ENFERMAGEM": "ortopediaalaj.hcid@gmail.com",
        "ALA L - ENFERMAGEM": "alalortopedia.hcid@gmail.com",
        "ALMOXARIFADO": "almoxarifadosocorrao2@gmail.com",
        "AMBULATÓRIO": "ambulatorio.hcid@gmail.com",
        "ASSESSORIA JURÍDICA": "fernandoacmoraes.advogado@gmail.com",
        "CAC": "cac.socorrao2@gmail.com",
        "CENTRAL DE CÓPIAS": "centraldecopias.hcid@gmail.com",
        "CCIH": "ccih.hcid@gmail.com",
        "CENTRO CIRÚRGICO - ENFERMAGEM": "centrocirurgicosoc2@gmail.com",
        "COMISSÃO DE ÓBITOS": "comissaodeobito.hcid@gmail.com",
        "COMISSÃO DE ÉTICA EM ENFERMAGEM": "comissaodeetica.soc2@gmail.com",
        "CME": "cmeesterilizacao@gmail.com",
        "DOCUMENTOS INSTITUCIONAIS": "documentos.soc2@gmail.com",
        "ECP": "carlosemilioecp@gmail.com",
        "ENDOSCOPIA": "endoscopiasoc2@gmail.com",
        "EDOT": "edot.hcid@gmail.com",
        "EMTN": "emtnhospitaldacidade@gmail.com",
        "ESCALAS": "escalahcid25@gmail.com",
        "FARMÁCIA": "farmacia.hcid@gmail.com",
        "FARMACOVIGILÂNCIA": "farmacovigilanciahcid@gmail.com",
        "FATURAMENTO": "contasmedicassocorrao2@gmail.com",
        "FISIOTERAPIA /PS - ENFERMARIAS": "fisioreabsoc2@gmail.com",
        "FISIOTERAPIA - UTI": "fisioterapiasocorrao@gmail.com",
        "FONOAUDIOLOGIA": "fonosocorrao2@yahoo.com",
        "GERÊNCIA DE ENFERMAGEM": "gerenciadeenf.hcid@gmail.com",
        "GERÊNCIA DA REABILITAÇÃO": "gerenciareab.hcid@gmail.com",
        "GERENCIAMENTO DE RESÍDUOS": "residuos.hcid@gmail.com",
        "HOTELARIA": "hotelaria.hcid@gmail.com",
        "LABORATÓRIO": "hcid.laboratorio@gmail.com",
        "MANUTENÇÃO": "manutencao.hcid@gmail.com",
        "MAQUEIROS": "centraltransporte.hcid@gmail.com",
        "NAQH": "naqhsoc2@gmail.com",
        "NCE": "nce.socorrao2@gmail.com",
        "NEP": "nep.hcid@gmail.com",
        "NEPEX": "nepexhcid@gmail.com",
        "NSP": "nspsoc2@gmail.com",
        "NIR": "coordenanirs2@gmail.com",
        "NUTRIÇÃO": "nutricaohospitaldacidade@gmail.com",
        "NVE": "nhe.hcid@gmail.com",
        "ODONTOLOGIA": "robertapalacioferes@gmail.com",
        "OUVIDORIA": "ouvidoriahcid@gmail.com",
        "PARECER DE EXAMES": "coodsaet2026@gmail.com",
        "PATRIMÔNIO": "patrimonio.hcid@gmail.com",
        "PORTARIA": "portaria.hcid@gmail.com",
        "PSICOLOGIA": "psicologia.hcid@gmail.com",
        "RECURSOS HUMANOS": "setorpessoal.hcid@gmail.com",
        "ROUPARIA": "rouparia.hcid@gmail.com",
        "SAEEM": "saeem.hcid@gmail.com",
        "SAME": "same25.hcid@gmail.com",
        "SAET": "coodsaet2026@gmail.com",
        "SADT": "sadthosp@gmail.com",
        "SDM/VERMELHA - ENFERMAGEM": "salavermelhasdm@gmail.com",
        "SERVIÇO SOCIAL": "servicosocialhcid@gmail.com",
        "SETOR DE IMAGEM": "sadt.hosp@gmail.com",
        "TECNOLOGIA DA INFORMAÇÃO": "suporteti.hcid@gmail.com",
        "TERAPIA OCUPACIONAL": "terapiaocupacional.hcid@gmail.com",
        "TRANSPORTE EXTERNO": "centraltransporte.hcid@gmail.com",
        "CENTRAL DE TRANSPORTE INTRA HOSPITALAR": "centraltransporte.hcid@gmail.com",
        "UNIDADE SENTINELA": "unidadesentinela@gmail.com",
        "UTI A - MÉDICOS": "pedrohfrota@gmail.com",
        "UTI B - MÉDICOS": "livia.goreth@ufma.br",
        "UTI C - MÉDICOS": "hugogama19@hotmail.com",
        "ORTOPEDIA - MÉDICOS": "abraaojunior@hotmail.com",
        "PRONTO SOCORRO - MÉDICOS": "gustavo1422@hotmail.com",
        "ALA A - MÉDICOS": "kledsonmarques10@gmail.com",
        "ALA B - MÉDICOS": "aliomarcamaradecarvalho@gmail.com",
        "CENTRO CIRÚRGICO - MÉDICOS": "felipeabreu805@gmail.com",
        "CARDIOLOGIA - MÉDICOS": "jordach.xerez@gmail.com",
        "ANESTESISTAS - MÉDICO": "rafaelcbarrosneto@yahoo.com.br",
        "XEROX/REPROGRAFIA": "centraldecopias.hcid@gmail.com",
    }
    
    for chave, email in emails.items():
        if chave in nome or nome in chave:
            return email
    
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
                        key=f"botao_{num_notif}_{idx}"
                    )
    except Exception as e:
        st.error(f"Erro ao processar planilha: {str(e)}")
