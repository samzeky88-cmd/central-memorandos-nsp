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
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
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
            for m_col in [15, 19, 24, 28]:
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
            localizacao = tratar_texto(linha.iloc[5]) if len(linha) > 5 else ""
            tipo_incidente = tratar_texto(linha.iloc[6]) if len(linha) > 6 else ""
            classificacao_incidente = tratar_texto(linha.iloc[7]) if len(linha) > 7 else ""
            descricao_notificacao = tratar_texto(linha.iloc[8]) if len(linha) > 8 else ""
            leito = tratar_texto(linha.iloc[10]) if len(linha) > 10 else ""
            setor_notificante = tratar_texto(linha.iloc[11]) if len(linha) > 11 else ""
            sugestao = tratar_texto(linha.iloc[12]) if len(linha) > 12 else ""
            
            m_manha = "☐"
            m_tarde = "☐"
            m_noite = "☐"
            if turno:
                if "MANHÃ" in turno or "MANHA" in turno:
                    m_manha = "☑"
                elif "TARDE" in turno:
                    m_tarde = "☑"
                elif "NOITE" in turno:
                    m_noite = "☑"
            
            blocos_memo = []
            for g_col, s_col, m_col in [(13,14,15), (17,18,19), (22,23,24), (26,27,28)]:
                if m_col >= len(linha):
                    continue
                num_memo = limpar_numero(linha.iloc[m_col])
                setor_destino = tratar_texto(linha.iloc[s_col])
                gestor = tratar_texto(linha.iloc[g_col])
                if num_memo and num_memo != "" and "Nº MEMO" not in num_memo:
                    blocos_memo.append({
                        "num_memo": num_memo,
                        "setor": setor_destino,
                        "gestor": gestor
                    })
            
            if not blocos_memo:
                continue
            
            if status_registro == "ENVIADO":
                st.markdown(f"~~Notificação {num_notif} — {nome_paciente}~~ **ENVIADO**")
                continue
            
            qtd_processados += 1
            with st.expander(f"Notificação {num_notif} — {nome_paciente} ({len(blocos_memo)} memo(s))", expanded=True):
                st.write(f"Data de Emiss o: {data_extenso}")
                st.write(f"Ocorr ncia: {data_ocorrencia} | Notifica o: {data_notificacao}")
                st.write(f"Local: {localizacao} | Setor Notificante: {setor_notificante}")
                st.write(f"Tipo: {tipo_incidente} | Classifica o: {classificacao_incidente}")
                st.write(f"Paciente: {nome_paciente} | Leito: {leito}")
                st.write(f"Turno: {m_manha} MANH  {m_tarde} TARDE  {m_noite} NOITE")
                if descricao_notificacao:
                    st.write(f"Descri o: {descricao_notificacao}")
                if sugestao:
                    st.write(f"Sugest o: {sugestao}")
                
                for b in blocos_memo:
                    email_destino = buscar_email_setor(b["setor"])
                    st.subheader(f"MEMO: NSP N {b['num_memo']}/{ano_atual}")
                    st.write(f"Para: {b['gestor']} da {b['setor']}")
                    st.write(f"Assunto: N {num_notif}")
                    st.info(f"E-mail: {email_destino}")
                    
                    corpo_email = "Boa Tarde, Prezados, ou bom dia de acordo com o hor rio.\n\n"
                    corpo_email += "Segue em anexo o Memorando N " + b["num_memo"] + "/" + str(ano_atual)
                    corpo_email += " referente notifica o N " + num_notif + ", para ser analisado e "
                    corpo_email += "respondido (via e-mail) em at 15 dias ap s a data presente.\n\n"
                    corpo_email += "ATEN O: A resposta via e-mail deve constar um arquivo em forma de Word "
                    corpo_email += "ou PDF para arquivamento de respostas conforme rotina institucional. "
                    corpo_email += "N o ser o aceitas mensagens via e-mail sem arquivo como resposta.\n\n"
                    corpo_email += "Atenciosamente,\n\nEzequias S. Santos\nAgente Administrativo - NAQH & NSP\nnspsoc2@gmail.com"
                    
                    with st.expander("Texto do E-mail"):
                        st.text(corpo_email)
                    
                    dados_docx = {
                        "{{numero_memorando}}": b["num_memo"] + "/" + str(ano_atual),
                        "{{notificacao_n}}": num_notif,
                        "{{gestor}}": b["gestor"],
                        "{{setor}}": b["setor"],
                        "{{data_ocorrencia}}": data_ocorrencia,
                        "{{data_notificacao}}": data_notificacao,
                        "{{m}}": m_manha,
                        "{{t}}": m_tarde,
                        "{{n}}": m_noite,
                        "{{localizacao}}": localizacao,
                        "{{tipo_incidente}}": tipo_incidente,
                        "{{classificacao_incidente}}": classificacao_incidente,
