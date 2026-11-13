# =================================================================
# 📄 GERADOR DE MEMORANDOS — HOSPITAL DR. JACKSON LAGO
# Versão Final: Lista Completa de E-mails + Contagem Correta
# =================================================================

import os
import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document
import plotly.express as px

# =================================================================
# 📚 LISTA COMPLETA DE E-MAILS — 86 Setores
# =================================================================
LISTA_SETORES = [
    ("1. DIREÇÃO", "direcao.hospitaldacidade@gmail.com"),
    ("2. GABINETE DA DIREÇÃO", "gabinete.hospitaldacidade@gmail.com"),
    ("3. AGÊNCIA TRANSFUSIONAL", "agenciatransfusional.hcid@gmail.com"),
    ("4. ALA A - ENFERMAGEM", "alaaclinicamedica@gmail.com"),
    ("5. ALA B - ENFERMAGEM", "alab.hcid@gmail.com"),
    ("6. ALA C - ENFERMAGEM", "cirurgica.hcid@gmail.com"),
    ("7. ALA D - ENFERMAGEM", "ortopediaalad.hcid@gmail.com"),
    ("8. ALA G/H", "aladgh.hcid@gmail.com"),
    ("9. ALA E - ENFERMAGEM", "Ala não existe"),
    ("10. ALA F - ENFERMAGEM", "utihcid@gmail.com"),
    ("11. ALA I - ENFERMAGEM", "ortopediaalai.hcid@gmail.com"),
    ("12. ALA J - ENFERMAGEM", "ortopediaalaj.hcid@gmail.com"),
    ("13. ALA L - ENFERMAGEM", "alalortopedia.hcid@gmail.com"),
    ("14. ALMOXARIFADO", "almoxarifadosocorrao2@gmail.com"),
    ("15. AMBULATÓRIO", "ambulatorio.hcid@gmail.com"),
    ("16. ASSESSORIA JURÍDICA", "fernandoacmoares.advogado@gmail.com"),
    ("17. CAC", "cac.socorrao2@gmail.com"),
    ("18. CENTRAL DE CÓPIAS", "centraldecopias.hcid@gmail.com"),
    ("19. CCIH", "ccih.hcid@gmail.com"),
    ("20. CENTRO CIRÚRGICO - ENFERMAGEM", "centrocirurgicosoc2@gmail.com"),
    ("21. COMISSÃO DE ÓBITOS", "comissaodeobito.hdacidade@gmail.com"),
    ("22. COMISSÃO DE ÉTICA EM ENFERMAGEM", "comissaodeetica.soc2@gmail.com"),
    ("23. CME", "cmeesterilizacao@gmail.com"),
    ("24. COMISSÃO DE ÉTICA MÉDICA", "cmeesterilizacao@gmail.com"),
    ("25. COMISSÃO DE CUIDADOS PALIATIVOS", "utihcid@gmail.com"),
    ("26. COMISSÃO DE REVISÃO DE PRONTUÁRIO", "documentos.soc2@gmail.com"),
    ("27. CENTRAL DE TRANSPORTE INTRA HOSPITALAR", "documentos.soc2@gmail.com"),
    ("28. DOCUMENTOS INSTITUCIONAIS (NAQH)", "documentos.soc2@gmail.com"),
    ("29. ECP", "carlosemilioecp@gmail.com"),
    ("30. ENDOSCOPIA", "endoscopiasoc2@gmail.com"),
    ("31. EDOT", "edot.hcid@gmail.com"),
    ("32. EMTN", "emtnhospitaldacidade@gmail.com"),
    ("33. ESCALAS", "escalas.hcid25@gmail.com"),
    ("34. FARMÁCIA", "farmacia.hcid@gmail.com"),
    ("35. FARMACOVIGILÂNCIA", "farmacovigilanciahcid@gmail.com"),
    ("36. FATURAMENTO", "contasmedicassocorrao2@gmail.com"),
    ("37. FISIOTERAPIA / PS - ENFERMARIAS", "fisioreabsoc2@gmail.com"),
    ("38. FISIOTERAPIA - UTI", "fisioterapiasocorrao@gmail.com"),
    ("39. FONOAUDIOLOGIA", "fonosocorrao2@yahoo.com"),
    ("40. GERÊNCIA DE ENFERMAGEM", "gerenciadeenf.hcid@gmail.com"),
    ("41. GERÊNCIA DA REABILITAÇÃO", "gerenciareab.hcid@gmail.com"),
    ("42. GERENCIAMENTO DE RESÍDUOS", "residuos.hcid@gmail.com"),
    ("43. HOTELARIA", "hotelaria.hcid@gmail.com"),
    ("44. LABORATÓRIO", "hcid.laboratorio@gmail.com"),
    ("45. MANUTENÇÃO", "manutencao.hcid@gmail.com"),
    ("46. MAQUEIROS", "centraltransporte.hcid@gmail.com"),
    ("47. NAQH", "naqhsoc2@gmail.com"),
    ("48. NCE", "ce.socorrao2@gmail.com"),
    ("49. NEP", "nep.hcid@gmail.com"),
    ("50. NEPEX", "nepexhcid@gmail.com"),
    ("51. NIR", "coordenanirs2@gmail.com"),
    ("52. NSP", "nspsoc2@gmail.com"),
    ("53. NUTRIÇÃO", "nutricaohospitaldacidade@gmail.com"),
    ("54. NVE", "nhe.hcid@gmail.com"),
    ("55. ODONTOLOGIA", "robertapalacioferes@gmail.com"),
    ("56. OUVIDORIA", "ouvidoriahcid@gmail.com"),
    ("57. PARECER DE EXAMES", "coodsaet2026@gmail.com"),
    ("58. PATRIMÔNIO", "patrimonio.hcid@gmail.com"),
    ("59. PORTARIA", "portaria.hcid@gmail.com"),
    ("60. RECURSOS HUMANOS", "setorpessoal.hcid@gmail.com"),
    ("61. ROUPARIA", "rouparia.hcid@gmail.com"),
    ("62. SAEEM", "saeem.hcid@gmail.com"),
    ("63. SAME", "same25.hcid@gmail.com"),
    ("64. SAET", "coodsaet2026@gmail.com"),
    ("65. SADT", "sadthosp@gmail.com"),
    ("66. SDM/VERMELHA - ENFERMAGEM", "salavermelhasdm@gmail.com"),
    ("67. SERVIÇO SOCIAL", "servicosocialhcid@gmail.com"),
    ("68. SETOR DE IMAGEM", "sadt.hosp@gmail.com"),
    ("69. TECNOLOGIA DA INFORMAÇÃO", "suporteti.hcid@gmail.com"),
    ("70. TERAPIA OCUPACIONAL", "terapiaocupacional.hcid@gmail.com"),
    ("71. TRANSPORTE EXTERNO", "centraltransporte.hcid@gmail.com"),
    ("72. UNIDADE SENTINELA", "unidadesentinela@gmail.com"),
    ("73. UTI A - MÉDICOS", "pedrohfrota@gmail.com"),
    ("74. UTI B - MÉDICOS", "livia.goreth@ufma.br"),
    ("75. UTI C - MÉDICOS", "hugogama19@hotmail.com"),
    ("76. ORTOPEDIA - MÉDICOS", "abraaojunior@hotmail.com"),
    ("77. PRONTO SOCORRO - MÉDICOS", "gustavo1422@hotmail.com"),
    ("78. ALA A - MÉDICOS", "kledsonmarques10@gmail.com"),
    ("79. ALA B - MÉDICOS", "aliomarcamaradecarvalho@gmail.com"),
    ("80. CENTRO CIRÚRGICO - MÉDICOS", "felipeabreu805@gmail.com"),
    ("81. CARDIOLOGIA - MÉDICOS", "jordach.xerez@gmail.com"),
    ("82. ANESTESISTAS - MÉDICO", "rafaelcbarrosneto@yahoo.com.br"),
    ("83. XEROX/REPROGRAFIA", "centraldecopias.hcid@gmail.com"),
]

df_setores = pd.DataFrame(LISTA_SETORES, columns=["SETOR", "E-MAIL"])
df_setores = df_setores[df_setores["E-MAIL"].str.strip() != ""].reset_index(drop=True)

# =================================================================
# 🎯 CONFIGURAÇÃO DA PÁGINA
# =================================================================
st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
aba_processar, aba_estatisticas = st.tabs(["📝 Lista de Memorandos", "📊 Estatísticas"])

# =================================================================
# 🛠️ FUNÇÕES AUXILIARES
# =================================================================
def substituir_texto_no_docx(doc, dados):
    for p in doc.paragraphs:
        texto_completo = "".join(run.text for run in p.runs)
        alterado = False
        for chave, valor in dados.items():
            if chave in texto_completo:
                texto_completo = texto_completo.replace(chave, str(valor))
                alterado = True
        if alterado:
            p.clear()
            p.add_run(texto_completo)
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                substituir_texto_no_docx(celula, dados)

def formatar_data_br(valor):
    try:
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan":
            return ""
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except:
        return str(valor).strip()

def limpar_numero(valor):
    if pd.isna(valor) or str(valor).strip().lower() == "nan":
        return ""
    try:
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor).strip()
    except:
        return str(valor)

def tratar_texto(valor):
    if pd.isna(valor) or str(valor).strip().lower() in ["nan", "none", ""]:
        return ""
    return str(valor).strip()

def data_por_extenso(dt):
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    return f"{dt.day} de {meses[dt.month-1]} de {dt.year}"

def buscar_email_setor(nome_setor):
    if not nome_setor or str(nome_setor).strip() == "":
        return ""
    nome_setor = str(nome_setor).strip().upper()
    for _, linha in df_setores.iterrows():
        setor_lista = linha["SETOR"].upper()
        if nome_setor in setor_lista or setor_lista in nome_setor:
            return linha["E-MAIL"]
    return ""

def montar_corpo_email(num_memo, num_notif, gestor, setor, nome_paciente, data_envio):
    assunto = f"RETIFICAÇÃO E COMPLEMENTAÇÃO — Memorando Nº {num_memo} / Notificação Nº {num_notif}"
    corpo = f"""
Bom Dia, Boa Tarde, Boa Noite - Prezados,

Apenas Retificando e Complementando algumas Informações no Memorando — Campo Descrição.

Segue em Anexo o Memorando Nº {num_memo} para ser analisado e respondido (via e-mail) em até 15 dias após a data presente.

DADOS:
• Paciente: {nome_paciente}
• Setor Destinatário: {setor}
• Memorando Nº: {num_memo}
• Notificação Nº: {num_notif}

ATENÇÃO: A resposta via e-mail deve constar um arquivo em forma de Word ou PDF para arquivamento de respostas conforme rotina institucional. Não serão aceitas mensagens via e-mail sem arquivo como resposta.

Atenciosamente,

Ezequias S. Santos
Agente Administrativo - NAQH & NSP
Hospital da Cidade Dr. Jackson Lago
E-mail: nspsoc2@gmail.com
"""
    return assunto, corpo.strip()

# =================================================================
# 📄 FUNÇÃO DE PROCESSAR BLOCO
# =================================================================
def processar_bloco(linha, bloco_nome, idx, nome_paciente, data_extenso,
                    col_gestor, col_setor, col_memo,
                    dt_ocorr, dt_notif, local, tipo,
                    classificacao, desc, leito, setor_notificante, sugestao, turno_bruto):
    
    gestor = tratar_texto(linha.iloc[col_gestor]) if col_gestor < len(linha) else ""
    setor_destino = tratar_texto(linha.iloc[col_setor]) if col_setor < len(linha) else ""
    num_memo = limpar_numero(linha.iloc[col_memo]) if col_memo < len(linha) else ""
    
    if not num_memo or num_memo == "" or num_memo.upper() == "Nº MEMO 01":
        return None
    
    num_notif = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
    status = tratar_texto(linha.iloc[1]).upper() if len(linha) > 1 else "PENDENTE"
    email_destino = buscar_email_setor(setor_destino)
    
    if pd.isna(turno_bruto):
        m, t, n = "", "", ""
    else:
        tx = str(turno_bruto).strip().upper()
        m = "X" if "MANH" in tx else ""
        t = "X" if "TARD" in tx else ""
        n = "X" if "NOIT" in tx else ""
    
    dados_memorando = {
        "{{numero_memorando}}": num_memo,
        "{{gestor}}": gestor,
        "{{setor}}": setor_destino,
        "{{notificacao_n}}": num_notif,
        "{{data_notificacao}}": dt_notif,
        "{{data_ocorrencia}}": dt_ocorr,
        "{{localizacao}}": local,
        "{{tipo_incidente}}": tipo,
        "{{classificacao_incidente}}": classificacao,
        "{{descricao_notificacao}}": desc,
        "{{nome_paciente}}": nome_paciente,
        "{{leito}}": leito,
        "{{setor_notificante}}": setor_notificante,
        "{{sugestao}}": sugestao,
        "{{data_envio}}": data_extenso,
        "{{m}}": m,
        "{{t}}": t,
        "{{n}}": n,
    }
    
    assunto_email, corpo_email = montar_corpo_email(num_memo, num_notif, gestor, setor_destino, nome_paciente, data_extenso)
    
    return {
        "bloco": bloco_nome,
        "num_memo": num_memo,
        "num_notif": num_notif,
        "status": status,
        "gestor": gestor,
        "setor": setor_destino,
        "email": email_destino,
        "dados_docx": dados_memorando,
        "assunto_email": assunto_email,
        "corpo_email": corpo_email,
    }

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
        st.error(f"Erro: {str(e)}")
        return None

# =================================================================
# 📊 COLETAR TODOS OS MEMORANDOS — CONTAGEM REAL
# =================================================================
def coletar_todos_memorandos(df):
    todos_memorandos = []
    for _, linha in df.iterrows():
        status = tratar_texto(linha.iloc[1]) if len(linha) > 1 else ""
        dt_notif = formatar_data_br(linha.iloc[3]) if len(linha) > 3 else ""
        
        for g_col, s_col, m_col in [(13,14,15), (17,18,19), (22,23,24), (26,27,28)]:
            num_memo = limpar_numero(linha.iloc[m_col]) if m_col < len(linha) else ""
            setor = tratar_texto(linha.iloc[s_col]) if s_col < len(linha) else ""
            
            if num_memo and num_memo != "" and "Nº MEMO" not in num_memo:
                data_obj = None
                try:
                    if dt_notif:
                        data_obj = pd.to_datetime(dt_notif, dayfirst=True)
                except:
                    pass
                
                todos_memorandos.append({
                    "numero_memo": num_memo,
                    "setor": setor,
                    "status": status.upper() if status else "PENDENTE",
                    "data": data_obj,
                    "notificacao": limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
                })
    
    return pd.DataFrame(todos_memorandos)

# =================================================================
# 📑 ABA 1 — LISTA DE MEMORANDOS
# =================================================================
with aba_processar:
    st.title("📝 Lista de Memorandos")
    data_selecionada = st.date_input("Data de Envio", datetime.now())
    arquivo_excel = st.file_uploader("Planilha com Incidentes (.xlsx)", type=["xlsx"])
    
    if arquivo_excel:
        try:
            df = pd.read_excel(arquivo_excel, engine="openpyxl")
            df_memos = coletar_todos_memorandos(df)
            
            # ✅ CONTAGEM CORRETA — APENAS DE MEMORANDOS VÁLIDOS
            total_memorandos = len(df_memos)
            enviados = len(df_memos[df_memos["status"] == "ENVIADO"])
            pendentes = total_memorandos - enviados
            
            st.subheader("📊 Resumo — Por Memorando")
            col1, col2, col3 = st.columns(3)
            col1.metric("📋 Total de Memorandos", total_memorandos)
            col2.metric("✅ Já Enviados", enviados)
            col3.metric("⏳ Pendentes", pendentes)
            
            st.info(f"Mostrando todos os {total_memorandos} memorandos — enviados aparecem riscados")
            
            data_extenso = data_por_extenso(data_selecionada)
            
            # Processar TODAS as linhas
            for idx, linha in df.iterrows():
                num_notif = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
                if not num_notif or num_notif.upper() == "STATUS" or "NOTIF" in str(num_notif).upper():
