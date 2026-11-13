# =================================================================
# 📄 GERADOR DE MEMORANDOS — HOSPITAL DR. JACKSON LAGO
# Versão: Setores Corrigidos + Texto do E-mail Atualizado
# =================================================================

import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document

# =================================================================
# 📌 VALOR FIXO — TOTAL GERAL DE MEMORANDOS
# =================================================================
TOTAL_GERAL_MEMORANDOS = 964

# =================================================================
# 📚 LISTA COMPLETA DE E-MAILS — CORRIGIDA E ATUALIZADA
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
    ("10. ALA F - ENFERMAGEM", "utihcid@gmail.com"),
    ("11. ALA I - ENFERMAGEM", "ortopediaalai.hcid@gmail.com"),
    ("12. ALA J - ENFERMAGEM", "ortopediaalaj.hcid@gmail.com"),
    ("13. ALA L - ENFERMAGEM", "alalortopedia.hcid@gmail.com"),
    ("14. ALMOXARIFADO", "almoxarifadosocorrao2@gmail.com"),
    ("15. AMBULATÓRIO", "ambulatorio.hcid@gmail.com"),
    ("16. ASSESSORIA JURÍDICA", "fernandoacmoraes.advogado@gmail.com"),
    ("17. CAC", "cac.socorrao2@gmail.com"),
    ("18. CENTRAL DE CÓPIAS", "centraldecopias.hcid@gmail.com"),
    ("19. CCIH", "ccih.hcid@gmail.com"),
    ("20. CENTRO CIRÚRGICO - ENFERMAGEM", "centrocirurgicosoc2@gmail.com"),
    ("21. COMISSÃO DE ÓBITOS", "comissaodeobito.hdacidade@gmail.com"),
    ("22. COMISSÃO DE ÉTICA EM ENFERMAGEM", "comissaodeetica.soc2@gmail.com"),
    ("23. CME", "cmeesterilizacao@gmail.com"),
    ("27. CENTRAL DE TRANSPORTE INTRA HOSPITALAR", "centraltransporte.hcid@gmail.com"),
    ("CENTRAL INTERNA DE TRANSPORTE", "centraltransporte.hcid@gmail.com"),
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
    ("FISIOTERAPIA INTERNAÇÃO/PS", "fisioreabsoc2@gmail.com"),
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
    ("48. NCE", "nce.socorrao2@gmail.com"),
    ("49. NEP", "nep.hcid@gmail.com"),
    ("50. NEPEX", "nepexhcid@gmail.com"),
    ("51. NIR", "coordenanirs2@gmail.com"),
    ("52. NSP", "nspsoc2@gmail.com"),
    ("53. NUTRIÇÃO", "nutricaohospitaldacidade@gmail.com"),
    ("54. NVE", "nhe.hcid@gmail.com"),
    ("56. ODONTOLOGIA", "robertapalacioferes@gmail.com"),
    ("57. OUVIDORIA", "ouvidoriahcid@gmail.com"),
    ("58. PARECER DE EXAMES", "coodsaet2026@gmail.com"),
    ("59. PATRIMÔNIO", "patrimonio.hcid@gmail.com"),
    ("60. PORTARIA", "portaria.hcid@gmail.com"),
    ("61. PSICOLOGIA", "psicologia.hcid@gmail.com"),
    ("62. RECURSOS HUMANOS", "setorpessoal.hcid@gmail.com"),
    ("63. ROUPARIA", "rouparia.hcid@gmail.com"),
    ("64. SAEEM", "saeem.hcid@gmail.com"),
    ("65. SAME", "same25.hcid@gmail.com"),
    ("66. SAET", "coodsaet2026@gmail.com"),
    ("67. SADT", "sadthosp@gmail.com"),
    ("68. SDM/VERMELHA - ENFERMAGEM", "salavermelhasdm@gmail.com"),
    ("SALA VERMELHA", "salavermelhasdm@gmail.com"),
    ("69. SERVIÇO SOCIAL", "servicosocialhcid@gmail.com"),
    ("70. SETOR DE IMAGEM", "sadt.hosp@gmail.com"),
    ("71. TECNOLOGIA DA INFORMAÇÃO", "suporteti.hcid@gmail.com"),
    ("72. TERAPIA OCUPACIONAL", "terapiaocupacional.hcid@gmail.com"),
    ("73. TRANSPORTE EXTERNO", "centraltransporte.hcid@gmail.com"),
    ("74. UNIDADE SENTINELA", "unidadesentinela@gmail.com"),
    ("75. UTI A - MÉDICOS", "pedrohfrota@gmail.com"),
    ("76. UTI B - MÉDICOS", "livia.goreth@ufma.br"),
    ("77. UTI C - MÉDICOS", "hugogama19@hotmail.com"),
    ("78. ORTOPEDIA - MÉDICOS", "abraaojunior@hotmail.com"),
    ("80. PRONTO SOCORRO - MÉDICOS", "gustavo1422@hotmail.com"),
    ("81. ALA A - MÉDICOS", "kledsonmarques10@gmail.com"),
    ("82. ALA B - MÉDICOS", "aliomarcamaradecarvalho@gmail.com"),
    ("83. CENTRO CIRÚRGICO - MÉDICOS", "felipeabreu805@gmail.com"),
    ("84. CARDIOLOGIA - MÉDICOS", "jordach.xerez@gmail.com"),
    ("85. ANESTESISTAS - MÉDICO", "rafaelcbarrosneto@yahoo.com.br"),
    ("86. XEROX/REPROGRAFIA", "centraldecopias.hcid@gmail.com"),
]

df_setores = pd.DataFrame(LISTA_SETORES, columns=["SETOR", "E-MAIL"])
df_setores = df_setores[df_setores["E-MAIL"].str.strip() != ""].reset_index(drop=True)

# =================================================================
# 🎯 CONFIGURAÇÃO DA PÁGINA
# =================================================================
st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")

# =================================================================
# 🛠️ FUNÇÕES AUXILIARES
# =================================================================
def tratar_texto(valor):
    if pd.isna(valor) or str(valor).strip().lower() in ["nan", "none", ""]:
        return ""
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

def data_por_extenso(dt):
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    return f"{dt.day} de {meses[dt.month-1]} de {dt.year}"

# ✅ CORRIGIDA BUSCA DE E-MAIL — agora busca parcial e ignora maiúsculas/minúsculas
def buscar_email_setor(nome_setor):
    if not nome_setor or str(nome_setor).strip() == "":
        return "nspsoc2@gmail.com"
    nome_setor = str(nome_setor).strip().upper()
    for _, linha in df_setores.iterrows():
        setor_lista = linha["SETOR"].upper()
        if nome_setor in setor_lista or setor_lista in nome_setor:
            return linha["E-MAIL"]
    return "nspsoc2@gmail.com"

# ✅ TEXTO DO E-MAIL EXATAMENTE COMO VOCÊ PEDIU
def montar_corpo_email(num_memo, ano):
    return f"""Boa Tarde, Prezados, ou bom dia de acordo com o horário.

Segue em anexo o Memorando Nº {num_memo}/{ano} para ser analisado e respondido (via e-mail) em até 15 dias após a data presente.

ATENÇÃO: A resposta via e-mail deve constar um arquivo em forma de Word ou PDF para arquivamento de respostas conforme rotina institucional. Não serão aceitas mensagens via e-mail sem arquivo como resposta.

Atenciosamente,

Ezequias S. Santos
Agente Administrativo - NAQH & NSP
nspsoc2@gmail.com
"""

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

# =================================================================
# 📊 CONTAR ENVIADOS E CALCULAR PENDENTES
# =================================================================
def contar_enviados(df):
    qtd = 0
    for _, linha in df.iterrows():
        status = tratar_texto(linha.iloc[1]).upper() if len(linha) > 1 else ""
        if status == "ENVIADO":
            for m_col in [15, 19, 24, 28]:
                if m_col < len(linha):
                    num = limpar_numero(linha.iloc[m_col])
                    if num and num != "" and "Nº MEMO" not in num:
                        qtd += 1
    return qtd

# =================================================================
# 📑 INTERFACE PRINCIPAL
# =================================================================
st.title("📄 Emissor de Memorandos — Hospital da Cidade")
st.divider()

arquivo_excel = st.file_uploader("📂 Carregue a Planilha de Incidentes (.xlsx)", type=["xlsx"])

if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel, engine="openpyxl")
        
        # ✅ RESUMO GERAL
        enviados = contar_enviados(df)
        total_memorandos = TOTAL_GERAL_MEMORANDOS
        pendentes = total_memorandos - enviados
        
        st.subheader("📊 RESUMO GERAL")
        col1, col2, col3 = st.columns(3)
        col1.metric("📋 Total Geral de Memorandos", f"{total_memorandos}")
        col2.metric("✅ Já Enviados", f"{enviados}")
        col3.metric("⏳ Pendentes", f"{pendentes}")
        st.info(f"Fórmula aplicada: **{total_memorandos} − {enviados} = {pendentes}** memorandos a processar")
        st.divider()
        
        data_selecionada = st.date_input("📅 Data de Envio", datetime.now())
        ano_atual = data_selecionada.year
        data_extenso = data_por_extenso(data_selecionada)
        
        qtd_processados = 0
        
        for idx, linha in df.iterrows():
            num_notif = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
            if not num_notif or num_notif.upper() == "STATUS" or "NOTIF" in str(num_notif).upper():
                continue
            
            nome_paciente = tratar_texto(linha.iloc[9]) if len(linha) > 9 else "Paciente"
            if nome_paciente.upper() == "PACIENTE":
                continue
            
            status_registro = tratar_texto(linha.iloc[1]).upper() if len(linha) > 1 else "PENDENTE"
            
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
                st.markdown(f"~~📋 Notificação {num_notif} — {nome_paciente} — {len(blocos_memo)} memorando(s)~~ ✅ **ENVIADO**")
                continue
            else:
                qtd_processados += 1
                with st.expander(f"📋 Notificação {num_notif} — {nome_paciente} — {len(blocos_memo)} memorando(s)", expanded=True):
                    st.write(f"📅 Data: {data_extenso} | 📝 Total de Memorandos: {len(blocos_memo)}")
                    
                    for b in blocos_memo:
                        email_destino = buscar_email_setor(b["setor"])
                        
                        # ✅ CORRIGIDO — sem ano duplicado!
                        st.subheader(f"📄 Memorando 01 — Nº {b['num_memo']}/{ano_atual}")
                        st.write(f"**Setor:** {b['setor']} | **Gestor:** {b['gestor']}")
                        st.info(f"📧 E-mail: {email_destino}")
                        
                        # ✅ TEXTO DO E-MAIL EXIBIDO NA TELA
                        corpo_email = montar_corpo_email(b['num_memo'], ano_atual)
                        with st.expander("📋 Ver texto do E-mail"):
                            st.text(corpo_email)
                        
                        dados_docx = {
                            "{{numero_memorando}}": f"{b['num_memo']}/{ano_atual}",
                            "{{gestor}}": b["gestor"],
                            "{{setor}}": b["setor"],
                            "{{notificacao_n}}": num_notif,
                            "{{nome_paciente}}": nome_paciente,
                            "{{data_envio}}": data_extenso,
                        }
                        
                        arquivo_docx = gerar_docx(dados_docx)
                        if arquivo_docx:
                            st.download_button(
                                label="📄 Baixar Memorando (DOCX)",
                                data=arquivo_docx,
                                file_name=f"MEMORANDO_{b['num_memo']}_{ano_atual}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )
                        st.divider()
        
        st.success(f"✅ Processamento concluído! {qtd_processados} memorandos disponíveis para download.")
    
    except Exception as e:
        st.error(f"❌ Erro ao processar arquivo: {str(e)}")
else:
    st.info("👆 Faça o upload da planilha .xlsx para começar")
