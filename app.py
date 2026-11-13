# =================================================================
# 📄 GERADOR DE MEMORANDOS — HOSPITAL DR. JACKSON LAGO
# Versão com todos os campos do modelo oficial
# =================================================================

import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document

# =================================================================
# 📌 CONFIGURAÇÕES FIXAS
# =================================================================
TOTAL_GERAL_MEMORANDOS = 964

# =================================================================
# 📚 LISTA DE E-MAILS
# =================================================================
LISTA_SETORES = [
    ("SALA VERMELHA", "salavermelhasdm@gmail.com"),
    ("SDM/VERMELHA", "salavermelhasdm@gmail.com"),
    ("FISIOTERAPIA", "fisioreabsoc2@gmail.com"),
    ("FISIOTERAPIA INTERNAÇÃO/PS", "fisioreabsoc2@gmail.com"),
    ("CENTRAL DE TRANSPORTE", "centraltransporte.hcid@gmail.com"),
    ("CENTRAL INTERNA DE TRANSPORTE", "centraltransporte.hcid@gmail.com"),
    ("ALA A", "alaaclinicamedica@gmail.com"),
    ("ALA B", "alab.hcid@gmail.com"),
    ("ALA C", "cirurgica.hcid@gmail.com"),
    ("ALA D", "ortopediaalad.hcid@gmail.com"),
    ("ALA F", "utihcid@gmail.com"),
    ("ALA G/H", "aladgh.hcid@gmail.com"),
    ("ALA I", "ortopediaalai.hcid@gmail.com"),
    ("ALA J", "ortopediaalaj.hcid@gmail.com"),
    ("ALA L", "alalortopedia.hcid@gmail.com"),
    ("NSP", "nspsoc2@gmail.com"),
    ("NAQH", "naqhsoc2@gmail.com"),
    ("SERVIÇO SOCIAL", "servicosocialhcid@gmail.com"),
    ("SADT", "sadthosp@gmail.com"),
    ("SETOR DE IMAGEM", "sadt.hosp@gmail.com"),
    ("PRONTO SOCORRO", "gustavo1422@hotmail.com"),
    ("CENTRO CIRÚRGICO", "centrocirurgicosoc2@gmail.com"),
    ("FARMACIA", "farmacia.hcid@gmail.com"),
    ("NIR", "coordenanirs2@gmail.com"),
    ("HOTELARIA", "hotelaria.hcid@gmail.com"),
    ("OUVIDORIA", "ouvidoriahcid@gmail.com"),
    ("RECURSOS HUMANOS", "setorpessoal.hcid@gmail.com"),
    ("SAME", "same25.hcid@gmail.com"),
    ("PSICOLOGIA", "psicologia.hcid@gmail.com"),
    ("TERAPIA OCUPACIONAL", "terapiaocupacional.hcid@gmail.com"),
    ("EMTN", "emtnhospitaldacidade@gmail.com"),
    ("ENDOSCOPIA", "endoscopiasoc2@gmail.com"),
    ("FARMACOVIGILÂNCIA", "farmacovigilanciahcid@gmail.com"),
    ("GESTOR DE ENFERMAGEM", "gerenciadeenf.hcid@gmail.com"),
]

df_setores = pd.DataFrame(LISTA_SETORES, columns=["SETOR", "E-MAIL"])
df_setores = df_setores[df_setores["E-MAIL"].str.strip() != ""].reset_index(drop=True)

# =================================================================
# 🎯 CONFIGURAÇÃO DA PÁGINA
# =================================================================
st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
st.title("📄 Emissor de Memorandos — Hospital da Cidade")
st.divider()

# =================================================================
# 🛠️ FUNÇÕES AUXILIARES
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

def data_por_extenso(dt):
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    return f"{dt.day} de {meses[dt.month-1]} de {dt.year}"

def buscar_email_setor(nome_setor):
    if not nome_setor:
        return "nspsoc2@gmail.com"
    nome_upper = str(nome_setor).strip().upper()
    for _, linha in df_setores.iterrows():
        if nome_upper in linha["SETOR"].upper() or linha["SETOR"].upper() in nome_upper:
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
        st.error("❌ Arquivo modelo_memorando.docx não encontrado!")
        return None
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")
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
# 📑 PRINCIPAL
# =================================================================
arquivo_excel = st.file_uploader("📂 Carregue a Planilha (.xlsx)", type=["xlsx"])

if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel, engine="openpyxl")
        
        enviados = contar_enviados(df)
        total = TOTAL_GERAL_MEMORANDOS
        pendentes = total - enviados
        
        st.subheader("📊 RESUMO GERAL")
        col1, col2, col3 = st.columns(3)
        col1.metric("📋 Total de Memorandos", total)
        col2.metric("✅ Já Enviados", enviados)
        col3.metric("⏳ Pendentes", pendentes)
        st.info(f"**{total} − {enviados} = {pendentes}** memorandos pendentes")
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
            
            # Extrair todos os campos do Excel
            data_ocorrencia = tratar_texto(linha.iloc[2]) if len(linha) > 2 else ""
            data_notificacao = tratar_texto(linha.iloc[3]) if len(linha) > 3 else ""
            turno = tratar_texto(linha.iloc[4]).upper() if len(linha) > 4 else ""
            localizacao = tratar_texto(linha.iloc[5]) if len(linha) > 5 else ""
            tipo_incidente = tratar_texto(linha.iloc[6]) if len(linha) > 6 else ""
            classificacao_incidente = tratar_texto(linha.iloc[7]) if len(linha) > 7 else ""
            descricao_notificacao = tratar_texto(linha.iloc[8]) if len(linha) > 8 else ""
            leito = tratar_texto(linha.iloc[10]) if len(linha) > 10 else ""
            setor_notificante = tratar_texto(linha.iloc[11]) if len(linha) > 11 else ""
            sugestao = tratar_texto(linha.iloc[12]) if len(linha) > 12 else ""
            
            # Definir marcações de turno
            m_manha = "☐"
            m_tarde = "☐"
            m_noite = "☐"
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
                st.markdown(f"~~📋 Notificação {num_notif} — {nome_paciente}~~ ✅ **ENVIADO**")
                continue
            
            qtd_processados += 1
            with st.expander(f"📋 Notificação {num_notif} — {nome_paciente} ({len(blocos_memo)} memo(s))", expanded=True):
                st.write(f"📅 Data de Emissão: {data_extenso}")
                st.write(f"📌 Data da Ocorrência: {data_ocorrencia} | 📌 Data da Notificação: {data_notificacao}")
                st.write(f"📍 Local: {localizacao} | 🏥 Setor Notificante: {setor_notificante}")
                st.write(f"⚠️ Tipo: {tipo_incidente} | Classificação: {classificacao_incidente}")
                st.write(f"🛏️ Leito: {leito} | 🔄 Turno: {turno}")
                if sugestao:
                    st.write(f"💡 Sugestão: {sugestao}")
                
                for b in blocos_memo:
                    email_destino = buscar_email_setor(b["setor"])
                    st.subheader(f"📄 MEMO: NSP Nº {b['num_memo']}/{ano_atual}")
                    st.write(f"**Para:** {b['gestor']} da {b['setor']} | **Assunto:** Nº {num_notif}")
                    st.info(f"📧 E-mail: {email_destino}")
                    
                    # Texto do e-mail
                    corpo_email = f"""Boa Tarde, Prezados, ou bom dia de acordo com o horário.

Segue em anexo o Memorando Nº {b['num_memo']}/{ano_atual} referente à notificação Nº {num_notif}, para ser analisado e respondido (via e-mail) em até 15 dias após a data presente.

ATENÇÃO: A resposta via e-mail deve constar um arquivo em forma de Word ou PDF para arquivamento de respostas conforme rotina institucional. Não serão aceitas mensagens via e-mail sem arquivo como resposta.

Atenciosamente,

Ezequias S. Santos
Agente Administrativo - NAQH & NSP
nspsoc2@gmail.com
"""
                    with st.expander("✉️ Texto do E-mail"):
                        st.text(corpo_email)
                    
                    # Todos os campos para substituir no Word
                    dados_docx = {
                        "{{numero_memorando}}": f"{b['num_memo']}/{ano_atual}",
                        "{{assunto_numero}}": num_notif,
                        "{{gestor}}": b["gestor"],
                        "{{setor_destino}}": b["setor"],
                        "{{data_ocorrencia}}": data_ocorrencia,
                        "{{data_notificacao}}": data_notificacao,
                        "{{m}}": m_manha,
                        "{{t}}": m_tarde,
                        "{{n}}": m_noite,
                        "{{localizacao}}": localizacao,
                        "{{tipo_incidente}}": tipo_incidente,
                        "{{classificacao_incidente}}": classificacao_incidente,
                        "{{descricao_notificacao}}": descricao_notificacao,
                        "{{paciente}}": nome_paciente,
                        "{{leito}}": leito,
                        "{{setor_notificante}}": setor_notificante,
                        "{{sugestao}}": sugestao,
                        "{{data_extenso}}": data_extenso,
                    }
                    
                    arquivo_docx = gerar_docx(dados_docx)
                    if arquivo_docx:
                        st.download_button(
                            label="📄 Baixar Memorando (DOCX)",
                            data=arquivo_docx,
                            file_name=f"NSP_{b['num_memo']}_{ano_atual}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    st.divider()
        
        st.success(f"✅ Pronto! {qtd_processados} memorandos disponíveis.")
    
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")
else:
    st.info("👆 Faça upload do arquivo .xlsx para começar")
