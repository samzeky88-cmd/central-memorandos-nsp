import os
import pandas as pd
import streamlit as st
import io
import urllib.parse
from datetime import datetime
from docx import Document

st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
st.title("📄 Emissor de Memorandos Individuais - Hospital Dr. Jackson Lago")

st.markdown("### 🗓️ Configuração da Data de Envio")
data_selecionada = st.date_input("Selecione a data que sairá no cabeçalho do Memorando:", datetime.now())

arquivo_excel = st.file_uploader("Suba a planilha contendo os incidentes (.xlsx)", type=["xlsx"])
caminho_modelo = "modelo_memorando.docx"

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
    """Garante que as datas sejam exibidas no formato brasileiro DD/MM/AAAA"""
    try:
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan":
            return ""
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except:
        v_str = str(valor).strip().split(" ")[0]
        if "-" in v_str:
            parts = v_str.split("-")
            if len(parts) == 3:
                return f"{parts[2]}/{parts[1]}/{parts[0]}"
        return str(valor).strip()

def limpar_numero_float(valor):
    """Remove o .0 de números inteiros vindos do Excel (ex: 886.0 vira 886)"""
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
    """Evita que campos vazios ou nulos exibam a palavra 'nan'"""
    if pd.isna(valor) or str(valor).strip().lower() == "nan" or str(valor).strip().lower() == "none":
        return ""
    return str(valor).strip()

def obter_data_por_extenso(dt):
    """Gera a data selecionada por extenso em português brasileiro"""
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho",
        7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return f"{dt.day} de {meses[dt.month]} de {dt.year}"

@st.fragment
def renderizar_linha_paciente_sob_demanda(index, linha, col_paciente, col_notif, col_data_notif, col_data_ocorr, col_turno, col_tipo, col_classif, col_desc, col_leito, col_notificante, col_sugestao, col_localizacao, col_email, data_extenso_envio):
    nome_do_paciente = str(linha[col_paciente]).strip()
    
    # Coleta inteligente do número do memorando nas duas colunas possíveis (Memo 01 ou Memo 02)
    memo_01 = str(linha.get("Nº Memo 01", "")).strip()
    memo_02 = str(linha.get("Nº Memo 02", "")).strip()
    if memo_01 == "" or memo_01.lower() == "nan":
        num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" else "S-N"
    else:
        num_memo_cru = memo_01
        
    # Coleta o número da notificação da coluna correta
    num_notif = limpar_numero_float(linha.get(col_notif, "S-N"))
    
    # Padronização limpa do nome do arquivo final
    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
    nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
    
    # Lógica minuciosa e resiliente para marcação dos turnos com X
    turno_planilha = str(linha.get(col_turno, "")).strip().upper() if col_turno else ""
    marca_manha = "X" if "MANH" in turno_planilha or "MANHÃ" in turno_planilha else " "
    marca_tarde = "X" if "TARD" in turno_planilha or "TARDE" in turno_planilha else " "
    marca_noite = "X" if "NOIT" in turno_planilha or "NOITE" in turno_planilha else " "
    
    # Coleta resiliente do setor notificante para eliminar o erro do 'nan'
    setor_notificante_bruto = tratar_str_limpa(linha.get(col_notificante, "")) if col_notificante else ""
    if setor_notificante_bruto == "":
        setor_notificante_bruto = "NSP - NÚCLEO DE SEGURANÇA DO PACIENTE"

    # Vinculando os dados dinâmicos às tags do seu modelo Word
    dados_memorando = {
        "{{numero_memorando}}": num_memo_cru,
        "{{gestor}}": str(linha.get("Gestor 01", "")).strip(),
        "{{setor}}": str(linha.get("SETOR NOTIFICADO", "")).strip(),
        "{{notificacao_n}}": num_notif,
        "{{data_notificacao}}": formatar_data_br(linha.get(col_data_notif, "")) if col_data_notif else "",
        "{{data_ocorrencia}}": formatar_data_br(linha.get(col_data_ocorr, "")) if col_data_ocorr else "",
        "{{localizacao}}": str(linha.get(col_localizacao, "")).strip() if col_localizacao else "",
        "{{tipo_incidente}}": str(linha.get(col_tipo, "")).strip() if col_tipo else "",
        "{{classificacao_incidente}}": str(linha.get(col_classif, "")).strip() if col_classif else "",
        "{{descricao_notificacao}}": str(linha.get(col_desc, "")).strip() if col_desc else "",
        "{{nome_paciente}}": nome_do_paciente,
        "{{leito}}": limpar_numero_float(linha.get(col_leito, "")) if col_leito else "",
        "{{setor_notificante}}": setor_notificante_bruto,
        "{{sugestao}}": str(linha.get(col_sugestao, "")).strip() if col_sugestao else "",
        "{{m}}": marca_manha,
        "{{t}}": marca_tarde,
        "{{n}}": marca_noite,
        "{{data_envio}}": data_extenso_envio
    }
    
    # --- CONFIGURAÇÃO EXCLUSIVA DO TEXTO PADRÃO DO GMAIL ---
    hora_atual = datetime.now().hour
    saudacao = "Bom Dia Prezados" if hora_atual < 12 else "Boa Tarde Prezados"
    
    email_destino = tratar_str_limpa(linha.get(col_email, "")) if col_email else ""
    assunto_email = f"NSP - Memorando Nº {num_memo_cru} (Notificação Nº {num_notif})"
    
    corpo_email = (
        f"{saudacao},\n\n"
        f"Segue em Anexo o Memorando Nº {num_memo_cru} em anexo para ser analisado e respondido "
        f"(via e-mail) em até 15 dias após a data presente.\n\n"
        f"ATENÇÃO: A resposta via e-mail deve constar um arquivo em forma de word ou PDF para "
        f"arquivamento de respostas conforme rotina institucional. Não serão aceitas mensagens "
        f"via e-mail sem arquivo como resposta.\n\n"
        f"Segue abaixo a notificação para análise do incidente em equipe e resposta ao NSP.\n\n"
        f"Atenciosamente,\n"
        f"Ezequias S. Santos\n"
        f"Agente Administrativo"
    )
    
    link_gmail = (
        f"https://google.com?"
        f"view=cm&fs=1&tf=1"
        f"&to={urllib.parse.quote(email_destino)}"
        f"&su={urllib.parse.quote(assunto_email)}"
        f"&body={urllib.parse.quote(corpo_email)}"
    )

    col_nome, col_word, col_pdf, col_email_btn = st.columns(4)
    
    with col_nome:
        st.markdown(f"**🔹 {nome_do_paciente}**")
        
    with col_word:
        doc_word = Document(caminho_modelo)
        substituir_texto_protegendo_logos(doc_word, dados_memorando)
        word_io = io.BytesIO()
        doc_word.save(word_io)
        word_io.seek(0)
        st.download_button(
            label="📝 WORD",
            data=word_io.getvalue(),
            file_name=f"{nome_base_arquivo}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"w_{index}"
        )
        
    with col_pdf:
        doc_pdf = Document(caminho_modelo)
        substituir_texto_protegendo_logos(doc_pdf, dados_memorando)
        pdf_io = io.BytesIO()
        doc_pdf.save(pdf_io)
        pdf_io.seek(0)
        st.download_button(
            label="📕 PDF",
            data=pdf_io.getvalue(),
            file_name=f"{nome_base_arquivo}.pdf",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"p_{index}"
        )
        
    with col_email_btn:
        st.link_button(
            label="📧 E-mail",
            url=link_gmail,
            key=f"e_{index}"
        )
        
    st.markdown("---")

if arquivo_excel:
    df = pd.read_excel(arquivo_excel)
    
    # FIXADO: Puxa o nome exato da primeira coluna (Coluna A) para ser o índice da Notificação
    col_notif_forcada = df.columns[0]
    df.columns = df.columns.str.strip()
    data_extenso_envio = obter_data_por_extenso(data_selecionada)
    
    col_paciente = None
    col_data_notif = None
    col_data_ocorr = None
    col_turno = None
    col_tipo = None
    col_classif = None
    col_desc = None
    col_leito = None
    col_notificante = None
    col_sugestao = None
    col_localizacao = None
    col_email = None
    
    for col in df.columns:
        c_upper = col.upper()
        if "PACIENTE" in c_upper or "NOME" in c_upper: col_paciente = col
        elif "DATA" in c_upper and "NOTIF" in c_upper: col_data_notif = col
        elif "DATA" in c_upper and ("OCORR" in c_upper or "OCOR" in c_upper): col_data_ocorr = col
        elif "TURNO" in c_upper: col_turno = col
        elif "TIPO" in c_upper or "INCIDENTE" in c_upper: col_tipo = col
