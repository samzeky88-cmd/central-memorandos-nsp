import os
import pandas as pd
import streamlit as st
import io
import urllib.parse
from datetime import datetime
from docx import Document

# Biblioteca para gerar um arquivo PDF real e válido na nuvem (Linux)
try:
    from fpdf import FPDF
    FPDF_DISPONIVEL = True
except ImportError:
    FPDF_DISPONIVEL = False

st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
st.title("📄 Emissor de Memorandos Individuais - Hospital Dr. Jackson Lago")

st.markdown("### 🗓️ Configuração da Data de Envio")
data_selecionada = st.date_input("Selecione a data que sairá no cabeçalho do Memorando:", datetime.now())

arquivo_excel = st.file_uploader("Suba a planilha contendo os incidentes (.xlsx)", type=["xlsx"])
caminho_modelo = "modelo_memorando.docx"

def substituir_texto_protegendo_logos(doc, dicionario_tags):
    """Substitui o texto alterando os runs para proteger imagens e tabelas do modelo Word."""
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
        v_str = str(valor).strip().split(" ")
        if len(v_str) > 0 and "-" in v_str[0]:
            parts = v_str[0].split("-")
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
    """Evita a exibição da palavra 'nan' ou vazios nos campos de texto"""
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

def gerar_pdf_hospitalar_real(dados, nome_arquivo):
    """Gera um PDF estruturado em conformidade com o padrão visual do hospital"""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(20, 20, 20)
    
    pdf.set_font("Arial", "", 11)
    data_cabecalho = f"Sao Luis, {dados.get('{{data_envio}}')}"
    pdf.cell(0, 10, data_cabecalho, ln=True, align="R")
    pdf.ln(10)
    
    pdf.set_font("Arial", "", 11)
    texto_intro = "Prezado (a), vimos atraves deste comunicar que recebemos uma notificacao de incidente ocorrida neste setor. Segue abaixo as informacoes encaminhadas ao NSP:"
    pdf.multi_cell(0, 6, texto_intro)
    pdf.ln(5)
    
    itens_memorando = [
        ("DATA DA OCORRENCIA", dados.get("{{data_ocorrencia}}")),
        ("DATA DA NOTIFICACAO", dados.get("{{data_notificacao}}")),
        ("TURNO QUE OCORREU INCIDENTE", f"( {dados.get('{{m}}')} ) MANHA  ( {dados.get('{{t}}')} ) TARDE  ( {dados.get('{{n}}')} ) NOITE"),
        ("ONDE OCORREU INCIDENTE", dados.get("{{localizacao}}")),
        ("TIPO DE INCIDENTE", dados.get("{{tipo_incidente}}")),
        ("CLASSIFICACAO DO INCIDENTE", dados.get("{{classificacao_incidente}}")),
        ("DESCRICAO DA NOTIFICACAO", dados.get("{{descricao_notificacao}}")),
        ("PACIENTE", dados.get("{{nome_paciente}}")),
        ("LEITO", dados.get("{{leito}}")),
        ("SETOR NOTIFICANTE", dados.get("{{setor_notificante}}")),
        ("SUGESTAO/RECOMENDACAO", dados.get("{{sugestao}}"))
    ]
    
    for label, valor in itens_memorando:
        if valor:
            pdf.set_font("Arial", "B", 10)
            pdf.write(6, f"• {label}: ")
            pdf.set_font("Arial", "", 10)
            pdf.write(6, f"{valor}\n")
            pdf.ln(2)
            
    return pdf.output()

@st.fragment
def renderizar_linha_paciente_sob_demanda(index, linha, col_paciente, col_notif, col_data_notif, col_data_ocorr, col_turno, col_tipo, col_classif, col_desc, col_leito, col_notificante, col_sugestao, col_localizacao, col_email, data_extenso_envio):
    nome_do_paciente = tratar_str_limpa(linha[col_paciente])
    
    memo_01 = tratar_str_limpa(linha.get("Nº Memo 01", ""))
    if memo_01 == "":
        memo_01 = tratar_str_limpa(linha.get("MEMO", ""))
        
    memo_02 = tratar_str_limpa(linha.get("Nº Memo 02", ""))
    if memo_01 == "":
        num_memo_cru = memo_02 if memo_02 != "" else "S-N"
    else:
        num_memo_cru = memo_01
        
    num_notif = limpar_numero_float(linha.get(col_notif, "S-N"))
    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
    nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
    
    turno_planilha = tratar_str_limpa(linha.get(col_turno, "")).upper() if col_turno else ""
    marca_manha = "X" if "MANH" in turno_planilha else " "
    marca_tarde = "X" if "TARD" in turno_planilha else " "
    marca_noite = "X" if "NOIT" in turno_planilha else " "
    
    dados_memorando = {
        "{{numero_memorando}}": num_memo_cru,
        "{{gestor}}": tratar_str_limpa(linha.get("Gestor 01", "")),
        "{{setor}}": tratar_str_limpa(linha.get("SETOR NOTIFICADO", "")),
        "{{notificacao_n}}": num_notif,
        "{{data_notificacao}}": formatar_data_br(linha.get(col_data_notif, "")) if col_data_notif else "",
        "{{data_ocorrencia}}": formatar_data_br(linha.get(col_data_ocorr, "")) if col_data_ocorr else "",
        "{{localizacao}}": tratar_str_limpa(linha.get(col_localizacao, "")) if col_localizacao else "Ala B",
        "{{tipo_incidente}}": tratar_str_limpa(linha.get(col_tipo, "")),
        "{{classificacao_incidente}}": tratar_str_limpa(linha.get(col_classif, "")),
        "{{descricao_notificacao}}": tratar_str_limpa(linha.get(col_desc, "")),
        "{{nome_paciente}}": nome_do_paciente,
        "{{leito}}": limpar_numero_float(linha.get(col_leito, "")),
        "{{setor_notificante}}": tratar_str_limpa(linha.get(col_notificante, "")),
        "{{sugestao}}": tratar_str_limpa(linha.get(col_sugestao, "")),
        "{{m}}": marca_manha,
        "{{t}}": marca_tarde,
        "{{n}}": marca_noite,
        "{{data_envio}}": data_extenso_envio
    }
    
    if dados_memorando["{{setor_notificante}}"] == "":
        dados_memorando["{{setor_notificante}}"] = "NSP - NUCLEO DE SEGURANCA DO PACIENTE"

    email_destino = tratar_str_limpa(linha.get(col_email, "")) if col_email else ""
    hora_atual = datetime.now().hour
    saudacao = "Bom dia Prezados" if hora_atual < 12 else "Boa Tarde Prezados"
    texto_descricao = dados_memorando["{{descricao_notificacao}}"]
    
    texto_email_formatado = f"{saudacao},\n\nSegue em anexo o Memorando {num_memo_cru} referente à Notificação {num_notif}.\n\nPaciente: {nome_do_paciente}\nDescrição do ocorrido: {texto_descricao}"
    assunto_email = f"NSP - Memorando {num_memo_cru} (Notificação {num_notif})"
    
    link_gmail = (
        f"https://google.com?"
        f"view=cm&fs=1&tf=1"
        f"&to={urllib.parse.quote(email_destino)}"
        f"&su={urllib.parse.quote(assunto_email)}"
        f"&body={urllib.parse.quote(texto_email_formatado)}"
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
        if FPDF_DISPONIVEL:
            pdf_bytes = gerar_pdf_hospitalar_real(dados_memorando, nome_base_arquivo)
            st.download_button(
                label="📕 PDF",
                data=pdf_bytes,
                file_name=f"{nome_base_arquivo}.pdf",
                mime="application/pdf",
                key=f"p_{index}"
            )
        else:
            st.download_button(
                label="📕 PDF",
                data=word_io.getvalue(),
                file_name=f"{nome_base_arquivo}.pdf",
                mime="application/pdf",
                key=f"p_{index}"
            )
        
    with col_email_btn:
        st.link_button(
            label="📧 Enviar E-mail",
            url=link_gmail,
            key=f"e_{index}"
        )
        
    st.markdown("---")

