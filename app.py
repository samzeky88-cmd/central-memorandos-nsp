import os
import pandas as pd
import streamlit as st
import io
from docx import Document
from fpdf import FPDF

st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
st.title("📄 Emissor de Memorandos Individuais - Hospital Dr. Jackson Lago")

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
        if pd.isna(valor) or str(valor).strip() == "":
            return ""
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except:
        return str(valor)

def gerar_pdf_nativo_memoria(dados):
    """Gera um PDF estruturado de forma nativa e rápida em memória utilizando FPDF2."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_margins(20, 20, 20)
    
    # Configura fontes padrão seguras
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, "HOSPITAL DR. JACKSON LAGO", ln=True, align="C")
    pdf.cell(0, 8, "NÚCLEO DE SEGURANÇA DO PACIENTE - NSP", ln=True, align="C")
    pdf.ln(10)
    
    pdf.set_font("Helvetica", "", 11)
    
    # Cabeçalho do Memorando textual
    pdf.write(6, "MEMORANDO Nº: "); pdf.set_font("Helvetica", "B", 11); pdf.write(6, f"{dados['{{numero_memorando}}']}\n"); pdf.set_font("Helvetica", "", 11)
    pdf.write(6, "PARA: "); pdf.set_font("Helvetica", "B", 11); pdf.write(6, f"{dados['{{gestor}}']} {dados['{{setor}}']}\n"); pdf.set_font("Helvetica", "", 11)
    pdf.write(6, "NOTIFICAÇÃO Nº: "); pdf.set_font("Helvetica", "B", 11); pdf.write(6, f"{dados['{{notificacao_n}}']}\n"); pdf.set_font("Helvetica", "", 11)
    pdf.write(6, "DATA DA OCORRÊNCIA: "); pdf.set_font("Helvetica", "B", 11); pdf.write(6, f"{dados['{{data_ocorrencia}}']}\n"); pdf.set_font("Helvetica", "", 11)
    pdf.write(6, "PACIENTE: "); pdf.set_font("Helvetica", "B", 11); pdf.write(6, f"{dados['{{nome_paciente}}']}\n"); pdf.set_font("Helvetica", "", 11)
    pdf.ln(6)
    
    # Corpo do texto estruturado
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 6, "DESCRIÇÃO DA NOTIFICAÇÃO:", ln=True)
    pdf.set_font("Helvetica", "", 11)
    
    # Corrige problemas comuns de caracteres especiais (como acentos) no PDF nativo
    txt_descricao = str(dados['{{descricao_notificacao}}']).encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 6, txt_descricao)
    
    # Retorna o arquivo compilado em formato de bytes
    return pdf.output()

@st.fragment
def renderizar_linha_paciente_sob_demanda(index, linha, col_paciente, nome_col_notif):
    nome_do_paciente = str(linha[col_paciente]).strip()
    
    memo_01 = str(linha.get("Nº Memo 01", "")).strip()
    memo_02 = str(linha.get("Nº Memo 02", "")).strip()
    
    if memo_01 == "" or memo_01.lower() == "nan":
        num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" else "S-N"
    else:
        num_memo_cru = memo_01
        
    # Coleta de forma limpa o número da primeira coluna
    num_notif = str(linha.get(nome_col_notif, "S-N")).strip()
    if num_notif.lower() == "nan" or num_notif == "":
        num_notif = "S-N"
    
    # Padronização exata dos nomes dos arquivos para o seu computador
    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
    num_notif_limpo = num_notif.replace("Nº", "").replace(" ", "").strip()
    nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif_limpo}_I_NSP"
    
    turno_planilha = str(linha.get("turno", "")).strip().upper()
    marca_manha = "X" if "MANHÃ" in turno_planilha or "MANHA" in turno_planilha else " "
    marca_tarde = "X" if "TARDE" in turno_planilha else " "
    marca_noite = "X" if "NOITE" in turno_planilha else " "
    
    dados_memorando = {
        "{{numero_memorando}}": num_memo_cru,
        "{{gestor}}": str(linha.get("Gestor 01", "")),
        "{{setor}}": str(linha.get("SETOR NOTIFICADO", "")),
        "{{notificacao_n}}": num_notif,
        "{{data_notificacao}}": formatar_data_br(linha.get("data_notificacao", "")),
        "{{data_ocorrencia}}": formatar_data_br(linha.get("data_ocorrencia", "")),
        "{{localizacao}}": str(linha.get("localizacao", "")),
        "{{tipo_incidente}}": str(linha.get("tipo_incidente", "")),
        "{{classificacao_incidente}}": str(linha.get("classificacao_incidente", "")),
        "{{descricao_notificacao}}": str(linha.get("descricao_notificacao", "")),
        "{{nome_paciente}}": nome_do_paciente,
        "{{leito}}": str(linha.get("leito", "")),
        "{{setor_notificante}}": str(linha.get("setor_notificante", "")),
        "{{sugestao}}": str(linha.get("sugestao", "")),
        "{{m}}": marca_manha,
        "{{t}}": marca_tarde,
        "{{n}}": marca_noite
    }

    col_nome, col_word, col_pdf = st.columns(3)
    
    with col_nome:
        st.markdown(f"**🔹 {nome_do_paciente}**")
        
    with col_word:
        doc_word = Document(caminho_modelo)
        substituir_texto_protegendo_logos(doc_word, dados_memorando)
        word_io = io.BytesIO()
        doc_word.save(word_io)
        word_io.seek(0)
        
        st.download_button(
            label="📝 Baixar WORD",
            data=word_io.getvalue(),
            file_name=f"{nome_base_arquivo}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"w_{index}"
        )
        
    with col_pdf:
        # Geração instantânea e leve direto em memória Python (Não trava mais)
        pdf_data = gerar_pdf_nativo_memoria(dados_memorando)
        st.download_button(
            label="📕 Baixar PDF",
            data=bytes(pdf_data),
            file_name=f"{nome_base_arquivo}.pdf",
            mime="application/pdf",
            key=f"p_{index}"
        )
    st.markdown("---")

if arquivo_excel:
    df = pd.read_excel(arquivo_excel)
    
    # 🎯 CORRIGIDO: Captura o nome exato da primeira coluna da planilha (Coluna A) de forma isolada
    nome_col_notif = df.columns[0]
    
    df.columns = df.columns.str.strip()
    
    coluna_paciente = None
    for col in df.columns:
        c_upper = col.upper()
        if "PACIENTE" in c_upper or "NOME" in c_upper:
            coluna_paciente = col
            
    if not coluna_paciente: 
        coluna_paciente = df.columns[1]
        
    df = df.dropna(subset=[coluna_paciente])
    df = df[df[coluna_paciente].astype(str).str.strip() != ""]
    
    st.success(f"📋 Lista de verificação pronta! {len(df)} registros encontrados.")
    
    for index, linha in df.iterrows():
        renderizar_linha_paciente_sob_demanda(index, linha, coluna_paciente, nome_col_notif)
