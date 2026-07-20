import os
import pandas as pd
import streamlit as st
import io
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

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

def converter_docx_para_pdf_bytes(doc):
    """Extrai o texto estruturado do Word e compila em PDF de forma nativa e estável."""
    pdf_buffer = io.BytesIO()
    pdf = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
    story = []
    
    styles = getSampleStyleSheet()
    estilo_normal = ParagraphStyle('TextoNormal', parent=styles['Normal'], fontSize=11, leading=14)
    estilo_negrito = ParagraphStyle('TextoNegrito', parent=styles['Normal'], fontSize=11, leading=14, fontName='Helvetica-Bold')
    
    # Adiciona o conteúdo textual preservando parágrafos simples
    for paragrafo in doc.paragraphs:
        texto = paragrafo.text.strip()
        if texto:
            if "MEMO:" in texto or "DE:" in texto or "PARA:" in texto or "ASSUNTO:" in texto:
                story.append(Paragraph(texto, estilo_negrito))
            else:
                story.append(Paragraph(texto, estilo_normal))
            story.append(Spacer(1, 8))
            
    pdf.build(story)
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()

if arquivo_excel:
    df = pd.read_excel(arquivo_excel)
    df.columns = df.columns.str.strip()
    
    # Identificação automática da coluna de paciente
    coluna_paciente = None
    for col in df.columns:
        if col.upper() == "PACIENTE" or "PACIENTE" in col.upper() or "NOME" in col.upper():
            coluna_paciente = col
            break
            
    if not coluna_paciente:
        coluna_paciente = df.columns[0]
        
    df = df.dropna(subset=[coluna_paciente])
    df = df[df[coluna_paciente].astype(str).str.strip() != ""]
    
    st.success(f"📋 Planilha validada! {len(df)} registros encontrados. Confira e baixe individualmente abaixo:")
    
    # Cria uma lista limpa e organizada na tela
    for index, linha in df.iterrows():
        nome_do_paciente = str(linha[coluna_paciente]).strip()
        
        # Extrai variáveis do nome do arquivo padronizado
        num_memo_cru = str(linha.get("numero_memorando", "S-N")).strip()
        num_memo = num_memo_cru.replace("NSP", "").replace("Memo", "").replace("MEMO", "").replace("/", "-").strip()
        
        num_notif = str(linha.get("notificacao_n", "S-N")).strip()
        
        # Estrutura o nome exatamente igual à sua imagem de referência
        nome_base_arquivo = f"MEMORANDO Nº {num_memo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
        
        # Processamento individual em memória
        doc = Document(caminho_modelo)
        
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
        
        substituir_texto_protegendo_logos(doc, dados_memorando)
        
        # Transforma o arquivo WORD modificado em memória para download imediato
        word_buffer = io.BytesIO()
        doc.save(word_buffer)
        word_buffer.seek(0)
        
        # Renderiza o layout de colunas no Streamlit para os botões ficarem lado a lado
        col_nome, col_word, col_pdf = st.columns([3, 1, 1])
        
        with col_nome:
            st.markdown(f"**🔹 {nome_do_paciente}**")
            
        with col_word:
            st.download_button(
                label="📝 Baixar WORD",
                data=word_buffer.getvalue(),
                file_name=f"{nome_base_arquivo}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"word_{index}"
            )
            
        with col_pdf:
            try:
                pdf_bytes = converter_docx_para_pdf_bytes(doc)
                st.download_button(
                    label="📕 Baixar PDF",
                    data=pdf_bytes,
                    file_name=f"{nome_base_arquivo}.pdf",
                    mime="application/pdf",
                    key=f"pdf_{index}"
                )
            except Exception as e:
                st.error("Erro ao gerar PDF")
        st.divider()
