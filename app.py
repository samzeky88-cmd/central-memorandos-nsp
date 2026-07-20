import os
import pandas as pd
import streamlit as st
import io
import subprocess
from docx import Document

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

def converter_docx_para_pdf_via_sistema(caminho_docx):
    """Utiliza o LibreOffice do Linux para gerar um PDF idêntico ao Word sem corromper imagens."""
    try:
        # Executa a conversão headless silenciosa no Linux do Streamlit
        comando = ["soffice", "--headless", "--convert-to", "pdf", caminho_docx]
        subprocess.run(comando, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        
        nome_pdf_gerado = caminho_docx.replace(".docx", ".pdf")
        if os.path.exists(nome_pdf_gerado):
            with open(nome_pdf_gerado, "rb") as f:
                dados_pdf = f.read()
            os.remove(nome_pdf_gerado)  # Limpa o PDF temporário do servidor
            return dados_pdf
    except Exception as e:
        pass
    return None

if arquivo_excel:
    df = pd.read_excel(arquivo_excel)
    df.columns = df.columns.str.strip()
    
    # Identificação automática da coluna de paciente baseada na sua planilha
    coluna_paciente = None
    for col in df.columns:
        if col.upper() == "PACIENTE" or "PACIENTE" in col.upper() or "NOME" in col.upper():
            coluna_paciente = col
            break
            
    if not coluna_paciente:
        coluna_paciente = df.columns[0]
        
    # Limpeza de linhas inválidas ou fantasmas
    df = df.dropna(subset=[coluna_paciente])
    df = df[df[coluna_paciente].astype(str).str.strip() != ""]
    
    st.success(f"📋 Planilha validada! {len(df)} registros encontrados. Confira e baixe individualmente abaixo:")
    
    for index, linha in df.iterrows():
        nome_do_paciente = str(linha[coluna_paciente]).strip()
        
        # Extrai e limpa as variáveis para construir o nome do arquivo padronizado
        num_memo_cru = str(linha.get("numero_memorando", "S-N")).strip()
        num_memo = num_memo_cru.replace("NSP", "").replace("Memo", "").replace("MEMO", "").replace("/", "-").strip()
        num_notif = str(linha.get("notificacao_n", "S-N")).strip()
        
        # Nome do arquivo padronizado exatamente igual ao modelo solicitado
        nome_base_arquivo = f"MEMORANDO Nº {num_memo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
        
        # Inicia uma cópia limpa do Word em disco para a conversão
        doc = Document(caminho_modelo)
        
        # Lógica para marcar 'X' ou deixar espaço em branco nos turnos
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
        
        # Substituição cirúrgica preservando imagens do cabeçalho
        substituir_texto_protegendo_logos(doc, dados_memorando)
        
        # Salva o arquivo Word temporário para processamento do LibreOffice
        caminho_docx_temp = f"temp_{index}.docx"
        doc.save(caminho_docx_temp)
        
        # Carrega os bytes do Word criado para o botão correspondente
        with open(caminho_docx_temp, "rb") as f:
            word_bytes = f.read()
            
        # Converte para PDF nativo de forma paralela e estável
        pdf_bytes = converter_docx_para_pdf_via_sistema(caminho_docx_temp)
        
        # Deleta o arquivo Word temporário do servidor imediatamente
        if os.path.exists(caminho_docx_temp):
            os.remove(caminho_docx_temp)
            
        # Renderização organizada dos botões lado a lado na interface
        col_nome, col_word, col_pdf = st.columns([2, 1, 1])
        
        with col_nome:
            st.markdown(f"**🔹 {nome_do_paciente}**")
            
        with col_word:
            st.download_button(
                label="📝 Baixar WORD",
                data=word_bytes,
                file_name=f"{nome_base_arquivo}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"word_{index}"
            )
            
        with col_pdf:
            if pdf_bytes:
                st.download_button(
                    label="📕 Baixar PDF",
                    data=pdf_bytes,
                    file_name=f"{nome_base_arquivo}.pdf",
                    mime="application/pdf",
                    key=f"pdf_{index}"
                )
            else:
                st.warning("⚠️ Aguardando LibreOffice")
        st.divider()
