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
    """Utiliza o LibreOffice do Linux para gerar um PDF idêntico ao Word."""
    try:
        comando = ["soffice", "--headless", "--convert-to", "pdf", caminho_docx]
        subprocess.run(comando, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
        
        nome_pdf_gerado = caminho_docx.replace(".docx", ".pdf")
        if os.path.exists(nome_pdf_gerado):
            with open(nome_pdf_gerado, "rb") as f:
                dados_pdf = f.read()
            os.remove(nome_pdf_gerado)
            return dados_pdf
    except Exception as e:
        pass
    return None

if arquivo_excel:
    # Evita reprocessar a planilha toda vez que você clica em um botão
    if "dados_processados" not in st.session_state:
        with st.spinner("📦 Preparando e gerando arquivos de verificação... Aguarde um instante."):
            df = pd.read_excel(arquivo_excel)
            df.columns = df.columns.str.strip()
            
            # 🔍 PROCURA AS COLUNAS COM BASE NAS TRÊS IMAGENS ENVIADAS
            coluna_paciente = None
            coluna_notificacao = None
            
            for col in df.columns:
                c_upper = col.upper()
                if "PACIENTE" in c_upper or "NOME" in c_upper:
                    coluna_paciente = col
                elif col == "Nº" or "NOTIF" in c_upper:
                    coluna_notificacao = col
            
            # Garantia caso falte a identificação automática
            if not coluna_paciente: coluna_paciente = df.columns
            if not coluna_notificacao: coluna_notificacao = "Nº"
            
            df = df.dropna(subset=[coluna_paciente])
            df = df[df[coluna_paciente].astype(str).str.strip() != ""]
            
            lista_final = []
            
            for index, Board in df.iterrows():
                nome_do_paciente = str(Board[coluna_paciente]).strip()
                
                # 🔄 ESTRATÉGIA DE BUSCA NAS DUAS COLUNAS (Memo 01 e Memo 02)
                memo_01 = str(Board.get("Nº Memo 01", "")).strip()
                memo_02 = str(Board.get("Nº Memo 02", "")).strip()
                
                # Se a coluna 01 estiver vazia ou for nan, ele assume o valor da coluna 02
                if memo_01 == "" or memo_01.lower() == "nan":
                    num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" else "S-N"
                else:
                    num_memo_cru = memo_01
                
                num_notif = str(Board.get(coluna_notificacao, "S-N")).strip()
                
                # Limpa os textos para gerar o nome do arquivo padronizado
                num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
                num_notif_limpo = num_notif.replace("Nº", "").replace(" ", "").strip()
                
                nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif_limpo}_I_NSP"
                
                doc = Document(caminho_modelo)
                
                turno_planilha = str(Board.get("turno", "")).strip().upper()
                marca_manha = "X" if "MANHÃ" in turno_planilha or "MANHA" in turno_planilha else " "
                marca_tarde = "X" if "TARDE" in turno_planilha else " "
                marca_noite = "X" if "NOITE" in turno_planilha else " "
                
                # Passa o número correto para preencher dentro do documento Word corporativo
                dados_memorando = {
                    "{{numero_memorando}}": num_memo_cru,
                    "{{gestor}}": str(Board.get("Gestor 01", "")),
                    "{{setor}}": str(Board.get("SETOR NOTIFICADO", "")),
                    "{{notificacao_n}}": num_notif,
                    "{{data_notificacao}}": formatar_data_br(Board.get("data_notificacao", "")),
                    "{{data_ocorrencia}}": formatar_data_br(Board.get("data_ocorrencia", "")),
                    "{{localizacao}}": str(Board.get("localizacao", "")),
                    "{{tipo_incidente}}": str(Board.get("tipo_incidente", "")),
                    "{{classificacao_incidente}}": str(Board.get("classificacao_incidente", "")),
                    "{{descricao_notificacao}}": str(Board.get("descricao_notificacao", "")),
                    "{{nome_paciente}}": nome_do_paciente,
                    "{{leito}}": str(Board.get("leito", "")),
                    "{{setor_notificante}}": str(Board.get("setor_notificante", "")),
                    "{{sugestao}}": str(Board.get("sugestao", "")),
                    "{{m}}": marca_manha,
                    "{{t}}": marca_tarde,
                    "{{n}}": marca_noite
                }
                
                substituir_texto_protegendo_logos(doc, dados_memorando)
                
                caminho_docx_temp = f"temp_{index}.docx"
                doc.save(caminho_docx_temp)
                
                with open(caminho_docx_temp, "rb") as f:
                    word_bytes = f.read()
                    
                pdf_bytes = converter_docx_para_pdf_via_sistema(caminho_docx_temp)
                
                if os.path.exists(caminho_docx_temp):
                    os.remove(caminho_docx_temp)
                    
                lista_final.append({
                    "paciente": nome_do_paciente,
                    "nome_arquivo": nome_base_arquivo,
                    "word": word_bytes,
                    "pdf": pdf_bytes
                })
                
            st.session_state["dados_processados"] = lista_final

    # Exibe a lista na tela usando o cache fixado
    if "dados_processados" in st.session_state:
        st.success(f"📋 Lista prontas! {len(st.session_state['dados_processados'])} registros encontrados.")
        
        for idx, item in enumerate(st.session_state["dados_processados"]):
            col_nome, col_word, col_pdf = st.columns()
            
            with col_nome:
                st.markdown(f"**🔹 {item['paciente']}**")
                
            with col_word:
                st.download_button(
                    label="📝 Baixar WORD",
                    data=item["word"],
                    file_name=f"{item['nome_arquivo']}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"w_btn_{idx}"
                )
                
            with col_pdf:
                if item["pdf"]:
                    st.download_button(
                        label="📕 Baixar PDF",
                        data=item["pdf"],
                        file_name=f"{item['nome_arquivo']}.pdf",
                        mime="application/pdf",
                        key=f"p_btn_{idx}"
                    )
                else:
                    st.error("Falha ao gerar PDF")
            st.markdown("---")
