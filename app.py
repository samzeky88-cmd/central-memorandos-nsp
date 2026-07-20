import os
import pandas as pd
import streamlit as st
import io
import zipfile
from docx import Document

st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄")
st.title("📄 Emissor de Memorandos - Hospital Dr. Jackson Lago")

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

if arquivo_excel:
    df = pd.read_excel(arquivo_excel)
    
    # Remove espaços invisíveis das colunas
    df.columns = df.columns.str.strip()
    
    # Identificação automática da coluna de paciente
    coluna_paciente = None
    for col in df.columns:
        if col.upper() == "PACIENTE" or "PACIENTE" in col.upper() or "NOME" in col.upper():
            coluna_paciente = col
            break
            
    if not coluna_paciente:
        coluna_paciente = df.columns[0]
        
    # Limpa linhas inválidas
    df = df.dropna(subset=[coluna_paciente])
    df = df[df[coluna_paciente].astype(str).str.strip() != ""]
    
    st.success(f"📋 Planilha validada com sucesso! {len(df)} memorandos prontos. (Coluna identificada: '{coluna_paciente}')")
    
    if st.button("🚀 Compactar e Gerar Todos os Memorandos (.ZIP)"):
        # Cria um arquivo ZIP na memória do computador para juntar os arquivos
        arquivo_zip_memoria = io.BytesIO()
        
        # Barra de progresso visual para você acompanhar os 113 documentos
        barra_progresso = st.progress(0)
        status_texto = st.empty()
        
        with zipfile.ZipFile(arquivo_zip_memoria, "w", zipfile.ZIP_DEFLATED) as zip_file:
            for index, linha in df.iterrows():
                # Atualiza a barra de progresso na tela
                percentual_concluido = (index + 1) / len(df)
                barra_progresso.progress(percentual_concluido)
                
                doc = Document(caminho_modelo)
                nome_do_paciente = str(linha[coluna_paciente]).strip()
                status_texto.text(f"Processando ({index + 1}/{len(df)}): {nome_do_paciente}")
                
                # Lógica para marcar 'X' nos turnos
                turno_planilha = str(linha.get("turno", "")).strip().upper()
                marca_manha = "X" if "MANHÃ" in turno_planilha or "MANHA" in turno_planilha else " "
                marca_tarde = "X" if "TARDE" in turno_planilha else " "
                marca_noite = "X" if "NOITE" in turno_planilha else " "
                
                # Tratamento de datas brasileiras
                dt_notif = formatar_data_br(linha.get("data_notificacao", ""))
                dt_ocorr = formatar_data_br(linha.get("data_ocorrencia", ""))
                
                # Mapeamento completo das tags do Word
                dados_memorando = {
                    "{{numero_memorando}}": str(linha.get("numero_memorando", "")),
                    "{{gestor}}": str(linha.get("Gestor 01", "")),
                    "{{setor}}": str(linha.get("SETOR NOTIFICADO", "")),
                    "{{notificacao_n}}": str(linha.get("notificacao_n", "")),
                    "{{data_notificacao}}": dt_notif,
                    "{{data_ocorrencia}}": dt_ocorr,
                    "{{localizacao}}": str(linha.get("localizacao", "")),
                    "{{tipo_incidente}}": str(linha.get("tipo_incidente", "")),
                    "{{classificacao_incidente}}": str(linha.get("classificacao_incidente", "")),
                    "{{descricao_notificacao}}": str(linha.get("descricao_notificacao", "")),
                    "{{nome_paciente}}": nome_do_paciente,
                    "{{leito}}": str(linha.get("leito", "")),
                    "{{setor_notificante}}": str(linha.get("setor_notificante", "")),
                    "{{sugestao}}": str(linha.get("sugestao", "")),
                    
                    # Tags dos turnos
                    "{{m}}": marca_manha,
                    "{{t}}": marca_tarde,
                    "{{n}}": marca_noite
                }
                
                # Aplica as alterações sem afetar os logotipos corporativos
                substituir_texto_protegendo_logos(doc, dados_memorando)
                
                # Salva o arquivo temporário diretamente na memória
                num_memo = str(linha.get("numero_memorando", f"S-N-{index}")).replace("/", "-")
                nome_arquivo_word = f"Memo_{num_memo}_{nome_do_paciente}.docx"
                
                # Converte o documento para bytes e adiciona ao ZIP
                doc_stream = io.BytesIO()
                doc.save(doc_stream)
                doc_stream.seek(0)
                zip_file.writestr(nome_arquivo_word, doc_stream.getvalue())
        
        # Prepara o arquivo final compactado para o download
        arquivo_zip_memoria.seek(0)
        status_texto.text("🎉 Todos os 113 memorandos foram empacotados com sucesso!")
        
        # Cria O ÚNICO botão de download na tela
        st.download_button(
            label="📥 BAIXAR TODOS OS 113 MEMORANDOS (.ZIP)",
            data=arquivo_zip_memoria,
            file_name="Todos_Os_Memorandos_NSP.zip",
            mime="application/zip",
            use_container_width=True
        )
        st.balloons()
