import os
import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄")
st.title("📄 Emissor de Memorandos - Hospital Dr. Jackson Lago")

arquivo_excel = st.file_uploader("Suba a planilha contendo os incidentes (.xlsx)", type=["xlsx"])
caminho_modelo = "modelo_memorando.docx"

def substituir_texto_protegendo_logos(doc, dicionario_tags):
    """
    Substitui o texto alterando apenas os 'runs' (fragmentos de texto puro).
    Isso impede totalmente que o python-docx toque nas imagens ou no cabeçalho.
    """
    # 1. Procura nos parágrafos normais do documento
    for paragrafo in doc.paragraphs:
        for tag, valor in dicionario_tags.items():
            if tag in paragrafo.text:
                for run in paragrafo.runs:
                    if tag in run.text:
                        run.text = run.text.replace(tag, valor)
                        
    # 2. Procura dentro de tabelas (caso existam no documento)
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
    # Lê a planilha tratando os nomes das colunas sem espaços extras nas pontas
    df = pd.read_excel(arquivo_excel)
    df.columns = df.columns.str.strip()
    
    # Remove as linhas sem paciente para evitar criar documentos vazios
    df = df.dropna(subset=["nome_paciente"])
    
    st.success(f"📋 Planilha validada com sucesso! {len(df)} memorandos prontos.")
    
    if st.button("🚀 Gerar Memorandos"):
        for index, linha in df.iterrows():
            doc = Document(caminho_modelo)
            
            # Lógica para marcar 'X' ou deixar espaço em branco nos turnos
            turno_planilha = str(linha.get("turno", "")).strip().upper()
            marca_manha = "X" if "MANHÃ" in turno_planilha or "MANHA" in turno_planilha else " "
            marca_tarde = "X" if "TARDE" in turno_planilha else " "
            marca_noite = "X" if "NOITE" in turno_planilha else " "
            
            # Formatação segura das datas lidas da planilha
            dt_notif = formatar_data_br(linha.get("data_notificacao", ""))
            dt_ocorr = formatar_data_br(linha.get("data_ocorrencia", ""))
            
            # Dicionário mapeando com precisão absoluta as tags da sua imagem
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
                "{{nome_paciente}}": str(linha.get("nome_paciente", "")),
                "{{leito}}": str(linha.get("leito", "")),
                "{{setor_notificante}}": str(linha.get("setor_notificante", "")),
                "{{sugestao}}": str(linha.get("sugestao", "")),
                
                # Tags de marcação dos turnos entre parênteses da imagem
                "{{m}}": marca_manha,
                "{{t}}": marca_tarde,
                "{{n}}": marca_noite
            }
            
            # Executa a substituição sem apagar as logos corporativas
            substituir_texto_protegendo_logos(doc, dados_memorando)
            
            # Salva o arquivo final de maneira estruturada
            num_memo = str(linha.get("numero_memorando", f"S-N-{index}")).replace("/", "-")
            nome_arquivo_final = f"Memo_{num_memo}_{linha['nome_paciente']}.docx"
            doc.save(nome_arquivo_final)
            
            # Apresenta o botão de download personalizado na interface
            with open(nome_arquivo_final, "rb") as f:
                st.download_button(
                    label=f"📥 Baixar Memorando - {linha['nome_paciente']}",
                    data=f,
                    file_name=nome_arquivo_final,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            # Remove o arquivo do servidor local imediatamente após disponibilizar o download
            if os.path.exists(nome_arquivo_final):
                os.remove(nome_arquivo_final)
                
        st.balloons()
