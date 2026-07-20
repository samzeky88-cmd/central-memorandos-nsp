import os
import pandas as pd
import streamlit as st
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
    
    # Remove espaços invisíveis que possam atrapalhar o nome das colunas
    df.columns = df.columns.str.strip()
    
    # 🔍 Identificação precisa e segura da coluna do paciente baseado na sua imagem
    coluna_paciente = None
    for col in df.columns:
        if col.upper() == "PACIENTE" or "PACIENTE" in col.upper() or "NOME" in col.upper():
            coluna_paciente = col
            break
            
    # Se por acaso a planilha vier totalmente desconfigurada, usa a primeira coluna
    if not coluna_paciente:
        coluna_paciente = df.columns[0]
        
    # Remove linhas onde o nome do paciente esteja vazio ou nulo (evita criar arquivo fantasma)
    df = df.dropna(subset=[coluna_paciente])
    df = df[df[coluna_paciente].astype(str).str.strip() != ""]
    
    st.success(f"📋 Planilha validada com sucesso! {len(df)} memorandos prontos. (Coluna identificada: '{coluna_paciente}')")
    
    if st.button("🚀 Gerar Memorandos"):
        for index, linha in df.iterrows():
            doc = Document(caminho_modelo)
            
            # Puxa o nome do paciente de forma segura
            nome_do_paciente = str(linha[coluna_paciente]).strip()
            
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
                "{{nome_paciente}}": nome_do_paciente,
                "{{leito}}": str(linha.get("leito", "")),
                "{{setor_notificante}}": str(linha.get("setor_notificante", "")),
                "{{sugestao}}": str(linha.get("sugestao", "")),
                
                # Tags de marcação dos turnos
                "{{m}}": marca_manha,
                "{{t}}": marca_tarde,
                "{{n}}": marca_noite
            }
            
            # Executa a substituição segura sem remover os logotipos do topo
            substituir_texto_protegendo_logos(doc, dados_memorando)
            
            # Salva o arquivo final
            num_memo = str(linha.get("numero_memorando", f"S-N-{index}")).replace("/", "-")
            nome_arquivo_final = f"Memo_{num_memo}_{nome_do_paciente}.docx"
            doc.save(nome_arquivo_final)
            
            # Disponibiliza o botão de download
            with open(nome_arquivo_final, "rb") as f:
                st.download_button(
                    label=f"📥 Baixar Memorando - {nome_do_paciente}",
                    data=f,
                    file_name=nome_arquivo_final,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            # Remove o arquivo temporário do servidor
            if os.path.exists(nome_arquivo_final):
                os.remove(nome_arquivo_final)
                
        st.balloons()
