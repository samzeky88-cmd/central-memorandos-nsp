import os
import pandas as pd
import streamlit as st
import io
import subprocess
import time
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

def converter_docx_para_pdf_via_sistema(caminho_docx, index):
    """Utiliza o LibreOffice de forma isolada e robusta para evitar erros de permissão e concorrência."""
    try:
        # Define uma pasta de perfil temporária exclusiva para esta conversão não travar no Linux
        user_profile_dir = f"/tmp/libreoffice_profile_{index}_{int(time.time())}"
        
        comando = [
            "soffice",
            f"-env:UserInstallation=file://{user_profile_dir}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            "/tmp",
            caminho_docx
        ]
        
        subprocess.run(comando, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True, timeout=30)
        
        # O LibreOffice joga o PDF convertido na pasta /tmp
        nome_pdf_gerado = os.path.join("/tmp", os.path.basename(caminho_docx).replace(".docx", ".pdf"))
        
        if os.path.exists(nome_pdf_gerado):
            with open(nome_pdf_gerado, "rb") as f:
                dados_pdf = f.read()
            os.remove(nome_pdf_gerado)
            
            # Limpa o perfil temporário criado do LibreOffice
            subprocess.run(["rm", "-rf", user_profile_dir], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return dados_pdf
    except Exception as e:
        pass
    return None

@st.fragment
def renderizar_linha_paciente_sob_demanda(index, linha, col_paciente, col_notif):
    nome_do_paciente = str(linha[col_paciente]).strip()
    
    memo_01 = str(linha.get("Nº Memo 01", "")).strip()
    memo_02 = str(linha.get("Nº Memo 02", "")).strip()
    
    if memo_01 == "" or memo_01.lower() == "nan":
        num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" else "S-N"
    else:
        num_memo_cru = memo_01
        
    num_notif = str(linha.get(col_notif, "S-N")).strip()
    
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

    col_nome, col_word, col_pdf = st.columns()
    
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
        if st.button("⚡ Gerar PDF", key=f"p_btn_{index}"):
            with st.spinner("Compilando PDF..."):
                doc_pdf = Document(caminho_modelo)
                substituir_texto_protegendo_logos(doc_pdf, dados_memorando)
                
                # Salva o arquivo temporário diretamente na pasta /tmp do Linux
                caminho_temp = f"/tmp/doc_{index}_{int(time.time())}.docx"
                doc_pdf.save(caminho_temp)
                
                pdf_bytes = converter_docx_para_pdf_via_sistema(caminho_temp, index)
                
                if os.path.exists(caminho_temp):
                    os.remove(caminho_temp)
                    
                if pdf_bytes:
                    st.download_button(
                        label="📥 Salvar PDF",
                        data=pdf_bytes,
                        file_name=f"{nome_base_arquivo}.pdf",
                        mime="application/pdf",
                        key=f"dl_pdf_{index}"
                    )
                else:
                    st.error("Erro na conversão. Tente novamente.")
    st.markdown("---")

if arquivo_excel:
    df = pd.read_excel(arquivo_excel)
    df.columns = df.columns.str.strip()
    
    coluna_paciente = None
    coluna_notificacao = None
    
    for col in df.columns:
        c_upper = col.upper()
        if "PACIENTE" in c_upper or "NOME" in c_upper:
            coluna_paciente = col
        elif col == "Nº" or "NOTIF" in c_upper:
            coluna_notificacao = col
            
    if not coluna_paciente: 
        coluna_paciente = df.columns
    if not coluna_notificacao: 
        coluna_notificacao = "Nº"
        
    df = df.dropna(subset=[coluna_paciente])
    df = df[df[coluna_paciente].astype(str).str.strip() != ""]
    
    st.success(f"📋 Lista de verificação pronta! {len(df)} registros encontrados.")
    
    for index, linha in df.iterrows():
        renderizar_linha_paciente_sob_demanda(index, linha, coluna_paciente, coluna_notificacao)
