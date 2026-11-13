# =================================================================
# 📄 GERADOR DE MEMORANDOS — HOSPITAL DR. JACKSON LAGO
# Versão Completa e Fechada — 100% Funcional
# =================================================================

import os
import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document

# =================================================================
# 📚 LISTA DE E-MAILS DOS SETORES
# =================================================================
LISTA_SETORES = [
    ("DIREÇÃO", "direcao.hospitaldacidade@gmail.com"),
    ("GABINETE DA DIREÇÃO", "gabinete.hospitaldacidade@gmail.com"),
    ("ALA A - ENFERMAGEM", "alaaclinicamedica@gmail.com"),
    ("ALA B - ENFERMAGEM", "alab.hcid@gmail.com"),
    ("NSP - NÚCLEO DE SEGURANÇA DO PACIENTE", "nsp.hcid@gmail.com"),
]

df_setores = pd.DataFrame(LISTA_SETORES, columns=["SETOR", "E-MAIL"])
df_setores = df_setores[df_setores["E-MAIL"].str.strip() != ""].reset_index(drop=True)

# =================================================================
# 🎯 CONFIGURAÇÃO DA PÁGINA
# =================================================================
st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide")
st.title("📝 Emissor de Memorandos — NSP / HCID")

data_selecionada = st.date_input("Data de Envio", datetime.now())
arquivo_excel = st.file_uploader("Planilha com Incidentes (.xlsx)", type=["xlsx"])
caminho_modelo = "modelo_memorando.docx"

# =================================================================
# 🛠️ FUNÇÕES AUXILIARES
# =================================================================
def substituir_texto_no_docx(doc, dados):
    for p in doc.paragraphs:
        for chave, valor in dados.items():
            if chave in p.text:
                for run in p.runs:
                    if chave in run.text:
                        run.text = run.text.replace(chave, str(valor))
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                substituir_texto_no_docx(celula, dados)

def formatar_data_br(valor):
    try:
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan":
            return ""
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except:
        return str(valor).strip()

def limpar_numero(valor):
    if pd.isna(valor) or str(valor).strip().lower() == "nan":
        return ""
    try:
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        return str(valor).strip()
    except:
        return str(valor)

def data_por_extenso(dt):
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    return f"{dt.day} de {meses[dt.month-1]} de {dt.year}"

# =================================================================
# 📄 FUNÇÃO PRINCIPAL DE GERAÇÃO
# =================================================================
def gerar_e_exibir_documento(dados, nome_arquivo_base):
    try:
        with open(caminho_modelo, "rb") as f:
            doc = Document(f)
        substituir_texto_no_docx(doc, dados)
        
        saida_docx = io.BytesIO()
        doc.save(saida_docx)
        saida_docx.seek(0)
        docx_bytes = saida_docx.getvalue()
        
        st.download_button(
            label="📄 Baixar Memorando (DOCX)",
            data=docx_bytes,
            file_name=f"{nome_arquivo_base}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        return True
    except FileNotFoundError:
        st.error("❌ Arquivo modelo_memorando.docx não encontrado! Verifique se ele está na pasta do projeto.")
        return False
    except Exception as e:
        st.error(f"❌ Erro ao gerar documento: {str(e)}")
        return False

# =================================================================
# 🚀 EXECUÇÃO PRINCIPAL
# =================================================================
if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel, engine="openpyxl")
        st.success(f"✅ Planilha carregada com sucesso — {len(df)} registro(s) encontrado(s)")
        data_extenso = data_por_extenso(data_selecionada)
        
        for idx, linha in df.iterrows():
            num_notif = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
            if not num_notif or num_notif.upper() == "STATUS" or "NOTIF" in str(num_notif).upper():
                continue
            
            nome_paciente = str(linha.iloc[9]).strip() if len(linha) > 9 else "Paciente"
            if nome_paciente.upper() == "PACIENTE":
                continue
            
            num_memo = limpar_numero(linha.iloc[15]) if len(linha) > 15 else "S/N"
            dt_notif = formatar_data_br(linha.iloc[3]) if len(linha) > 3 else ""
            dt_ocorr = formatar_data_br(linha.iloc[2]) if len(linha) > 2 else ""
            local = str(linha.iloc[5]).strip() if len(linha) > 5 else ""
            tipo = str(linha.iloc[6]).strip() if len(linha) > 6 else ""
            desc = str(linha.iloc[8]).strip() if len(linha) > 8 else ""
            leito = limpar_numero(linha.iloc[10]) if len(linha) > 10 else ""
            setor = str(linha.iloc[14]).strip() if len(linha) > 14 else ""
            gestor = str(linha.iloc[13]).strip() if len(linha) > 13 else "GESTOR DE ENFERMAGEM"
            sugestao = str(linha.iloc[12]).strip() if len(linha) > 12 else ""
            
            dados_memorando = {
                "{{numero_memorando}}": num_memo,
                "{{gestor}}": gestor,
                "{{setor}}": setor,
                "{{notificacao_n}}": num_notif,
                "{{data_notificacao}}": dt_notif,
                "{{data_ocorrencia}}": dt_ocorr,
                "{{localizacao}}": local,
                "{{tipo_incidente}}": tipo,
                "{{descricao_notificacao}}": desc,
                "{{nome_paciente}}": nome_paciente,
                "{{leito}}": leito,
                "{{sugestao}}": sugestao,
                "{{data_envio}}": data_extenso,
            }
            
            with st.expander(f"📋 Notificação {num_notif} — {nome_paciente}"):
                st.write(f"**Memorando:** {num_memo} | **Data:** {data_extenso}")
                st.write(f"**Setor:** {setor} | **Gestor:** {gestor}")
                gerar_e_exibir_documento(dados_memorando, f"MEMORANDO_{num_notif}")
    
    except Exception as e:
        st.error(f"❌ Erro ao processar a planilha: {str(e)}")
else:
    st.info("👆 Por favor, faça o upload da sua planilha Excel (.xlsx) para começar")
