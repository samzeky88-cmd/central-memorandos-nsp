import os 
import pandas as pd 
import streamlit as st 
import io 
import urllib.parse 
from datetime import datetime, timedelta, timezone 
from docx import Document 

st.set_page_config(page_title="Gerador de Memorandos", page_icon="📄", layout="wide") 
st.title("📄 Emissor de Memorandos Individuais - Hospital Dr. Jackson Lago") 
st.markdown("### 🗓️ Configuração da Data de Envio") 

data_selecionada = st.date_input("Selecione a data que sairá no cabeçalho do Memorando:", datetime.now()) 
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
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan": 
            return "" 
        return pd.to_datetime(valor).strftime("%d/%m/%Y") 
    except: 
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
    """Evita que campos vazios ou nulos exibam a palavra 'nan'""" 
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

@st.fragment 
def renderizar_linha_paciente_sob_demanda(index, linha, num_colunas, data_extenso_envio): 
    num_notif = limpar_numero_float(linha.iloc[0]) if num_colunas > 0 else "S-N" 
    if num_notif.upper() == "STATUS" or "NOTIF" in num_notif.upper() or num_notif == "" or num_notif == "1": 
        return 
    dt_ocorr = formatar_data_br(linha.iloc[2]) if num_colunas > 2 else "" 
    dt_notif = formatar_data_br(linha.iloc[3]) if num_colunas > 3 else "" 
    turno_planilha = str(linha.iloc[4]).strip().upper() if num_colunas > 4 else "" 
    onde_ocorreu = tratar_str_limpa(linha.iloc[5]) if num_colunas > 5 else "Ala B" 
    tipo_incidente = tratar_str_limpa(linha.iloc[6]) if num_colunas > 6 else "" 
    classificacao_incidente = tratar_str_limpa(linha.iloc[7]) if num_colunas > 7 else "" 
    descricao_notificacao = tratar_str_limpa(linha.iloc[8]) if num_colunas > 8 else "" 
    nome_do_paciente = tratar_str_limpa(linha.iloc[9]) if num_colunas > 9 else "Paciente Não Identificado" 
    if nome_do_paciente.upper() == "PACIENTE": 
        return 
    leito_paciente = limpar_numero_float(linha.iloc[10]) if num_colunas > 10 else "" 
    setor_notificante_bruto = tratar_str_limpa(linha.iloc[11]) if num_colunas > 11 else "" 
    sugestao_nsp = tratar_str_limpa(linha.iloc[12]) if num_colunas > 12 else "" 
    gestor_destinatario = tratar_str_limpa(linha.iloc[13]) if num_colunas > 13 else "GESTOR DE ENFERMAGEM" 
    setor_notificado_final = tratar_str_limpa(linha.iloc[14]) if num_colunas > 14 else onde_ocorreu 
    num_memo_cru = tratar_str_limpa(linha.iloc[15]) if num_colunas > 15 else "" 
    if num_memo_cru == "" or num_memo_cru.upper() == "Nº MEMO 01": 
        num_memo_cru = "S-N" 
    email_destino = tratar_str_limpa(linha.iloc[21]) if num_colunas > 21 else "" 
    if email_destino.upper() == "EMAIL_SETOR": 
        return 
    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip() 
    nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP" 
    marca_manha = "X" if "MANH" in turno_planilha else " " 
    marca_tarde = "X" if "TARD" in turno_planilha else " " 
    marca_noite = "X" if "NOIT" in turno_planilha else " " 
    if setor_notificante_bruto == "": 
        setor_notificante_bruto = "NSP - NÚCLEO DE SEGURANÇA DO PACIENTE" 
    dados_memorando = { 
        "{{numero_memorando}}": num_memo_cru, 
        "{{gestor}}": gestor_destinatario, 
        "{{setor}}": setor_notificado_final, 
        "{{notificacao_n}}": num_notif, 
        "{{data_notificacao}}": dt_notif, 
        "{{data_ocorrencia}}": dt_ocorr, 
        "{{localizacao}}": onde_ocorreu, 
        "{{tipo_incidente}}": tipo_incidente, 
        "{{classificacao_incidente}}": classificacao_incidente, 
        "{{descricao_notificacao}}": descricao_notificacao, 
        "{{nome_paciente}}": nome_do_paciente, 
        "{{leito}}": leito_paciente, 
        "{{setor_notificante}}": setor_notificante_bruto, 
        "{{sugestao}}": sugestao_nsp, 
        "{{m}}": marca_manha, 
        "{{t}}": marca_tarde, 
        "{{n}}": marca_noite, 
        "{{data_envio}}": data_extenso_envio 
    } 
    fuso_brasilia = timezone(timedelta(hours=-3)) 
    hora_atual = datetime.now(fuso_brasilia).hour 
    saudacao = "Bom Dia Prezados" if hora_atual < 12 else "Boa Tarde Prezados" 
    corpo_email = ( 
        f"{saudacao},\n\n" 
        f"Segue em Anexo o Memorando Nº {num_memo_cru} para ser analisado e respondido " 
        f"(via e-mail) em até 15 dias após a data presente.\n\n" 
        f"ATENÇÃO: A resposta via e-mail deve constar um arquivo em forma de word ou PDF para " 
        f"arquivamento de respostas conforme rotina institutional. Não serão aceitas mensagens " 
        f"via e-mail sem arquivo como resposta.\n\n" 
        f"Segue abaixo a notificação para análise do incidente em equipe e resposta ao NSP:\n" 
        f"• Memorando: Nº {num_memo_cru}\n" 
        f"• Notificação: Nº {num_notif}\n\n" 
        f"Atenciosamente,\n" 
        f"Ezequias S. Santos\n" 
        f"Agente Administrativo" 
    ) 
    col_nome, col_word, col_pdf, col_copiar = st.columns([1.5, 0.8, 0.8, 1.8]) 
    with col_nome: 
        st.markdown(f"**🔹 {nome_do_paciente}**") 
        if email_destino: 
            st.caption(f"📧 Destinatário: {email_destino}") 
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
        st.download_button( 
            label="📕 PDF", 
            data=word_io.getvalue(), 
            file_name=f"{nome_base_arquivo}.pdf", 
            mime="application/pdf", 
            key=f"p_{index}" 
        ) 
    with col_copiar: 
        st.code(corpo_email, language="text") 
    st.markdown("---") 

if arquivo_excel: 
    df = pd.read_excel(arquivo_excel, header=None) 
    if len(df) > 1 and ("STATUS" in str(df.iloc[0]).upper() or str(df.iloc[0]) == "1"): 
        df = df.iloc[1:] 
    elif len(df) > 2 and ("STATUS" in str(df.iloc[1]).upper() or str(df.iloc[1]) == "1"): 
        df = df.iloc[2:] 
    data_extenso_envio = obter_data_por_extenso(data_selecionada) 
    num_colunas = len(df.columns) 
    df = df.dropna(subset=[df.columns[0]]) 
    if num_colunas > 9: 
        df = df.dropna(subset=[df.columns[9]]) 
        df = df[df[df.columns[9]].astype(str).str.strip() != ""] 
    st.success(f"📋 Lista de verificação pronta! {len(df)} memorandos estruturados e validados.") 
    for index, line in df.iterrows(): 
        renderizar_linha_paciente_sob_demanda(index, line, num_colunas, data_extenso_envio) 
else: 
    st.info("💡 Por favor, suba um arquivo Excel contendo os dados para iniciar o processamento automatizado.")
