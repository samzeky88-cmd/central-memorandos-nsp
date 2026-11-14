import os
import io
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta, timezone
from docx import Document

# ---------------------- CONFIGURAÇÃO DA PÁGINA ----------------------
st.set_page_config(page_title="Gerador de Memorandos", page_icon="🏥", layout="wide")

# ---------------------- CABEÇALHO COM ÍCONE ----------------------
st.markdown("""
<div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px;">
    <span style="font-size: 40px;">🏥</span>
    <div>
        <h1 style="margin: 0; font-size: 24px;">Emissor de Memorandos Individuais</h1>
        <p style="margin: 0; color: #666;">Hospital Dr. Jackson Lago — Núcleo de Segurança do Paciente</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ---------------------- FUNÇÕES AUXILIARES ----------------------
def substituir_texto_protegendo_logos(doc, dicionario_tags):
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
    try:
        if pd.isna(valor) or str(valor).strip() == "" or str(valor).strip().lower() == "nan":
            return ""
        return pd.to_datetime(valor).strftime("%d/%m/%Y")
    except Exception:
        return str(valor).strip()

def limpar_numero_float(valor):
    if pd.isna(valor) or str(valor).strip().lower() == "nan":
        return ""
    try:
        if isinstance(valor, float) and valor.is_integer():
            return str(int(valor))
        v_str = str(valor).strip()
        if v_str.endswith(".0"):
            return v_str[:-2]
        return v_str
    except Exception:
        return str(valor)

def tratar_str_limpa(valor):
    if pd.isna(valor) or str(valor).strip().lower() in ("nan", "none", ""):
        return ""
    return str(valor).strip()

def obter_data_por_extenso(dt):
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return f"{dt.day} de {meses[dt.month]} de {dt.year}"

def extrair_memorandos_da_linha(linha, num_colunas):
    lista_memos = []
    campos = [
        (15, 16, 17, 13, 14),
        (18, 19, 20, 21, 22),
        (23, 24, 25, 26, 27),
        (28, 29, 30, 31, 32),
        (33, 34, 35, 36, 37),
    ]
    for idx_memo, idx_resp, idx_email, idx_gestor, idx_setor in campos:
        if num_colunas > idx_memo:
            num_memo = tratar_str_limpa(linha.iloc[idx_memo])
            if num_memo and "Nº MEMO" not in num_memo.upper() and num_memo != "":
                resposta = tratar_str_limpa(linha.iloc[idx_resp]) if num_colunas > idx_resp else ""
                email = tratar_str_limpa(linha.iloc[idx_email]) if num_colunas > idx_email else ""
                gestor = tratar_str_limpa(linha.iloc[idx_gestor]) if num_colunas > idx_gestor else ""
                setor_notif = tratar_str_limpa(linha.iloc[idx_setor]) if num_colunas > idx_setor else ""
                lista_memos.append({
                    "numero_memo": num_memo,
                    "resposta": resposta,
                    "email": email,
                    "gestor": gestor,
                    "setor_notificado": setor_notif
                })
    return lista_memos

def renderizar_memorando(
    index, memo, linha, num_notif, data_extenso_envio, dt_ocorr, dt_notif,
    onde_ocorreu, tipo_incidente, classificacao_incidente, descricao_notificacao,
    nome_do_paciente, leito_paciente, setor_notificante_bruto, sugestao_nsp,
    turno_planilha, status_envio
):
    num_memo_cru = memo["numero_memo"]
    email_destino = memo["email"]
    gestor_destinatario = memo["gestor"] or "GESTOR DE ENFERMAGEM"
    setor_notificado_final = memo["setor_notificado"] or onde_ocorreu

    num_memo_limpo = (
        num_memo_cru.replace("Nº", "")
        .replace("No", "")
        .replace("NS", "")
        .replace("NSP", "")
        .replace("/", "-")
        .replace(" ", "")
        .strip()
    )
    nome_base_arquivo = f"MEMORANDO_N{num_memo_limpo}_NOTIF_N{num_notif}"

    marca_manha = "X" if "MANH" in turno_planilha else ""
    marca_tarde = "X" if "TARD" in turno_planilha else ""
    marca_noite = "X" if "NOIT" in turno_planilha else ""

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
        f"Segue em Anexo o Memorando {num_memo_cru} para ser analisado e respondido "
        f"(via e-mail) em até 15 dias após a data presente.\n\n"
        f"ATENÇÃO: A resposta via e-mail deve constar um arquivo em forma de word ou PDF para "
        f"arquivamento de respostas conforme rotina institucional. Não serão aceitas mensagens "
        f"via e-mail sem arquivo como resposta.\n\n"
        f"Segue abaixo a notificação para análise do incidente em equipe e resposta ao NSP:\n"
        f"• Memorando: {num_memo_cru}\n"
        f"• Notificação: N° {num_notif}\n"
        f"• Destinatário: {gestor_destinatario} — {setor_notificado_final}\n\n"
        f"Atenciosamente,\n"
        f"Ezequias S. Santos\n"
        f"Agente Administrativo - NAQH & NSP"
    )

    st.markdown(f"**📄 Memorando: {num_memo_cru}** → **{gestor_destinatario}** | *{setor_notificado_final}*")
    if email_destino:
        st.caption(f"📧 E-mail: {email_destino}")

    col_nome, col_word, col_pdf, col_copiar = st.columns([1.5, 0.8, 0.8, 1.8])

    with col_nome:
        if status_envio == "ENVIADO":
            st.markdown(f"~~{nome_do_paciente}~~ 🟢 *(Já Enviado)*")
        else:
            st.markdown(f"**🔹 {nome_do_paciente}**")

    word_io = None
    caminho_modelo = "modelo_memorando.docx"
    if os.path.exists(caminho_modelo):
        doc_word = Document(caminho_modelo)
        substituir_texto_protegendo_logos(doc_word, dados_memorando)
        word_io = io.BytesIO()
        doc_word.save(word_io)
        word_io.seek(0)

    with col_word:
        if word_io:
            st.download_button(
                label="📄 WORD",
                data=word_io.getvalue(),
                file_name=f"{nome_base_arquivo}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"w_{status_envio}_{index}_{num_memo_limpo}"
            )

    with col_pdf:
        if word_io:
            st.download_button(
                label="📕 PDF",
                data=word_io.getvalue(),
                file_name=f"{nome_base_arquivo}.pdf",
                mime="application/pdf",
                key=f"p_{status_envio}_{index}_{num_memo_limpo}"
            )

    with col_copiar:
        st.code(corpo_email, language="text")

    st.markdown("---")


def processar_linha(index, linha, num_colunas, data_extenso_envio):
    num_notif = limpar_numero_float(linha.iloc[0]) if num_colunas > 0 else "S-N"
    if num_notif.upper() == "STATUS" or "NOTIF" in num_notif.upper() or num_notif == "" or num_notif == "1":
        return None, None, 0

    status_envio = str(linha.iloc[1]).strip().upper() if num_colunas > 1 else "PENDENTE"
    dt_ocorr = formatar_data_br(linha.iloc[2]) if num_colunas > 2 else ""
    dt_notif = formatar_data_br(linha.iloc[3]) if num_colunas > 3 else ""
    turno_planilha = str(linha.iloc[4]).strip().upper() if num_colunas > 4 else ""
    onde_ocorreu = tratar_str_limpa(linha.iloc[5]) if num_colunas > 5 else "Não informado"
    tipo_incidente = tratar_str_limpa(linha.iloc[6]) if num_colunas > 6 else ""
    classificacao_incidente = tratar_str_limpa(linha.iloc[7]) if num_colunas > 7 else ""
    descricao_notificacao = tratar_str_limpa(linha.iloc[8]) if num_colunas > 8 else ""
    nome_do_paciente = tratar_str_limpa(linha.iloc[9]) if num_colunas > 9 else "Paciente Não Identificado"

    if nome_do_paciente.upper() == "PACIENTE":
        return None, None, 0

    leito_paciente = limpar_numero_float(linha.iloc[10]) if num_colunas > 10 else ""
    setor_notificante_bruto = tratar_str_limpa(linha.iloc[11]) if num_colunas > 11 else ""
    sugestao_nsp = tratar_str_limpa(linha.iloc[12]) if num_colunas > 12 else ""

    lista_memos = extrair_memorandos_da_linha(linha, num_colunas)
    if not lista_memos:
        return None, None, 0

    dados_linha = {
        "index": index,
        "linha": linha,
        "num_notif": num_notif,
        "data_extenso_envio": data_extenso_envio,
        "dt_ocorr": dt_ocorr,
        "dt_notif": dt_notif,
        "onde_ocorreu": onde_ocorreu,
        "tipo_incidente": tipo_incidente,
        "classificacao_incidente": classificacao_incidente,
        "descricao_notificacao": descricao_notificacao,
        "nome_do_paciente": nome_do_paciente,
        "leito_paciente": leito_paciente,
        "setor_notificante_bruto": setor_notificante_bruto,
        "sugestao_nsp": sugestao_nsp,
        "turno_planilha": turno_planilha,
        "status_envio": status_envio,
        "lista_memos": lista_memos
    }

    return status_envio, dados_linha, len(lista_memos)


# ---------------------- INTERFACE PRINCIPAL ----------------------
arquivo_planilha = st.file_uploader("📂 Carregue a planilha de notificações (.xlsx ou .csv)", type=["xlsx", "csv"])

if arquivo_planilha is not None:
    if arquivo_planilha.name.endswith(".csv"):
        df = pd.read_csv(arquivo_planilha)
    else:
        df = pd.read_excel(arquivo_planilha)

    fuso_brasilia = timezone(timedelta(hours=-3))
    data_extenso_envio = obter_data_por_extenso(datetime.now(fuso_brasilia))
    num_colunas = len(df.columns)

    for index, linha in df.iterrows():
        status, dados, qtd = processar_linha(index, linha, num_colunas, data_extenso_envio)
        if dados:
            for memo in dados["lista_memos"]:
                renderizar_memorando(
                    index=dados["index"],
                    memo=memo,
                    linha=dados["linha"],
                    num_notif=dados["num_notif"],
                    data_extenso_envio=dados["data_extenso_envio"],
                    dt_ocorr=dados["dt_ocorr"],
                    dt_notif=dados["dt_notif"],
                    onde_ocorreu=dados["onde_ocorreu"],
                    tipo_incidente=dados["tipo_incidente"],
                    classificacao_incidente=dados["classificacao_incidente"],
                    descricao_notificacao=dados["descricao_notificacao"],
                    nome_do_paciente=dados["nome_do_paciente"],
                    leito_paciente=dados["leito_paciente"],
                    setor_notificante_bruto=dados["setor_notificante_bruto"],
                    sugestao_nsp=dados["sugestao_nsp"],
                    turno_planilha=dados["turno_planilha"],
                    status_envio=dados["status_envio"]
                )
