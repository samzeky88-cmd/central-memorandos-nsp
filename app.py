import os
import io
import subprocess
import tempfile
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
    titulos_negrito = (
        "MEMO:", "DE:", "PARA:", "ASSUNTO:",
        "DATA DA OCORRÊNCIA:", "DATA DA NOTIFICAÇÃO:",
        "TURNO QUE OCORREU INCIDENTE:", "ONDE OCORREU INCIDENTE:",
        "TIPO DE INCIDENTE:", "CLASSIFICAÇÃO DO INCIDENTE:",
        "DESCRIÇÃO DA NOTIFICAÇÃO:", "PACIENTE:", "LEITO:",
        "SETOR NOTIFICANTE:", "SUGESTÃO:"
    )

    def processar_paragrafo(paragrafo):
        texto_atual = paragrafo.text
        if not any(tag in texto_atual for tag in dicionario_tags):
            return

        nome_fonte = None
        tamanho_fonte = None
        for r in paragrafo.runs:
            if r.font.name:
                nome_fonte = r.font.name
            if r.font.size:
                tamanho_fonte = r.font.size
            if nome_fonte and tamanho_fonte:
                break

        for tag, valor in dicionario_tags.items():
            if tag in texto_atual:
                texto_atual = texto_atual.replace(tag, str(valor))

        for run in paragrafo.runs:
            if not run._element.xpath('.//w:drawing | .//w:pict'):
                run.text = ""

        def aplicar_estilo(run, negrito):
            run.bold = negrito
            if nome_fonte:
                run.font.name = nome_fonte
            if tamanho_fonte:
                run.font.size = tamanho_fonte

        texto_limpo = texto_atual.lstrip("• \t")
        if ":" in texto_atual and any(texto_limpo.startswith(t) for t in titulos_negrito):
            partes = texto_atual.split(":", 1)
            run_titulo = paragrafo.add_run(partes[0] + ":")
            aplicar_estilo(run_titulo, True)

            run_valor = paragrafo.add_run(partes[1])
            aplicar_estilo(run_valor, False)
        else:
            run_geral = paragrafo.add_run(texto_atual)
            aplicar_estilo(run_geral, False)

    for paragrafo in doc.paragraphs:
        processar_paragrafo(paragrafo)

    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    processar_paragrafo(paragrafo)

def converter_docx_para_pdf_bytes(docx_bytes):
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            caminho_docx = os.path.join(tmpdir, "temp.docx")
            caminho_pdf = os.path.join(tmpdir, "temp.pdf")
            with open(caminho_docx, "wb") as f:
                f.write(docx_bytes)
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf", caminho_docx, "--outdir", tmpdir],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            if os.path.exists(caminho_pdf):
                with open(caminho_pdf, "rb") as f:
                    return f.read()
    except Exception:
        return None
    return None

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
            num_memo = tratar_str_limpa(linha.iat[idx_memo])
            if num_memo and "Nº MEMO" not in num_memo.upper() and num_memo != "":
                resposta = tratar_str_limpa(linha.iat[idx_resp]) if num_colunas > idx_resp else ""
                email = tratar_str_limpa(linha.iat[idx_email]) if num_colunas > idx_email else ""
                gestor = tratar_str_limpa(linha.iat[idx_gestor]) if num_colunas > idx_gestor else ""
                setor_notif = tratar_str_limpa(linha.iat[idx_setor]) if num_colunas > idx_setor else ""
                lista_memos.append({
                    "numero_memo": num_memo,
                    "resposta": resposta,
                    "email": email,
                    "gestor": gestor,
                    "setor_notificado": setor_notif
                })
    return lista_memos
def renderizar_memorando(
    index, memo_idx, memo, linha, num_notif, data_extenso_envio, dt_ocorr, dt_notif,
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
        "{{sugestão}}": sugestao_nsp,
        "{{m}}": marca_manha,
        "{{t}}": marca_tarde,
        "{{n}}": marca_noite,
        "{{data_envio}}": data_extenso_envio
    }

    fuso_brasilia = timezone(timedelta(hours=-3))
    hora_atual = datetime.now(fuso_brasilia).hour
    if hora_atual < 12:
        saudacao = "Bom Dia Prezados"
    elif hora_atual < 18:
        saudacao = "Boa Tarde Prezados"
    else:
        saudacao = "Boa Noite Prezados"

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
        elif status_envio == "ANALISE":
            st.markdown(f"**🟡 {nome_do_paciente}** *(Em Análise)*")
        else:
            st.markdown(f"**🔹 {nome_do_paciente}**")

    word_bytes = None
    caminho_modelo = "modelo_memorando.docx"
    if os.path.exists(caminho_modelo):
        doc_word = Document(caminho_modelo)
        substituir_texto_protegendo_logos(doc_word, dados_memorando)
        word_io = io.BytesIO()
        doc_word.save(word_io)
        word_bytes = word_io.getvalue()
    else:
        st.error("Arquivo 'modelo_memorando.docx' não encontrado no servidor.")

    with col_word:
        if word_bytes:
            st.download_button(
                label="📄 WORD",
                data=word_bytes,
                file_name=f"{nome_base_arquivo}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"w_{status_envio}_{index}_{memo_idx}_{num_memo_limpo}"
            )

    with col_pdf:
        if word_bytes:
            pdf_bytes = converter_docx_para_pdf_bytes(word_bytes)
            if pdf_bytes:
                st.download_button(
                    label="📕 PDF",
                    data=pdf_bytes,
                    file_name=f"{nome_base_arquivo}.pdf",
                    mime="application/pdf",
                    key=f"p_{status_envio}_{index}_{memo_idx}_{num_memo_limpo}"
                )
            else:
                st.caption("⚠️ PDF indisponível")

    with col_copiar:
        st.code(corpo_email, language="text")

    st.markdown("---")
