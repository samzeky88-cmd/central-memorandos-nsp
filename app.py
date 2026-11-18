import os
import io
import subprocess
import tempfile
import pandas as pd
import streamlit as st
from datetime import datetime, timedelta, timezone
from docx import Document

st.set_page_config(page_title="Gerador de Memorandos", page_icon="🏥", layout="wide")

st.markdown("""
<div style="display: flex; align-items: center; gap: 15px; margin-bottom: 20px;">
    <span style="font-size: 40px;">🏥</span>
    <div>
        <h1 style="margin: 0; font-size: 24px;">Emissor de Memorandos Individuais</h1>
        <p style="margin: 0; color: #666;">Hospital Dr. Jackson Lago — Núcleo de Segurança do Paciente</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ============================================================
# FUNÇÃO INTELIGENTE DE MONTAGEM DO DESTINATÁRIO
# ============================================================
def montar_destinatario(gestor_col, setor_col):
    def tratar_texto(t):
        if not t or str(t).strip() == "":
            return ""
        return str(t).strip().upper()
    
    g = tratar_texto(gestor_col)
    s = tratar_texto(setor_col)
    
    if not g:
        return s if s else "GESTOR RESPONSÁVEL"
    if not s:
        return g
    
    if g == s or s in g:
        return g
    
    abreviaturas = {"SDM": "DA", "NIR": "DO", "SADT": "DO", "SAET": "DO"}
    if s in abreviaturas:
        preposicao = abreviaturas[s]
        if f" {preposicao} " not in g and not g.endswith(f" {preposicao[:-1]}"):
            return f"{g} {preposicao} {s}"
        return f"{g} {s}"
    
    palavras_g = set(g.replace(" DE ", " ").replace(" DA ", " ").replace(" DO ", " ").split())
    palavras_s = s.split()
    if palavras_s[0] in palavras_g:
        s_limpo = " ".join(palavras_s[1:]).strip()
        if s_limpo:
            return f"{g} - {s_limpo}"
        return g
    
    return f"{g} - {s}"

def limpar_valor(valor):
    if pd.isna(valor):
        return ""
    texto = str(valor).strip()
    if texto.lower() == "n/a":
        return "N/A"
    if texto == "":
        return ""
    return texto

def substituir_texto_protegido(doc, substituicoes):
    for paragrafo in doc.paragraphs:
        for chave, valor in substituicoes.items():
            if chave in paragrafo.text:
                paragrafo.text = paragrafo.text.replace(chave, valor)
    for tabela in doc.tables:
        for linha in tabela.rows:
            for celula in linha.cells:
                for paragrafo in celula.paragraphs:
                    for chave, valor in substituicoes.items():
                        if chave in paragrafo.text:
                            paragrafo.text = paragrafo.text.replace(chave, valor)

def converter_docx_para_pdf(docx_bytes):
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            caminho_docx = os.path.join(tmpdir, "temp.docx")
            caminho_pdf = os.path.join(tmpdir, "temp.pdf")
            with open(caminho_docx, "wb") as f:
                f.write(docx_bytes)
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", tmpdir, caminho_docx],
                check=True, capture_output=True
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
        return limpar_valor(valor)

def obter_data_por_extenso(dt):
    meses = {
        1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
        5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
        9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
    }
    return f"{dt.day} de {meses[dt.month]} de {dt.year}"

def extrair_memorandos_da_linha(linha, num_colunas):
    lista = []
    campos = [
        (15, 16, 17, 13, 14),
        (20, 21, 22, 18, 19),
        (25, 26, 27, 23, 24),
        (30, 31, 32, 28, 29),
        (35, 36, 37, 33, 34),
    ]
    for col_memo, col_email, col_resposta, col_gestor, col_setor in campos:
        if num_colunas > col_memo:
            numero_memo = limpar_valor(linha.iloc[col_memo])
            if numero_memo and numero_memo not in ("", "0", "-"):
                lista.append({
                    "numero_memo": numero_memo,
                    "email": limpar_valor(linha.iloc[col_email]),
                    "resposta": limpar_valor(linha.iloc[col_resposta]),
                    "gestor": limpar_valor(linha.iloc[col_gestor]),
                    "setor_notificado": limpar_valor(linha.iloc[col_setor])
                })
    return lista

def chave_numerica(valor):
    digitos = "".join(c for c in str(valor) if c.isdigit())
    return int(digitos) if digitos else 0

def processar_linha(index, linha, num_colunas, data_atual):
    def col(i):
        return linha.iloc[i] if num_colunas > i else ""
    
    num_notif = limpar_valor(col(0))
    if not num_notif or num_notif.upper() in ("Nº", "NUMERO", "NÚMERO", ""):
        return None, None, 0
    
    nome_paciente = limpar_valor(col(9))
    if nome_paciente.upper() in ("PACIENTE", "NOME", ""):
        return None, None, 0
    
    status_celula = limpar_valor(col(1)).upper()
    if "ENVIAD" in status_celula and "NÃO" not in status_celula and "NAO" not in status_celula:
        status = "ENVIADO"
    elif "TRATATIVA" in status_celula:
        status = "TRATATIVA"
    elif "RETIFICAR" in status_celula:
        status = "RETIFICAR"
    elif "NÃO ENVIAR" in status_celula or "NAO ENVIAR" in status_celula:
        status = "NAO_ENVIAR"
    elif "ANALIS" in status_celula or "ANÁLIS" in status_celula:
        status = "ANALISE"
    else:
        status = "PENDENTE"
    
    dados = {
        "index": index,
        "num_notif": num_notif,
        "status": status,
        "data_ocorrencia": formatar_data_br(col(2)),
        "data_notificacao": formatar_data_br(col(3)),
        "turno": limpar_valor(col(4)),
        "local": limpar_valor(col(5)),
        "tipo_incidente": limpar_valor(col(6)),
        "classificacao": limpar_valor(col(7)),
        "descricao": limpar_valor(col(8)),
        "nome_paciente": nome_paciente,
        "leito": limpar_valor(col(10)),
        "setor_origem": limpar_valor(col(11)),
        "sugestao": limpar_valor(col(12)),
        "memorandos": extrair_memorandos_da_linha(linha, num_colunas)
    }
    return status, dados, len(dados["memorandos"])

def gerar_memorando(dados_principais, dados_memo, data_emissao):
    modelo_path = "modelo_memorando.docx"
    if not os.path.exists(modelo_path):
        modelo_path = "modelo.docx"
    
    if not os.path.exists(modelo_path):
        st.error("❌ ARQUIVO MODELO NÃO ENCONTRADO!")
        st.warning("Por favor, verifique se o arquivo 'modelo_memorando.docx' está na pasta do projeto.")
        return None, "Modelo não encontrado"
    
    doc = Document(modelo_path)
    
    destinatario = montar_destinatario(
        dados_memo["gestor"],
        dados_memo["setor_notificado"]
    )
    
    sugestao_final = dados_principais["sugestao"]
    if not sugestao_final:
        sugestao_final = "N/A"
    
    substituicoes = {
        "{{numero_memorando}}": dados_memo["numero_memo"],
        "{{ano}}": str(data_emissao.year),
        "{{setor}}": destinatario,
        "{{notificacao_n}}": dados_principais["num_notif"],
        "{{data_envio}}": obter_data_por_extenso(data_emissao),
        "{{data_ocorrencia}}": dados_principais["data_ocorrencia"] or "N/A",
        "{{data_notificacao}}": dados_principais["data_notificacao"] or "N/A",
        "{{m}}": "( X )" if "MANHÃ" in dados_principais["turno"].upper() else "(   )",
        "{{t}}": "( X )" if "TARDE" in dados_principais["turno"].upper() else "(   )",
        "{{n}}": "( X )" if "NOITE" in dados_principais["turno"].upper() else "(   )",
        "{{localizacao}}": dados_principais["local"] or "N/A",
        "{{tipo_incidente}}": dados_principais["tipo_incidente"] or "N/A",
        "{{classificacao_incidente}}": dados_principais["classificacao"] or "N/A",
        "{{descricao_notificacao}}": dados_principais["descricao"] or "N/A",
        "{{nome_paciente}}": dados_principais["nome_paciente"] or "N/A",
        "{{leito}}": dados_principais["leito"] or "N/A",
        "{{setor_notificante}}": dados_principais["setor_origem"] or "N/A",
        "{{sugestao}}": sugestao_final
    }
    
    substituir_texto_protegido(doc, substituicoes)
    
    nome_arquivo = f"NOTIF_N{dados_principais['num_notif']}_MEMO_N{dados_memo['numero_memo']}-{data_emissao.year}.docx"
    
    doc_io = io.BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    conteudo = doc_io.getvalue()
    
    return conteudo, nome_arquivo

def exibir_memorando(dados_principais, dados_memo, data_emissao):
    destinatario = montar_destinatario(
        dados_memo["gestor"],
        dados_memo["setor_notificado"]
    )
    
    st.markdown(f"---")
    st.subheader(f"📝 Memorando Nº {dados_memo['numero_memo']} / {data_emissao.year}")
    st.markdown(f"""
    - **Para:** {destinatario}
    - **Notificação:** {dados_principais['num_notif']}
    - **Paciente:** {dados_principais['nome_paciente']}
    - **Sugestão:** {dados_principais['sugestao'] or 'N/A'}
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        docx_bytes, nome_docx = gerar_memorando(dados_principais, dados_memo, data_emissao)
        if docx_bytes:
            st.download_button(
                label="📄 Baixar DOCX",
                data=docx_bytes,
                file_name=nome_docx,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"docx_{dados_principais['index']}_{dados_memo['numero_memo']}"
            )
    with col2:
        pdf_bytes = converter_docx_para_pdf(docx_bytes) if docx_bytes else None
        if pdf_bytes:
            nome_pdf = nome_docx.replace(".docx", ".pdf")
            st.download_button(
                label="📕 Baixar PDF",
                data=pdf_bytes,
                file_name=nome_pdf,
                mime="application/pdf",
                key=f"pdf_{dados_principais['index']}_{dados_memo['numero_memo']}"
            )

@st.cache_data
def carregar_planilha(arquivo):
    if arquivo.name.endswith(".csv"):
        return pd.read_csv(arquivo)
    return pd.read_excel(arquivo)

# ============================================================
# 📂 UPLOAD E PROCESSAMENTO PRINCIPAL
# ============================================================
arquivo_planilha = st.file_uploader("📂 Carregue a planilha de notificações (.xlsx ou .csv)", type=["xlsx", "csv"])

if arquivo_planilha is not None:
    df = carregar_planilha(arquivo_planilha)
    
    fuso_brasilia = timezone(timedelta(hours=-3))
    data_emissao_atual = datetime.now(fuso_brasilia)
    num_colunas = len(df.columns)
    
    registros = []
    contadores = {
        "PENDENTE": 0, "ANALISE": 0, "ENVIADO": 0,
        "TRATATIVA": 0, "RETIFICAR": 0, "NAO_ENVIAR": 0
    }
    
    for indice, linha in df.iterrows():
        status, dados, qtd_memos = processar_linha(indice, linha, num_colunas, data_emissao_atual)
        if dados:
            registros.append(dados)
            contadores[status] += qtd_memos
    
    st.info(
        f"📊 **Resumo da Planilha:** "
        f"🔵 **{contadores['PENDENTE']}** Pendentes | "
        f"🟡 **{contadores['ANALISE']}** Em Análise | "
        f"🟢 **{contadores['ENVIADO']}** Já Enviados | "
        f"🟠 **{contadores['TRATATIVA']}** Em Tratativa | "
        f"🔴 **{contadores['RETIFICAR']}** Retificar | "
        f"⚫ **{contadores['NAO_ENVIAR']}** Não Enviar"
    )
    
    col_pesq, col_status, col_ordem = st.columns([2, 1.5, 1.2])
    with col_pesq:
        busca = st.text_input(
            "🔍 Pesquisar por Nº da Notificação ou Paciente:",
            placeholder="Ex: 886 ou Antônia Maria"
        ).strip().upper()
    with col_status:
        filtro = st.selectbox(
            "📌 Filtrar por Status:",
            ["Todos", "Pendentes", "Em Análise", "Em Tratativa", "Retificar", "Não Enviar", "Já Enviados"]
        )
    with col_ordem:
        ordem = st.selectbox(
            "🔢 Ordem:",
            ["Menor para o Maior", "Maior para o Menor"]
        )
    
    mapa_filtro = {
        "Pendentes": "PENDENTE",
        "Em Análise": "ANALISE",
        "Em Tratativa": "TRATATIVA",
        "Retificar": "RETIFICAR",
        "Não Enviar": "NAO_ENVIAR",
        "Já Enviados": "ENVIADO"
    }
    if filtro != "Todos":
        registros = [r for r in registros if r["status"] == mapa_filtro[filtro]]
    
    if busca:
        registros = [r for r in registros 
                     if busca in r["num_notif"] or busca in r["nome_paciente"].upper()]
    
    registros.sort(key=lambda r: chave_numerica(r["num_notif"]), reverse=(ordem == "Maior para o Menor"))
    
    # ============================================================
    # 📊 ABA: ESTATÍSTICAS E RANKING DE SETORES
    # ============================================================
    st.markdown("---")
    st.header("📊 Estatísticas dos Setores Mais Notificados")
    
    contagem_setores = {}
    for registro in registros:
        setor = registro["local"] or "Não informado"
        if setor in contagem_setores:
            contagem_setores[setor] += 1
        else:
            contagem_setores[setor] = 1
    
    ranking = sorted(contagem_setores.items(), key=lambda x: x[1], reverse=True)
    
    aba_ranking, aba_grafico = st.tabs(["🏆 Ranking", "📈 Gráfico"])
    
    with aba_ranking:
        st.subheader("🏆 Ranking dos Setores com Mais Notificações")
        st.markdown("Ordenado por quantidade de incidentes registrados:")
        
        tabela_ranking = []
        for posicao, (setor, quantidade) in enumerate(ranking, 1):
            medalha = "🥇" if posicao == 1 else "🥈" if posicao == 2 else "🥉" if posicao == 3 else f"{posicao}º"
            porcentagem = (quantidade / len(registros) * 100) if registros else 0
            tabela_ranking.append({
                "Posição": medalha,
                "Setor": setor,
                "Notificações": quantidade,
                "% do Total": f"{porcentagem:.1f}%"
            })
        
        df_ranking = pd.DataFrame(tabela_ranking)
        st.dataframe(df_ranking, use_container_width=True, hide_index=True)
        
        st.info(f"📋 Total de notificações analisadas: **{len(registros)}** em **{len(ranking)}** setores distintos")
    
    with aba_grafico:
        st.subheader("📈 Gráfico — Top 15 Setores Mais Notificados")
        
        top_setores = ranking[:15]
        nomes = [item[0] for item in top_setores]
        quantidades
