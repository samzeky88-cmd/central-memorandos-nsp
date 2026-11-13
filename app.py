# =================================================================
# 📄 GERADOR DE MEMORANDOS — HOSPITAL DR. JACKSON LAGO
# Versão: Leitura de 4 Blocos por Linha + E-mails + Texto Completo
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
    ("ALA A", "alaaclinicamedica@gmail.com"),
    ("ALA B", "alab.hcid@gmail.com"),
    ("ALA C", "cirurgica.hcid@gmail.com"),
    ("ALA D", "ortopediaalad.hcid@gmail.com"),
    ("ALA F (UTI B)", "utihcid@gmail.com"),
    ("ALA G/H", "aladgh.hcid@gmail.com"),
    ("ALA I", "ortopediaalai.hcid@gmail.com"),
    ("ALA J", "ortopediaalaj.hcid@gmail.com"),
    ("ALA L", "alalortopedia.hcid@gmail.com"),
    ("CENTRO CIRÚRGICO", "centrocirurgicosoc2@gmail.com"),
    ("CCIH", "ccih.hcid@gmail.com"),
    ("DIREÇÃO TÉCNICA", "direcaotecnica.hcid@gmail.com"),
    ("DIREÇÃO", "direcao.hospitaldacidade@gmail.com"),
    ("ECP", "carlosemilioecp@gmail.com"),
    ("ECP - EQUIPE DE CUIDADOS COM A PELE", "fisioreabsoc2@gmail.com"),
    ("ENDOSCOPIA", "endoscopiasoc2@gmail.com"),
    ("FISIOTERAPIA", "fisioterapiasocorrao@gmail.com"),
    ("FARMACIA", "farmacia.hcid@gmail.com"),
    ("FARMACOVIGILÂNCIA", "farmacovigilanciahcid@gmail.com"),
    ("GESTOR DE ENFERMAGEM", "gerenciadeenf.hcid@gmail.com"),
    ("HOTELARIA", "hotelaria.hcid@gmail.com"),
    ("NIR", "coordenanirs2@gmail.com"),
    ("NSP", "nspsoc2@gmail.com"),
    ("NÚCLEO INTERNO DE REGULAÇÃO DE LEITOS", "coordenanirs2@gmail.com"),
    ("SADT", "sadthosp@gmail.com"),
    ("SALA VERMELHA", "salavermelhasdm@gmail.com"),
    ("SERVIÇO SOCIAL", "servicosocialhcid@gmail.com"),
    ("SAET", "coodsaet2026@gmail.com"),
    ("SETOR DE IMAGEM", "sadt.hosp@gmail.com"),
    ("TERAPIA OCUPACIONAL", "terapiaocupacional.hcid@gmail.com"),
    ("EMTN - EQUIPE", "emtnhospitaldacidade@gmail.com"),
    ("FISIOTERAPIA", "fisioterapiasocorrao@gmail.com"),
    ("SOM - EQUIPE DE CUIDADOS MÉDICA", "sadm.hcid@gmail.com"),
    ("SADT - SETOR DE IMAGEM", "sadt.hosp@gmail.com"),
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
        texto_completo = "".join(run.text for run in p.runs)
        alterado = False
        for chave, valor in dados.items():
            if chave in texto_completo:
                texto_completo = texto_completo.replace(chave, str(valor))
                alterado = True
        if alterado:
            p.clear()
            p.add_run(texto_completo)
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

def tratar_texto(valor):
    if pd.isna(valor) or str(valor).strip().lower() in ["nan", "none", ""]:
        return ""
    return str(valor).strip()

def data_por_extenso(dt):
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    return f"{dt.day} de {meses[dt.month-1]} de {dt.year}"

def buscar_email_setor(nome_setor):
    if not nome_setor or str(nome_setor).strip() == "":
        return ""
    nome_setor = str(nome_setor).strip().upper()
    for _, linha in df_setores.iterrows():
        setor_lista = linha["SETOR"].upper()
        if nome_setor in setor_lista or setor_lista in nome_setor:
            return linha["E-MAIL"]
    return ""

def montar_corpo_email(num_memo, num_notif, gestor, setor, nome_paciente, data_envio):
    assunto = f"NOTIFICAÇÃO DE INCIDENTE — MEMO Nº {num_memo} / NOTIF Nº {num_notif}"
    
    corpo = f"""
Bom Dia / Boa Tarde, Prezado(a) {gestor}!

Segue anexo o Memorando Nº {num_memo} referente à Notificação de Incidente Nº {num_notif}, referente ao setor {setor}.

📋 DADOS:
• Paciente: {nome_paciente}
• Setor Destinatário: {setor}
• Memorando Nº: {num_memo}
• Notificação Nº: {num_notif}

Sugerimos analisar o incidente juntamente com a equipe assistencial e discutir propostas de cuidados e prevenção conforme Protocolo Institucional.

Conforme rotina institucional, o gestor tem o prazo de 15 dias para realizar comunicação do incidente com sua equipe e discutir barreiras para evitar a ocorrência de novos eventos.

Atenciosamente,

Fabrícia Rocha
Coordenadora do Núcleo de Segurança do Paciente
Hospital da Cidade Dr. Jackson Lago
E-mail: nspsoc2@gmail.com
"""
    return assunto, corpo.strip()

# =================================================================
# 📄 FUNÇÃO DE PROCESSAR 1 BLOCO DE MEMORANDO
# =================================================================
def processar_bloco(linha, bloco_nome, idx, nome_paciente, data_extenso, 
                    col_gestor, col_setor, col_memo, 
                    dt_ocorr="", dt_notif="", local="", tipo="", 
                    classificacao="", desc="", leito="", setor_notificante="", 
                    sugestao="", turno_bruto=""):
    
    gestor = tratar_texto(linha[col_gestor]) if col_gestor < len(linha) else ""
    setor_destino = tratar_texto(linha[col_setor]) if col_setor < len(linha) else ""
    num_memo = limpar_numero(linha[col_memo]) if col_memo < len(linha) else ""
    
    # Pular se não houver número de memorando
    if not num_memo or num_memo == "" or num_memo.upper() == "Nº MEMO 01":
        return None
    
    num_notif = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
    email_destino = buscar_email_setor(setor_destino)
    
    # Processar turno
    if pd.isna(turno_bruto):
        m, t, n = "", "", ""
    else:
        tx = str(turno_bruto).strip().upper()
        m = "X" if "MANH" in tx else ""
        t = "X" if "TARD" in tx else ""
        n = "X" if "NOIT" in tx else ""
    
    dados_memorando = {
        "{{numero_memorando}}": num_memo,
        "{{gestor}}": gestor,
        "{{setor}}": setor_destino,
        "{{notificacao_n}}": num_notif,
        "{{data_notificacao}}": dt_notif,
        "{{data_ocorrencia}}": dt_ocorr,
        "{{localizacao}}": local,
        "{{tipo_incidente}}": tipo,
        "{{classificacao_incidente}}": classificacao,
        "{{descricao_notificacao}}": desc,
        "{{nome_paciente}}": nome_paciente,
        "{{leito}}": leito,
        "{{setor_notificante}}": setor_notificante,
        "{{sugestao}}": sugestao,
        "{{data_envio}}": data_extenso,
        "{{m}}": m,
        "{{t}}": t,
        "{{n}}": n,
    }
    
    assunto_email, corpo_email = montar_corpo_email(num_memo, num_notif, gestor, setor_destino, nome_paciente, data_extenso)
    
    return {
        "bloco": bloco_nome,
        "num_memo": num_memo,
        "num_notif": num_notif,
        "gestor": gestor,
        "setor": setor_destino,
        "email": email_destino,
        "dados_docx": dados_memorando,
        "assunto_email": assunto_email,
        "corpo_email": corpo_email,
    }

def gerar_docx(dados, nome_arquivo_base):
    try:
        with open(caminho_modelo, "rb") as f:
            doc = Document(f)
        substituir_texto_no_docx(doc, dados)
        saida = io.BytesIO()
        doc.save(saida)
        saida.seek(0)
        return saida
    except FileNotFoundError:
        st.error("❌ Arquivo modelo_memorando.docx não encontrado!")
        return None
    except Exception as e:
        st.error(f"❌ Erro: {str(e)}")
        return None

# =================================================================
# 🚀 EXECUÇÃO PRINCIPAL
# =================================================================
if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel, engine="openpyxl")
        st.success(f"✅ Planilha carregada — {len(df)} registro(s) encontrado(s)")
        data_extenso = data_por_extenso(data_selecionada)
        
        # Mostrar mapeamento das colunas para conferir
        with st.expander("🔍 Ver colunas da planilha"):
            for i, nome in enumerate(df.columns):
                st.write(f"Coluna {i} → {nome}")
        
        for idx, linha in df.iterrows():
            # Dados comuns a TODOS os blocos da mesma linha
            num_notif = limpar_numero(linha.iloc[0]) if len(linha) > 0 else ""
            if not num_notif or num_notif.upper() == "STATUS" or "NOTIF" in str(num_notif).upper():
                continue
            
            nome_paciente = tratar_texto(linha.iloc[9]) if len(linha) > 9 else "Paciente"
            if nome_paciente.upper() == "PACIENTE":
                continue
            
            # Dados que se repetem em todos os memorandos da mesma linha
            dt_ocorr = formatar_data_br(linha.iloc[2]) if len(linha) > 2 else ""
            dt_notif = formatar_data_br(linha.iloc[3]) if len(linha) > 3 else ""
            local = tratar_texto(linha.iloc[5]) if len(linha) > 5 else ""
            tipo = tratar_texto(linha.iloc[6]) if len(linha) > 6 else ""
            classificacao = tratar_texto(linha.iloc[7]) if len(linha) > 7 else ""
            desc = tratar_texto(linha.iloc[8]) if len(linha) > 8 else ""
            leito = limpar_numero(linha.iloc[10]) if len(linha) > 10 else ""
            setor_notificante = tratar_texto(linha.iloc[11]) if len(linha) > 11 else ""
            sugestao = tratar_texto(linha.iloc[12]) if len(linha) > 12 else ""
            turno_bruto = tratar_texto(linha.iloc[4]) if len(linha) > 4 else ""
            
            # 🧩 PROCESSAR OS 4 BLOCOS — cada um com suas colunas
            # Bloco 1: Colunas N(13)=Gestor01, O(14)=Setor01, P(15)=Memo01
            bloco1 = processar_bloco(linha, "Memorando 01", idx, nome_paciente, data_extenso,
                                     col_gestor=13, col_setor=14, col_memo=15,
                                     dt_ocorr, dt_notif, local, tipo, classificacao,
                                     desc, leito, setor_notificante, sugestao, turno_bruto)
            
            # Bloco 2: Colunas R(17)=Gestor02, S(18)=Setor02, T(19)=Memo02
            bloco2 = processar_bloco(linha, "Memorando 02", idx, nome_paciente, data_extenso,
                                     col_gestor=17, col_setor=18, col_memo=19,
                                     dt_ocorr, dt_notif, local, tipo, classificacao,
                                     desc, leito, setor_notificante, sugestao, turno_bruto)
            
            # Bloco 3: Colunas W(22)=Gestor03, X(23)=Setor03, Y(24)=Memo03
            bloco3 = processar_bloco(linha, "Memorando 03", idx, nome_paciente, data_extenso,
                                     col_gestor=22, col_setor=23, col_memo=24,
                                     dt_ocorr, dt_notif, local, tipo, classificacao,
                                     desc, leito, setor_notificante, sugestao, turno_bruto)
            
            # Bloco 4: Colunas AA(26)=Gestor04, AB(27)=Setor04, AC(28)=Memo04
            bloco4 = processar_bloco(linha, "Memorando 04", idx, nome_paciente, data_extenso,
                                     col_gestor=26, col_setor=27, col_memo=28,
                                     dt_ocorr, dt_notif, local, tipo, classificacao,
                                     desc, leito, setor_notificante, sugestao, turno_bruto)
            
            # Coletar apenas os blocos que tem número de memorando
            blocos_ativos = [b for b in [bloco1, bloco2, bloco3, bloco4] if b is not None]
            
            with st.expander(f"📋 Notificação {num_notif} — {nome_paciente} — {len(blocos_ativos)} memorando(s)"):
                st.write(f"**📅 Data:** {data_extenso} | **📋 Total de Memorandos:** {len(blocos_ativos)}")
                
                for b in blocos_ativos:
                    st.markdown(f"### 📄 {b['bloco']} — Nº {b['num_memo']}")
                    st.write(f"**Setor:** {b['setor']} | **Gestor:** {b['gestor']}")
                    
                    if b['email']:
                        st.info(f"📧 **E-mail:** {b['email']}")
                    else:
                        st.warning("⚠️ E-mail não encontrado")
                    
                    # Botão de download
                    docx_bytes = gerar_docx(b['dados_docx'], f"MEMO_{b['num_memo']}_NOTIF_{num_notif}")
                    if docx_bytes:
                        st.download_button(
                            label=f"📥 Baixar {b['bloco']} (DOCX)",
                            data=docx_bytes,
                            file_name=f"MEMORANDO_{b['num_memo']}_NOTIF_{num_notif}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"doc_{idx}_{b['bloco']}"
                        )
                    
                    # Assunto e corpo do e-mail
                    st.write("**✉️ Assunto do E-mail:**")
                    st.code(b['assunto_email'], language=None
