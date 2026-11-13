# =================================================================
# 📄 GERADOR DE MEMORANDOS + E-MAILS — HOSPITAL DR. JACKSON LAGO
# Versão Completa: Tags + Turno + Classificação + E-mails + Texto
# =================================================================

import os
import pandas as pd
import streamlit as st
import io
from datetime import datetime
from docx import Document

# =================================================================
# 📚 LISTA DE E-MAILS DOS SETORES — COMPLETA
# =================================================================
LISTA_SETORES = [
    ("DIREÇÃO", "direcao.hospitaldacidade@gmail.com"),
    ("GABINETE DA DIREÇÃO", "gabinete.hospitaldacidade@gmail.com"),
    ("AGÊNCIA TRANSFUSIONAL", "agenciatransfusional.hcid@gmail.com"),
    ("ALA A - ENFERMAGEM", "alaaclinicamedica@gmail.com"),
    ("ALA B - ENFERMAGEM", "alab.hcid@gmail.com"),
    ("ALA C - ENFERMAGEM", "cirurgica.hcid@gmail.com"),
    ("ALA D - ENFERMAGEM", "ortopediaalad.hcid@gmail.com"),
    ("ALA G/H", "aladgh.hcid@gmail.com"),
    ("ALA F - ENFERMAGEM", "utihcid@gmail.com"),
    ("ALA I - ENFERMAGEM", "ortopediaalai.hcid@gmail.com"),
    ("ALA J - ENFERMAGEM", "ortopediaalaj.hcid@gmail.com"),
    ("ALA L - ENFERMAGEM", "alalortopedia.hcid@gmail.com"),
    ("ALMOXARIFADO", "almoxarifadosocorrao2@gmail.com"),
    ("AMBULATÓRIO", "ambulatorio.hcid@gmail.com"),
    ("ASSESSORIA JURÍDICA", "fernandoacmoraes.advogado@gmail.com"),
    ("CAC", "cac.socorrao2@gmail.com"),
    ("CENTRAL DE CÓPIAS", "centraldecopias.hcid@gmail.com"),
    ("CCIH", "ccih.hcid@gmail.com"),
    ("CENTRO CIRÚRGICO - ENFERMAGEM", "centrocirurgicosoc2@gmail.com"),
    ("COMISSÃO DE ÓBITOS", "comissaodeobito.hdacidade@gmail.com"),
    ("COMISSÃO DE ÉTICA EM ENFERMAGEM", "comissaodeetica.soc2@gmail.com"),
    ("CME", "cmeesterilizacao@gmail.com"),
    ("DOCUMENTOS INSTITUCIONAIS (NAQH)", "documentos.soc2@gmail.com"),
    ("ECP", "carlosemilioecp@gmail.com"),
    ("ENDOSCOPIA", "endoscopiasoc2@gmail.com"),
    ("EDOT", "edot.hcid@gmail.com"),
    ("EMTN", "emtnhospitaldacidade@gmail.com"),
    ("ESCALAS", "escalahcid25@gmail.com"),
    ("FARMÁCIA", "farmacia.hcid@gmail.com"),
    ("FARMACOVIGILÂNCIA", "farmacovigilanciahcid@gmail.com"),
    ("FATURAMENTO", "contasmedicassocorrao2@gmail.com"),
    ("FISIOTERAPIA /PS - ENFERMARIAS", "fisioreabsoc2@gmail.com"),
    ("FISIOTERAPIA - UTI", "fisioterapiasocorrao@gmail.com"),
    ("FONOAUDIOLOGIA", "fonosocorrao2@yahoo.com"),
    ("GERÊNCIA DE ENFERMAGEM", "gerenciadeenf.hcid@gmail.com"),
    ("GERÊNCIA DA REABILITAÇÃO", "gerenciareab.hcid@gmail.com"),
    ("GERENCIAMENTO DE RESÍDUOS", "residuos.hcid@gmail.com"),
    ("HOTELARIA", "hotelaria.hcid@gmail.com"),
    ("LABORATÓRIO", "hcid.laboratorio@gmail.com"),
    ("MANUTENÇÃO", "manutencao.hcid@gmail.com"),
    ("MAQUEIROS", "centraltransporte.hcid@gmail.com"),
    ("NAQH", "naqhsoc2@gmail.com"),
    ("NCE", "ce.socorrao2@gmail.com"),
    ("NEP", "nep.hcid@gmail.com"),
    ("NEPEX", "nepexhcid@gmail.com"),
    ("NSP", "nspsoc2@gmail.com"),
    ("NIR", "coordenanirs2@gmail.com"),
    ("NUTRIÇÃO", "nutricaohospitaldacidade@gmail.com"),
    ("NVE", "nhe.hcid@gmail.com"),
    ("ODONTOLOGIA", "robertapalacioferes@gmail.com"),
    ("OUVIDORIA", "ouvidoriahcid@gmail.com"),
    ("PARECER DE EXAMES", "coodsaet2026@gmail.com"),
    ("PATRIMÔNIO", "patrimonio.hcid@gmail.com"),
    ("PORTARIA", "portaria.hcid@gmail.com"),
    ("PSICOLOGIA", "psicologia.hcid@gmail.com"),
    ("RECURSOS HUMANOS", "setorpessoal.hcid@gmail.com"),
    ("ROUPARIA", "rouparia.hcid@gmail.com"),
    ("SAEEM", "saeem.hcid@gmail.com"),
    ("SAME", "same25.hcid@gmail.com"),
    ("SAET", "coodsaet2026@gmail.com"),
    ("SADT", "sadthosp@gmail.com"),
    ("SDM/VERMELHA - ENFERMAGEM", "salavermelhasdm@gmail.com"),
    ("SERVIÇO SOCIAL", "servicosocialhcid@gmail.com"),
    ("SETOR DE IMAGEM", "sadt.hosp@gmail.com"),
    ("TECNOLOGIA DA INFORMAÇÃO", "suporteti.hcid@gmail.com"),
    ("TERAPIA OCUPACIONAL", "terapiaocupacional.hcid@gmail.com"),
    ("TRANSPORTE EXTERNO", "centraltransporte.hcid@gmail.com"),
    ("UNIDADE SENTINELA", "unidadesentinela@gmail.com"),
    ("UTI A - MÉDICOS", "pedrohfrota@gmail.com"),
    ("UTI B - MÉDICOS", "livia.goreth@ufma.br"),
    ("UTI C - MÉDICOS", "hugogama19@hotmail.com"),
    ("ORTOPEDIA - MÉDICOS", "abraaojunior@hotmail.com"),
    ("PRONTO SOCORRO - MÉDICOS", "gustavo1422@hotmail.com"),
    ("ALA A - MÉDICOS", "kledsonmarques10@gmail.com"),
    ("ALA B - MÉDICOS", "aliomarcamaradecarvalho@gmail.com"),
    ("CENTRO CIRÚRGICO - MÉDICOS", "felipeabreu805@gmail.com"),
    ("CARDIOLOGIA - MÉDICOS", "jordach.xerez@gmail.com"),
    ("ANESTESISTAS - MÉDICO", "rafaelcbarrosneto@yahoo.com.br"),
    ("XEROX/REPROGRAFIA", "centraldecopias.hcid@gmail.com"),
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

def data_por_extenso(dt):
    meses = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
             "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    return f"{dt.day} de {meses[dt.month-1]} de {dt.year}"

def processar_turno(valor):
    if pd.isna(valor):
        return "", "", ""
    texto = str(valor).strip().upper()
    manha = "X" if "MANH" in texto else ""
    tarde = "X" if "TARD" in texto else ""
    noite = "X" if "NOIT" in texto else ""
    return manha, tarde, noite

def buscar_email_setor(nome_setor):
    """Busca o e-mail correspondente ao nome do setor"""
    if not nome_setor or str(nome_setor).strip() == "":
        return ""
    nome_setor = str(nome_setor).strip().upper()
    for _, linha in df_setores.iterrows():
        if nome_setor in linha["SETOR"].upper() or linha["SETOR"].upper() in nome_setor:
            return linha["E-MAIL"]
    return ""

def montar_corpo_email(dados_memorando, email_destino):
    """Monta o texto completo do e-mail com os dados do memorando"""
    assunto = f"NOTIFICAÇÃO DE INCIDENTE — Nº {dados_memorando['{{notificacao_n}}']}"
    
    corpo = f"""
Bom Dia / Boa Tarde, Prezado(a)!

Segue anexo o Memorando Nº {dados_memorando['{{numero_memorando}}']} referente à Notificação de Incidente Nº {dados_memorando['{{notificacao_n}}']}.

📋 INFORMAÇÕES:
• Setor: {dados_memorando['{{setor}}']}
• Paciente: {dados_memorando['{{nome_paciente}}']}
• Data da Ocorrência: {dados_memorando['{{data_ocorrencia}}']}
• Tipo de Incidente: {dados_memorando['{{tipo_incidente}}']}
• Classificação: {dados_memorando['{{classificacao_incidente}}']}

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
        
        st.download_button(
            label="📄 Baixar Memorando (DOCX)",
            data=saida_docx,
            file_name=f"{nome_arquivo_base}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        return saida_docx.getvalue()
    except FileNotFoundError:
        st.error("❌ Arquivo modelo_memorando.docx não encontrado!")
        return None
    except Exception as e:
        st.error(f"❌ Erro ao gerar documento: {str(e)}")
        return None

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
            
            # Extrair todos os campos da planilha
            num_memo = limpar_numero(linha.iloc[15]) if len(linha) > 15 else "S/N"
            dt_notif = formatar_data_br(linha.iloc[3]) if len(linha) > 3 else ""
            dt_ocorr = formatar_data_br(linha.iloc[2]) if len(linha) > 2 else ""
            local = str(linha.iloc[5]).strip() if len(linha) > 5 else ""
            tipo = str(linha.iloc[6]).strip() if len(linha) > 6 else ""
            classificacao = str(linha.iloc[7]).strip() if len(linha) > 7 else ""
            desc = str(linha.iloc[8]).strip() if len(linha) > 8 else ""
            leito = limpar_numero(linha.iloc[10]) if len(linha) > 10 else ""
            setor_notificante = str(linha.iloc[11]).strip() if len(linha) > 11 else ""
            sugestao = str(linha.iloc[12]).strip() if len(linha) > 12 else ""
            gestor = str(linha.iloc[13]).strip() if len(linha) > 13 else "GESTOR DE ENFERMAGEM"
            setor = str(linha.iloc[14]).strip() if len(linha) > 14 else ""
            turno_bruto = str(linha.iloc[4]).strip() if len(linha) > 4 else ""
            
            # Buscar e-mail do setor
            email_destino = buscar_email_setor(setor)
            
            # Processar turno
            t_manha, t_tarde, t_noite = processar_turno(turno_bruto)
            
            # Montar dicionário com TODAS as tags
            dados_memorando = {
                "{{numero_memorando}}": num_memo,
                "{{gestor}}": gestor,
                "{{setor}}": setor,
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
                "{{m}}": t_manha,
                "{{t}}": t_tarde,
                "{{n}}": t_noite,
            }
            
            # Montar assunto e corpo do e-mail
            assunto_email, corpo_email = montar_corpo_email(dados_memorando, email_destino)
            
            with st.expander(f"📋 Notificação {num_notif} — {nome_paciente}"):
                st.write(f"**📝 Memorando:** {num_memo} | **📅 Data:** {data_extenso}")
                st.write(f"**🏥 Setor:** {setor} | **👤 Gestor:** {gestor}")
                
                if email_destino:
                    st.info(f"**📧 E-mail do Setor:** {email_destino}")
                else:
                    st.warning("⚠️ E-mail não encontrado para este setor")
                
                # Gerar botão de download do DOCX
                gerar_e_exibir_documento(dados_memorando, f"MEMORANDO_{num_notif}")
                
                # Exibir assunto e corpo do e-mail
                st.subheader("📧 E-mail — Copie e Cole")
                st.write("**Assunto:**")
                st.code(assunto_email, language=None)
                st.write("**Corpo:**")
                st.text_area("Texto do E-mail", corpo_email, height=250, key=f"email_{idx}")
    
    except Exception as e:
        st.error(f"❌ Erro ao processar a planilha: {str(e)}")
else:
    st.info("👆 Faça upload da sua planilha Excel (.xlsx) para começar")
