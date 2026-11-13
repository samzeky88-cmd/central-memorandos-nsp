# =================================================================
# 📄 RENDERIZAÇÃO DE CADA LINHA/MEMORANDO
# =================================================================
@st.fragment
def renderizar_linha_paciente_sob_demanda(index, linha, num_colunas, data_extenso_envio):
    num_notif = limpar_numero_float(linha.iloc[0]) if num_colunas > 0 else "S-N"
    if num_notif.upper() == "STATUS" or "NOTIF" in num_notif.upper() or num_notif == "" or num_notif == "1":
        return
    status_envio = str(linha.iloc[1]).strip().upper() if num_colunas > 1 else ""
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
    if num_memo_cru == "" or num_memo_cru.upper() == "N° MEMO 01":
        num_memo_cru = "S-N"
    email_destino_planilha = tratar_str_limpa(linha.iloc[21]) if num_colunas > 21 else ""
    if email_destino_planilha.upper() == "EMAIL_SETOR":
        return
    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("No", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
    nome_base_arquivo = f"MEMORANDO N° {num_memo_limpo}_NOTIFICAÇÃO_N° {num_notif}_I_NSP"
    marca_manha = "X" if "MANH" in turno_planilha else ""
    marca_tarde = "X" if "TARD" in turno_planilha else ""
    marca_noite = "X" if "NOIT" in turno_planilha else ""
    if setor_notificante_bruto == "":
        setor_notificante_bruto = "NSP - NÚCLEO DE SEGURANÇA DO PACIENTE"
    
    # ✅ DICIONÁRIO COMPLETO E FECHADO
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
        "{{leito_paciente}}": leito_paciente,
        "{{setor_notificante}}": setor_notificante_bruto,
        "{{sugestao_nsp}}": sugestao_nsp,
        "{{data_envio}}": data_extenso_envio,
        "{{manha}}": marca_manha,
        "{{tarde}}": marca_tarde,
        "{{noite}}": marca_noite
    }  # ← Chave de fechamento que estava faltando!

    # ✅ Resto da função continua aqui...
    with st.expander(f"📋 {nome_do_paciente} — Notificação Nº {num_notif}"):
        col1, col2 = st.columns([3, 1])
        with col1:
            st.write(f"**Memorando:** {num_memo_cru} | **Notificação:** {num_notif}")
            st.write(f"**Setor Destino:** {setor_notificado_final} | **Gestor:** {gestor_destinatario}")
        with col2:
            if status_envio == "ENVIADO":
                st.success("✅ ENVIADO")
            else:
                st.info("📄 PENDENTE")

        # Gerar DOCX
        try:
            with open(caminho_modelo, "rb") as f:
                doc_bytes = f.read()
            doc = Document(io.BytesIO(doc_bytes))
            substituir_texto_protegendo_logos(doc, dados_memorando)
            saida_docx = io.BytesIO()
            doc.save(saida_docx)
            saida_docx.seek(0)
            docx_bytes = saida_docx.getvalue()

            col_a, col_b, col_c = st.columns(3)
            with col_a:
                st.download_button(
                    label="📥 Baixar DOCX",
                    data=saida_docx,
                    file_name=f"{nome_base_arquivo}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"docx_{index}"
                )

            # Converter para PDF
            with st.spinner("📄 Convertendo para PDF..."):
                pdf_bytes, erro_pdf = converter_docx_para_pdf(docx_bytes)
                if pdf_bytes:
                    with col_b:
                        st.download_button(
                            label="📥 Baixar PDF",
                            data=pdf_bytes,
                            file_name=f"{nome_base_arquivo}.pdf",
                            mime="application/pdf",
                            key=f"pdf_{index}"
                        )
                else:
                    st.warning(f"⚠️ Não foi possível gerar PDF: {erro_pdf}")

            # Área de Envio de E-mail
            st.markdown("---")
            st.subheader("✉️ Enviar por E-mail")

            # Escolher destinatário
            opcoes_setores = df_setores["SETOR"].tolist()
            email_setor_escolhido = st.selectbox(
                "Selecione o Setor",
                opcoes_setores,
                index=opcoes_setores.index(setor_notificado_final) if setor_notificado_final in opcoes_setores else 0,
                key=f"setor_{index}"
            )
            email_destino = df_setores[df_setores["SETOR"] == email_setor_escolhido]["E-MAIL"].values[0]
            st.info(f"📧 E-mail: **{email_destino}**")

            # Corpo e Assunto
            assunto_email = f"Notificação: N° {num_notif} _ Memorando: Nº {num_memo_cru}"
            saudacao = obter_saudacao()
            corpo_email = f"""{saudacao},

Segue em anexo o Memorando Nº {num_memo_cru} para ser analisado e respondido (via e-mail) em até 15 dias após a data presente.

ATENÇÃO: A resposta via e-mail deve constar um arquivo em forma de word ou PDF para arquivamento de respostas conforme rotina institucional. Não serão aceitas mensagens via e-mail sem arquivo como resposta.

Segue abaixo a notificação para análise do incidente em equipe e resposta ao NSP:
• Memorando: Nº {num_memo_cru}
• Notificação: N° {num_notif}

Atenciosamente,
{CONFIG_EMAIL['assinatura']}
"""
            with st.expander("👁️ Prévia do e-mail"):
                st.code(f"Assunto: {assunto_email}")
                st.text(corpo_email)

            # Carregar roteiro fixo
            roteiro_bytes = carregar_roteiro_fixo()
            if not roteiro_bytes:
                st.warning("⚠️ Arquivo ROTEIRO_COORDENADORES.pdf não encontrado — apenas o memorando será anexado")

            # Botão Enviar
            if st.button(f"📤 Enviar E-mail para {email_setor_escolhido}", type="primary", key=f"enviar_{index}"):
                anexos = [(f"{nome_base_arquivo}.pdf", pdf_bytes, "application/pdf")]
                if roteiro_bytes:
                    anexos.append(("1-ROTEIRO_COORDENADORES.pdf", roteiro_bytes, "application/pdf"))

                with st.spinner("📤 Enviando e-mail..."):
                    sucesso, mensagem = enviar_email(email_destino, assunto_email, corpo_email, anexos)
                    if sucesso:
                        st.success(mensagem)
                    else:
                        st.error(mensagem)

        except FileNotFoundError:
            st.error(f"❌ Modelo não encontrado: coloque 'modelo_memorando.docx' na pasta do app")
        except Exception as e:
            st.error(f"❌ Erro ao gerar memorando: {str(e)}")
