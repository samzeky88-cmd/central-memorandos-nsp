            localizacao = tratar_texto(vals[5]) if len(vals) > 5 else ""
            tipo_incidente = tratar_texto(vals[6]) if len(vals) > 6 else ""
            classificacao_incidente = tratar_texto(vals[7]) if len(vals) > 7 else ""
            descricao_notificacao = tratar_texto(vals[8]) if len(vals) > 8 else ""
            leito = tratar_texto(vals[10]) if len(vals) > 10 else ""
            setor_notificante = tratar_texto(vals[11]) if len(vals) > 11 else ""
            sugestao = tratar_texto(vals[12]) if len(vals) > 12 else ""

            m_manha = "☐"
            m_tarde = "☐"
            m_noite = "☐"
            if turno:
                if "MANHÃ" in turno or "MANHA" in turno:
                    m_manha = "☑"
                elif "TARDE" in turno:
                    m_tarde = "☑"
                elif "NOITE" in turno:
                    m_noite = "☑"

            blocos_memo = []
            for g_col, s_col, m_col in [(13, 14, 15), (17, 18, 19), (22, 23, 24), (26, 27, 28)]:
                if m_col >= len(vals):
                    continue
                num_memo = limpar_numero(vals[m_col])
                setor_destino = tratar_texto(vals[s_col])
                gestor = tratar_texto(vals[g_col])

                if num_memo and num_memo != "" and "Nº MEMO" not in num_memo:
                    blocos_memo.append({
                        "num_memo": num_memo,
                        "setor": setor_destino,
                        "gestor": gestor
                    })

            if not blocos_memo:
                continue

            if status_registro == "ENVIADO":
                st.markdown(f"~~Notificação {num_notif} — {nome_paciente}~~ **ENVIADO**")
                continue

            qtd_processados += 1

            with st.expander(f"Notificação {num_notif} — {nome_paciente} ({len(blocos_memo)} memo(s))", expanded=True):
                st.write(f"Data de Emissão: {data_extenso}")
                st.write(f"Ocorrência: {data_ocorrencia} | Notificação: {data_notificacao}")
                st.write(f"Local: {localizacao} | Setor Notificante: {setor_notificante}")
                st.write(f"Tipo: {tipo_incidente} | Classificação: {classificacao_incidente}")
                st.write(f"Paciente: {nome_paciente} | Leito: {leito}")
                st.write(f"Turno: {m_manha} MANHÃ {m_tarde} TARDE {m_noite} NOITE")
                if descricao_notificacao:
                    st.write(f"Descrição: {descricao_notificacao}")
                if sugestao:
                    st.write(f"Sugestão: {sugestao}")

                for b in blocos_memo:
                    email_destino = buscar_email_setor(b["setor"])
                    st.subheader(f"MEMO: NSP Nº {b['num_memo']}/{ano_atual}")
                    st.write(f"Para: {b['gestor']} da {b['setor']}")
                    st.write(f"Assunto: Nº {num_notif}")
                    st.info(f"E-mail: {email_destino}")

                    corpo_email = "Boa Tarde, Prezados, ou bom dia de acordo com o horário.\n\n"
                    corpo_email += "Segue em anexo o Memorando Nº " + b["num_memo"] + "/" + str(ano_atual)
                    corpo_email += " referente à notificação Nº " + num_notif + ", para ser analisado e "
                    corpo_email += "respondido (via e-mail) em até 15 dias após a data presente.\n\n"
                    corpo_email += "ATENÇÃO: A resposta via e-mail deve constar um arquivo em forma de Word "
                    corpo_email += "ou PDF para arquivamento de respostas conforme rotina institucional. "
                    corpo_email += "Não serão aceitas mensagens via e-mail sem arquivo como resposta.\n\n"
                    corpo_email += "Atenciosamente,\n\nFabrício S. Santos\nNúcleo de Segurança do Paciente\nnspsoc2@gmail.com"

                    with st.expander("Texto do E-mail"):
                        st.text(corpo_email)

                    dados_docx = {
                        "{{numero_memorando}}": b["num_memo"] + "/" + str(ano_atual),
                        "{{notificacao_n}}": num_notif,
                        "{{gestor}}": b["gestor"],
                        "{{setor}}": b["setor"],
                        "{{data_ocorrencia}}": data_ocorrencia,
                        "{{data_notificacao}}": data_notificacao,
                        "{{data_envio}}": data_extenso,
                        "{{m}}": m_manha,
                        "{{t}}": m_tarde,
                        "{{n}}": m_noite,
                        "{{localizacao}}": localizacao,
                        "{{tipo_incidente}}": tipo_incidente,
                        "{{classificacao_incidente}}": classificacao_incidente,
                        "{{descricao_notificacao}}": descricao_notificacao,
                        "{{nome_paciente}}": nome_paciente,
                        "{{leito}}": leito,
                        "{{setor_notificante}}": setor_notificante,
                        "{{sugestao}}": sugestao
                    }

                    arquivo_pronto = gerar_docx(dados_docx)
                    if arquivo_pronto:
                        st.download_button(
                            label="📄 Baixar Memorando",
                            data=arquivo_pronto,
                            file_name=f"NSP_Nº {num_notif}_{ano_atual}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"download_{idx}_{b['num_memo']}"
                        )

    except Exception as e:
        st.error(f"Erro ao processar planilha: {str(e)}")
