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
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
