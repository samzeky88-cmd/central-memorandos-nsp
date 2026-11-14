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
    num_notif = limpar_numero_float(linha.iloc) if num_colunas > 0 else "S-N"
    if num_notif.upper() == "STATUS" or "NOTIF" in num_notif.upper() or num_notif == "" or num_notif == "1":
        return None, None, 0

    status_envio = str(linha.iloc).strip().upper() if num_colunas > 1 else "PENDENTE"
    dt_ocorr = formatar_data_br(linha.iloc) if num_colunas > 2 else ""
    dt_notif = formatar_data_br(linha.iloc) if num_colunas > 3 else ""
    turno_planilha = str(linha.iloc).strip().upper() if num_colunas > 4 else ""
    onde_ocorreu = tratar_str_limpa(linha.iloc) if num_colunas > 5 else "Não informado"
    tipo_incidente = tratar_str_limpa(linha.iloc) if num_colunas > 6 else ""
    classificacao_incidente = tratar_str_limpa(linha.iloc) if num_colunas > 7 else ""
    descricao_notificacao = tratar_str_limpa(linha.iloc) if num_colunas > 8 else ""
    nome_do_paciente = tratar_str_limpa(linha.iloc) if num_colunas > 9 else "Paciente Não Identificado"

    if nome_do_paciente.upper() == "PACIENTE":
        return None, None, 0

    leito_paciente = limpar_numero_float(linha.iloc) if num_colunas > 10 else ""
    setor_notificante_bruto = tratar_str_limpa(linha.iloc) if num_colunas > 11 else ""
    sugestao_nsp = tratar_str_limpa(linha.iloc) if num_colunas > 12 else ""

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
