# Execução principal após upload do arquivo
if arquivo_excel:
    nome_chave_cache = f"dados_memos_{arquivo_excel.name}_{data_selecionada.strftime('%Y%m%d')}"
    
    if nome_chave_cache not in st.session_state:
        with st.spinner("📦 Processando e estruturando os memorandos na memória... Aguarde um instante."):
            try:
                df = pd.read_excel(arquivo_excel, header=None)
                
                # Alerta caso a planilha esteja completamente em branco
                if df.empty:
                    st.warning("⚠️ O arquivo Excel enviado está vazio.")
                    st.stop()
                    
                data_extenso_envio = obter_data_por_extenso(data_selecionada)
                arquivos_processados = []
                
                hora_atual = datetime.now().hour
                saudacao = "Bom dia Prezados" if hora_atual < 12 else "Boa Tarde Prezados"
                num_colunas_df = len(df.columns)
                
                # Índice de busca fixado para evitar erros de escopo
                c_busca = 0 
                
                for index, line in df.iterrows():
                    try:
                        # Ignora linhas que não possuem dados válidos na Coluna B
                        if index not in df.index or len(df.iloc[index]) < 2:
                            continue
                            
                        val_paciente = df.iloc[index, 1]  # Coluna B
                        texto_paciente_teste = str(val_paciente).strip().upper()
                        
                        if pd.isna(val_paciente) or texto_paciente_teste == "" or texto_paciente_teste == "NAN" or "PACIENTE" in texto_paciente_teste or "NOME" in texto_paciente_teste:
                            continue
                        nome_do_paciente = str(val_paciente).strip()
                    except Exception as e:
                        # Se falhar na identificação do paciente, pula para a próxima linha do Excel
                        continue
                    
                    # Captura individual de cada coluna com travas de segurança para planilhas irregulares
                    try: num_notif = limpar_numero_float(df.iloc[index, 0]) if num_colunas_df > 0 else str(index + 1)
                    except: num_notif = str(index + 1)
                    
                    try: dt_notif = formatar_data_br(df.iloc[index, 4]) if num_colunas_df > 4 else ""
                    except: dt_notif = ""
                    
                    try: dt_ocorr = formatar_data_br(df.iloc[index, 5]) if num_colunas_df > 5 else ""
                    except: dt_ocorr = ""
                    
                    try: turno_planilha = str(df.iloc[index, 6]).strip().upper() if num_colunas_df > 6 else ""
                    except: turno_planilha = ""
                    
                    try: tipo_incidente = limpar_nan(df.iloc[index, 10]).replace("_", " ") if num_colunas_df > 10 else ""
                    except: tipo_incidente = ""
                    
                    try: texto_sugestao = str(df.iloc[index, 12]).strip() if num_colunas_df > 12 else ""
                    except: texto_sugestao = ""
                    
                    try: memo_01 = str(df.iloc[index, 15]).strip() if num_colunas_df > 15 else ""
                    except: memo_01 = ""
                    
                    try: memo_02 = str(df.iloc[index, 19]).strip() if num_colunas_df > 19 else ""
                    except: memo_02 = ""
                    
                    if memo_01 == "" or memo_01.lower() == "nan" or memo_01.lower() == "none":
                        num_memo_cru = memo_02 if memo_02 != "" and memo_02.lower() != "nan" and memo_02.lower() != "none" else "S-N"
                    else:
                        num_memo_cru = memo_01
                        
                    num_memo_limpo = num_memo_cru.replace("Nº", "").replace("NS", "").replace("NSP", "").replace("/", "-").replace(" ", "").strip()
                    nome_base_arquivo = f"MEMORANDO Nº {num_memo_limpo}_NOTIFICAÇÃO_Nº {num_notif}_I_NSP"
                    
                    marca_manha = "X" if "MANH" in turno_planilha else " "
                    marca_tarde = "X" if "TARD" in turno_planilha else " "
                    marca_noite = "X" if "NOIT" in turno_planilha else " "
                    
                    gestor_val, setor_val, onde_val, classif_val, setor_notif_val, leito_val = "", "", "", "", "", ""
                    
                    if num_colunas_df > 0:
                        for r_busca in range(min(index, 15)):
                            for c_temp in range(num_colunas_df):
                                txt_c = str(df.iloc[r_busca, c_temp]).strip().upper()
                                if "GESTOR" in txt_c: gestor_val = str(df.iloc[index, c_temp])
                                elif "SETOR NOTIFICADO" in txt_c: setor_val = str(df.iloc[index, c_temp])
                                elif "ONDE OCORREU" in txt_c: onde_val = str(df.iloc[index, c_temp])
                                elif "CLASSIFICA" in txt_c: classif_val = str(df.iloc[index, c_temp])
                                elif "SETOR NOTIFICANTE" in txt_c: setor_notif_val = str(df.iloc[index, c_temp])
                                elif "LEITO" in txt_c: leito_val = str(df.iloc[index, c_temp])
                                
                    if not gestor_val: gestor_val = "GESTOR DE ENFERMAGEM"
                    if not setor_val: setor_val = "ALA B"
                    
                    email_destino = ""
                    try:
                        if num_colunas_df > c_busca:
                            email_destino = limpar_nan(df.iloc[index, c_busca])
                    except:
                        email_destino = ""
                    
                    texto_descricao = limpar_nan(df.iloc[index, 11]) if num_colunas_df > 11 else ""
                    texto_email_formatado = f"{saudacao},\n\nSegue em anexo o Memorando {num_memo_cru} referente à Notificação {num_notif}.\n\nPaciente: {nome_do_paciente}\nDescrição do ocorrido: {texto_descricao}"
                    assunto_email = f"NSP - Memorando {num_memo_cru} (Notificação {num_notif})"
                    
                    link_gmail = (
                        f"https://google.com?"
                        f"view=cm&fs=1&tf=1"
                        f"&to={urllib.parse.quote(email_destino)}"
                        f"&su={urllib.parse.quote(assunto_email)}"
                        f"&body={urllib.parse.quote(texto_email_formatado)}"
                    )
                    
                    dados_memorando = {
                        "{{numero_memorando}}": num_memo_cru, "{{gestor}}": limpar_nan(gestor_val),
                        "{{setor}}": limpar_nan(setor_val), "{{notificacao_n}}": num_notif,
                        "{{data_notificacao}}": dt_notif, "{{data_ocorrencia}}": dt_ocorr,
                        "{{localizacao}}": limpar_nan(onde_val) if onde_val else "Ala B",
                        "{{classificacao_incidente}}": limpar_nan(classif_val) if classif_val else "Incidente com dano moderado",
                        "{{setor_notificante}}": limpar_nan(setor_notif_val) if setor_notif_val else "SALA VERMELHA",
                        "{{tipo_incidente}}": tipo_incidente, "{{descricao_notificacao}}": texto_descricao,
                        "{{nome_paciente}}": nome_do_paciente,
                        "{{leito}}": limpar_numero_float(leito_val) if leito_val else (limpar_numero_float(df.iloc[index, 7]) if num_colunas_df > 7 else ""),
                        "{{sugestao}}": limpar_nan(texto_sugestao), "{{m}}": marca_manha, "{{t}}": marca_tarde, "{{n}}": marca_noite,
                        "{{data_envio}}": data_extenso_envio
                    }
                    
                    if os.path.exists(caminho_modelo):
                        doc_instancia = Document(caminho_modelo)
                        substituir_texto_protegendo_logos(doc_instancia, dados_memorando)
                        
                        buffer_bytes = io.BytesIO()
                        doc_instancia.save(buffer_bytes)
                        bytes_word = buffer_bytes.getvalue()
                        
                        arquivos_processados.append({
                            "paciente": nome_do_paciente,
                            "nome_arquivo": nome_base_arquivo,
                            "conteudo_word": bytes_word,
                            "link_email": link_gmail
                        })
                    else:
                        st.error(f"❌ Arquivo de modelo não encontrado no caminho: {caminho_modelo}")
                        st.stop()
                
                # Validação final: avisa se passou pelo loop mas nenhuma linha virou memorando
                if len(arquivos_processados) == 0:
                    st.warning("⚠️ Nenhuma linha válida encontrada na planilha. Verifique se os nomes dos pacientes estão localizados estritamente na segunda coluna (Coluna B).")
                
                st.session_state[nome_chave_cache] = arquivos_processados
                
            except Exception as erro_geral:
                st.error(f"💥 Ocorreu uma falha crítica no processamento dos dados: {str(erro_geral)}")

    # --- RENDERIZAÇÃO DA INTERFACE ---
    if nome_chave_cache in st.session_state and len(st.session_state[nome_chave_cache]) > 0:
        st.success(f"📋 Lista de verificação pronta! {len(st.session_state[nome_chave_cache])} memorandos estruturados.")
        
        for idx, item in enumerate(st.session_state[nome_chave_cache]):
            col_nome, col_word, col_pdf, col_email = st.columns(4)
            
            with col_nome:
                st.markdown(f"🔹 {item['paciente']}")
                
