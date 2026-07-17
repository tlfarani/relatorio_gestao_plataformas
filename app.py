# =========================================================================
        # ABA 4: CONSOLIDAÇÃO POR PRODUTO (FILTROS ATÔMICOS CORRIGIDOS)
        # =========================================================================
        with tab_produtos:
            st.markdown("### 🛢️ Inventário Consolidado de Produtos e Marcas Comerciais")
            st.write("Agregação unificada, padronizada e higienizada de produtos com vazamento em plataformas (2025).")
            st.write("---")
            
            # --- 1. FUNÇÕES INTERNAS DE HIGIENIZAÇÃO E PADRONIZAÇÃO ---
            def limpar_volume_safely(val):
                if pd.isna(val): return 0.0
                if isinstance(val, (int, float)): return float(val)
                val_str = str(val).strip()
                if val_str.upper() == 'PREENCHER' or val_str == '': return 0.0
                if 'E' in val_str.upper():
                    val_str = val_str.replace(',', '.')
                    try: return float(val_str)
                    except ValueError: pass
                if '.' in val_str and ',' in val_str:
                    val_str = val_str.replace('.', '').replace(',', '.')
                elif ',' in val_str:
                    val_str = val_str.replace(',', '.')
                try: return float(val_str)
                except ValueError: return 0.0

            def padronizar_nome_produto(nome):
                if pd.isna(nome): return "Não Informado"
                n = str(nome).strip()
                n_lower = n.lower()
                if n_lower.startswith('erifon'): return "Erifon HD 603 HP > 1,89%"
                if n_lower.startswith('stack'): return "Stack Magic Eco F ≥ 1%"
                if 'panolin' in n_lower: return "Panolins"
                if 'monoetilenoglicol' in n_lower or 'meg' in n_lower: return "Monoetilenoglicol"
                if 'br-mul' in n_lower or 'br_mul' in n_lower or 'brmul' in n_lower: return "BR-Mul"
                if 'água oleosa' in n_lower or 'agua oleosa' in n_lower: return "Água Oleosa"
                if 'petroleo' in n_lower or 'petróleo' in n_lower: return "Petróleo"
                if 'óleo diesel' in n_lower or 'oleo diesel' in n_lower: return "Óleo Diesel"
                if 'mobil dte' in n_lower or 'mobildte' in n_lower: return "Mobil DTE"
                if 'lubrax' in n_lower: return "Lubrax"
                if 'hyspin' in n_lower: return "Hyspin"
                if 'mobilgear' in n_lower: return "Mobilgear"
                if 'oceanic' in n_lower and '525' in n_lower: return "Oceanic HW 525"
                if 'oceanic' in n_lower and '443' in n_lower: return "Oceanic HW 443"
                if 'tellus' in n_lower: return "Shell Tellus"
                if 'transaqua' in n_lower: return "Transaqua DW"
                if 'fcba' in n_lower or 'completação aquoso' in n_lower or 'completação base água' in n_lower or 'completação base agua' in n_lower:
                    return "FCBA (Fluido de Completação de Base Aquosa)"
                if 'fpba' in n_lower or ('perfuração' in n_lower and 'base aquosa' in n_lower):
                    return "FPBA (Fluido de Perfuração de Base Aquosa)"
                if 'produto oleoso' in n_lower or n_lower in ['óleo lubrificante', 'oleo lubrificante']:
                    return "Produto Oleoso Genérico"
                return n

            # --- 2. MAPEAMENTO REAL DE EQUIPAMENTOS ATÔMICOS (ÚNICOS) ---
            equipamentos_unicos_filtro = set()
            if 'equipment' in df_plataformas_2025.columns:
                for eq_row in df_plataformas_2025['equipment']:
                    eq_str = str(eq_row).strip()
                    if pd.isna(eq_row) or eq_str == '' or eq_str.lower() == 'nan' or eq_str.lower() == 'não informado':
                        equipamentos_unicos_filtro.add('Não Informado')
                    else:
                        # Quebra as strings compostas por vírgula em itens atômicos separados
                        for item in eq_str.split(','):
                            item_clean = item.strip()
                            if item_clean:
                                equipamentos_unicos_filtro.add(item_clean)
            list_equip_ordenada = sorted(list(equipamentos_unicos_filtro))

            # --- 3. SEÇÃO DE FILTROS AVANÇADOS DA ABA ---
            st.subheader("Filtros Específicos de Produtos")
            col_p1, col_p2, col_p3 = st.columns(3)
            
            with col_p1:
                equip_selecionados = st.multiselect("Filtrar por Equipamento Envolvido (Itens Únicos):", options=list_equip_ordenada, default=list_equip_ordenada, key="ms_equip_p4")
            with col_p2:
                classes_selecionadas = st.multiselect("Filtrar por Classe de Risco:", options=["A obter"], default=["A obter"], key="ms_classe_p4")
            with col_p3:
                tipos_selecionados = st.multiselect("Filtrar por Tipo de Produto:", options=["A obter"], default=["A obter"], key="ms_tipo_p4")
            
            # --- 4. FILTRAGEM DOS ACIDENTES NA ORIGEM (EVITA DUPLICAR VOLUMES) ---
            def verificar_aderencia_equipamento(eq_val, selecionados):
                eq_str = str(eq_val).strip()
                if pd.isna(eq_val) or eq_str == '' or eq_str.lower() == 'nan' or eq_str.lower() == 'não informado':
                    return 'Não Informado' in selecionados
                items_do_acidente = [item.strip() for item in eq_str.split(',') if item.strip()]
                return any(item in selecionados for item in items_do_acidente)

            # Aplica a aderência de equipamentos
            df_plataformas_filtradas_p4 = df_plataformas_2025[
                df_plataformas_2025['equipment'].apply(lambda x: verificar_aderencia_equipamento(x, equip_selecionados))
            ].copy()
            
            # Validação dos placeholders "A obter"
            if "A obter" not in classes_selecionadas or "A obter" not in tipos_selecionados:
                df_plataformas_filtradas_p4 = df_plataformas_filtradas_p4.iloc[0:0]

            # --- 5. EXTRAÇÃO E UNIFICAÇÃO DOS PRODUTOS DA BASE FILTRADA ---
            registros_produtos = []
            for _, row in df_plataformas_filtradas_p4.iterrows():
                equipamento_atual = str(row.get('equipment', 'Não Informado')).strip()
                id_processo = str(row.get('num_processo', 'S/N'))
                
                classe_risco_atual = "A obter"
                tipo_atual = "A obter"
                
                for prefix in ['1', '2', '3']:
                    marca = str(row.get(f'marca_p{prefix}')).strip() if pd.notna(row.get(f'marca_p{prefix}')) else ''
                    if marca != '' and marca.upper() != 'PREENCHER' and marca.lower() != 'nan':
                        marca_padrao = padronizar_nome_produto(marca)
                        vol = limpar_volume_safely(row.get(f'qtd_p{prefix}'))
                        registros_produtos.append({
                            'Produto': marca_padrao, 
                            'Volume': vol, 
                            'Equipamento': equipamento_atual, 
                            'Classe de Risco': classe_risco_atual,
                            'Tipo': tipo_atual,
                            'Processo': id_processo
                        })
            
            if registros_produtos:
                df_prod_filtrado = pd.DataFrame(registros_produtos)
                
                st.write("---")
                # --- 6. INDICADOR: CONTAGEM DE PRODUTOS DIFERENTES ---
                col_metric, _ = st.columns([1, 2])
                with col_metric:
                    produtos_distintos = df_prod_filtrado['Produto'].nunique()
                    st.metric("Total de Produtos Distintos Filtrados", f"{produtos_distintos}")
                st.write("---")
                
                # --- 7. AGREGAÇÃO CONSOLIDADA SEM REPETIÇÕES DE STRINGS ---
                def obter_lista_equipamentos_unicos(series):
                    set_consolidado = set()
                    for texto_celula in series:
                        if pd.notna(texto_celula):
                            for item in str(texto_celula).split(','):
                                cleaned = item.strip()
                                if cleaned and cleaned.lower() not in ['nan', 'não informado']:
                                    set_consolidado.add(cleaned)
                    if not set_consolidado:
                        return "Não Informado"
                    return ", ".join(sorted(set_consolidado))

                df_prod_summary = df_prod_filtrado.groupby(['Produto', 'Classe de Risco', 'Tipo']).agg(
                    Qtd_Acidentes=('Processo', 'nunique'),
                    Vol_Total=('Volume', 'sum'),
                    Equipamentos_Lista=('Equipamento', obter_lista_equipamentos_unicos)
                ).reset_index()
                
                # Layout final da tabela corporativa
                df_prod_summary = df_prod_summary[[
                    'Produto', 'Qtd_Acidentes', 'Vol_Total', 'Classe de Risco', 'Tipo', 'Equipamentos_Lista'
                ]]
                
                df_prod_summary.columns = [
                    'Nome do Produto', 'Quantidade de Acidentes', 'Soma dos Volumes', 'Classe de Risco', 'Tipo', 'Equipamentos Envolvidos'
                ]
                
                df_prod_summary = df_prod_summary.sort_values(by='Nome do Produto')
                
                df_formatado_ibama = df_prod_summary.style.format(
                    {'Soma dos Volumes': '{:,.8f}'}, 
                    decimal=',', 
                    thousands='.'
                )
                st.dataframe(df_formatado_ibama, use_container_width=True)
            else:
                st.info("Nenhum produto corresponde aos critérios dos filtros selecionados.")
