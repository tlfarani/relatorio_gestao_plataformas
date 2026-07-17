import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# Configuração da página
st.set_page_config(
    page_title="Painel de Acidentes - Plataformas",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- ESTILIZAÇÃO DO TEMA DO IBAMA (CSS PREMIUM AJUSTADO) ---
st.markdown("""
<style>
    /* 1. Cor de fundo principal da página (Cinza Claro) */
    .stApp {
        background-color: #F4F6F4 !important;
    }
    
    /* 2. Garante contraste escuro para textos normais e rótulos */
    .stApp p, .stApp span, .stApp label {
        color: #2E3E2F !important;
    }

    /* 3. Estilo dos títulos e subtítulos (Verde Musgo Institucional) */
    h1, h2, h3, h4, h5, h6 {
        color: #1E4620 !important;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif !important;
        font-weight: bold !important;
    }
    
    /* 4. Rótulos de todos os widgets e filtros */
    div[data-testid="stWidgetLabel"] p {
        color: #1E4620 !important;
        font-weight: bold !important;
        font-size: 16px !important;
    }
    
    /* 5. Customização das caixas de seleção (Multiselect) */
    div[data-baseweb="select"] > div {
        background-color: #FFFFFF !important;
        color: #2E3E2F !important;
        border-color: #C2CDC2 !important;
    }
    
    /* 6. Estilo geral das tags selecionadas (chips) */
    [data-baseweb="tag"] {
        background-color: #D1E2D3 !important; 
        border: 1px solid #1E4620 !important; 
        border-radius: 4px !important;
    }
    [data-baseweb="tag"] * {
        color: #1E4620 !important;
        font-weight: bold !important;
    }
    
    /* 7. Arredondamento e estilo dos cards de métricas (KPIs no topo) */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF !important;
        border-radius: 12px !important;
        padding: 20px !important;
        border: 1px solid #E2E8E2 !important;
        box-shadow: 0 4px 6px rgba(0,0,0,0.03) !important;
    }
    div[data-testid="stMetricLabel"] p {
        color: #4A5D4E !important;
        font-weight: bold !important;
        font-size: 16px !important;
    }
    div[data-testid="stMetricValue"] > div {
        color: #1E4620 !important;
        font-weight: bold !important;
    }
    
    /* 8. ARREDONDAMENTO DAS BORDAS DAS FIGURAS */
    .stPlotlyChart {
        background-color: #FFFFFF !important;
        border-radius: 14px !important;
        border: 1px solid #E2E8E2 !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.04) !important;
        overflow: hidden !important; 
    }
    
    /* 9. Customização do painel lateral de filtros (Sidebar) */
    [data-testid="stSidebar"] {
        background-color: #E2E8E2 !important;
        border-right: 1px solid #C2CDC2;
    }

    /* 10. Customização das abas de navegação (Tabs) */
    button[data-baseweb="tab"] {
        color: #6E8B75 !important;
        font-size: 15px !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #1E4620 !important;
        border-bottom-color: #1E4620 !important;
        font-weight: bold !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("⚓ Consolidação de Acidentes em Plataformas de Petróleo")
st.markdown("Análise interativa e monitoramento de emergências ambientais baseadas em dados consolidados.")

NOME_ACIDENTES = "acidentes_2025.xlsx"
NOME_PRODUCAO = "Producao.xlsx"
NOME_ATENDIMENTO = "Tempo_Atendimento.xlsx"
NOME_ENCERRAMENTO = "Tempo_Encerramento.xlsx"
NOME_MAP_PRODUTOS = "produtos_consolidados.xlsx"

# --- FUNÇÕES DE LEITURA E LIMPEZA DE DADOS (COM CACHE) ---

@st.cache_data(ttl=1800)
def carregar_dados_2025(caminho):
    df = pd.read_excel(caminho, sheet_name="Geral")
    df.columns = df.columns.str.strip()
    
    dicionario_colunas = {
        "Processo SEI": "num_processo",
        "Endereço (especificar plataforma, navio, km da rodovia, ferrovia e duto)": "instalacao",
        "Bacia Sedimentar": "bacia_sedimentar",
        "Origem do acidente": "origem_acidente",
        "Marca Comercial Produto 1": "marca_p1",
        "Produto 1": "prod_1",
        "Quantidade do vazamento (Produto 1)": "qtd_p1",
        "Unidade do vazamento (Produto 1)": "uni_p1",
        "Marca Comercial Produto 2": "marca_p2",
        "Produto 2": "prod_2",
        "Quantidade do vazamento (Produto 2)": "qtd_p2",
        "Unidade do vazamento (Produto 2)": "uni_p2",
        "Marca Comercial Produto 3": "marca_p3",
        "Produto 3": "prod_3",
        "Quantidade do vazamento (Produto 3)": "qtd_p3",
        "Unidade do vazamento (Produto 3)": "uni_p3",
        "Causa do Evento Padronizada 1 (indicado no Relat. Pós-Acidente ou outro documento)": "causa_1",
        "Causa do Evento Padronizada 2 (indicado no Relat. Pós-Acidente ou outro documento)": "causa_2",
        "Causa do Evento Padronizada 3 (indicado no Relat. Pós-Acidente ou outro documento)": "causa_3",
        "Empresa": "empresa",
        "Acionado PEI ou similar?": "acionamento_pei",
        "Tempo de atendimento (em dias)": "tempo_atendimento",
        "Forma de atendimento (primeiro documento/ação)": "forma_atendimento",
        "Dias Até Encerramento da Investigação": "dias_encerramento",
        "Operação/ocorrência/sistema": "operacao",
        "Equipamento/sistema envolvido": "equipment",
        "Técnica de contenção/dispersão": "tecnica_contencao",
        "Mobilização de embarcação de resposta": "mobilizacao_barco",
        "Característica do produto": "caracteristica_produto"
    }
    
    colunas_existentes = [col for col in dicionario_colunas.keys() if col in df.columns]
    df_filtrado = df[colunas_existentes].rename(columns=dicionario_colunas)
    
    if 'empresa' in df_filtrado.columns:
        df_filtrado['empresa'] = df_filtrado['empresa'].fillna('Não Informado').astype(str).str.strip()
    if 'bacia_sedimentar' in df_filtrado.columns:
        df_filtrado['bacia_sedimentar'] = df_filtrado['bacia_sedimentar'].fillna('Não Informada').astype(str).str.strip()
    if 'dias_encerramento' in df_filtrado.columns:
        df_filtrado['dias_encerramento'] = pd.to_numeric(df_filtrado['dias_encerramento'], errors='coerce').fillna(0).astype(int)
    if 'tempo_atendimento' in df_filtrado.columns:
        df_filtrado['tempo_atendimento'] = pd.to_numeric(df_filtrado['tempo_atendimento'], errors='coerce')
        
    return df_filtrado

@st.cache_data(ttl=1800)
def carregar_producao_historica(caminho):
    df_total = pd.read_excel(caminho, sheet_name="Total")
    df_bacias = pd.read_excel(caminho, sheet_name="Bacias")
    return df_total, df_bacias

@st.cache_data(ttl=1800)
def carregar_atendimento_historico(caminho):
    df_tot = pd.read_excel(caminho, sheet_name="Total")
    df_tot.rename(columns={"Mais do que 30 dias": "Mais de 30 dias"}, inplace=True)
    df_b24 = pd.read_excel(caminho, sheet_name="Bacias_2024")
    return df_tot, df_b24

@st.cache_data(ttl=1800)
def carregar_encerramento_historico(caminho):
    return pd.read_excel(caminho, sheet_name="Total")

@st.cache_data(ttl=1800)
def carregar_mapeamento_produtos(caminho):
    return pd.read_excel(caminho, sheet_name="produtos_classe")


# --- VERIFICAÇÃO DE CONFIGURAÇÃO DE ARQUIVOS ---
if os.path.exists(NOME_ACIDENTES) and os.path.exists(NOME_PRODUCAO) and os.path.exists(NOME_ATENDIMENTO) and os.path.exists(NOME_ENCERRAMENTO):
    try:
        # Carga dos bancos de dados
        df_2025_bruto = carregar_dados_2025(NOME_ACIDENTES)
        df_total_prod, df_bacias_prod = carregar_producao_historica(NOME_PRODUCAO)
        df_atend_tot, df_atend_b24 = carregar_atendimento_historico(NOME_ATENDIMENTO)
        df_enc_hist = carregar_encerramento_historico(NOME_ENCERRAMENTO)
        df_map_prod = carregar_mapeamento_produtos(NOME_MAP_PRODUTOS)
        
        # Filtro de Plataformas para 2025
        if 'origem_acidente' in df_2025_bruto.columns:
            df_plataformas_2025 = df_2025_bruto[
                df_2025_bruto['origem_acidente'].astype(str).str.strip().str.lower() == 'plataforma'
            ].copy()
        else:
            df_plataformas_2025 = pd.DataFrame(columns=df_2025_bruto.columns)
            
        acid_total_2025 = len(df_plataformas_2025)
        
        # Cruzamento de Produção 2025
        df_plataformas_2025['bacia_clean'] = df_plataformas_2025['bacia_sedimentar'].astype(str).str.strip().str.lower()
        counts_2025_dict = df_plataformas_2025['bacia_clean'].value_counts().to_dict()
        df_bacias_prod['bacia_clean'] = df_bacias_prod['Bacia Sedimentar'].astype(str).str.strip().str.lower()
        df_bacias_prod['Acid_2025'] = df_bacias_prod['bacia_clean'].map(counts_2025_dict).fillna(0).astype(int)
        
        # --- PROCESSAMENTO LOGÍSTICO: ATENDIMENTO OCORRÊNCIAS 2025 ---
        def classificar_tempo(t):
            if pd.isna(t): return 'Não Atendidos'
            elif t <= 30: return 'Até 30 dias'
            else: return 'Mais de 30 dias'
            
        df_plataformas_2025['cat_atendimento'] = df_plataformas_2025['tempo_atendimento'].apply(classificar_tempo) if 'tempo_atendimento' in df_plataformas_2025.columns else 'Não Atendidos'
        
        t_ate30 = (df_plataformas_2025['cat_atendimento'] == 'Até 30 dias').sum()
        t_mais30 = (df_plataformas_2025['cat_atendimento'] == 'Mais de 30 dias').sum()
        t_nao = (df_plataformas_2025['cat_atendimento'] == 'Não Atendidos').sum()
        t_med = df_plataformas_2025['tempo_atendimento'].mean() if 'tempo_atendimento' in df_plataformas_2025.columns else 0
        
        stats_bacias_2025 = [{'Bacia': 'Total', 'Até 30 dias': t_ate30, 'Mais de 30 dias': t_mais30, 'Não Atendidos': t_nao, 'Tempo Médio': t_med}]
        for bacia, grupo in df_plataformas_2025.groupby('bacia_sedimentar'):
            ate = (grupo['cat_atendimento'] == 'Até 30 dias').sum()
            mais = (grupo['cat_atendimento'] == 'Mais de 30 dias').sum()
            nao = (grupo['cat_atendimento'] == 'Não Atendidos').sum()
            med = grupo['tempo_atendimento'].mean() if 'tempo_atendimento' in grupo.columns else 0
            stats_bacias_2025.append({'Bacia': bacia.strip(), 'Até 30 dias': ate, 'Mais de 30 dias': mais, 'Não Atendidos': nao, 'Tempo Médio': med})
        df_atend_b25 = pd.DataFrame(stats_bacias_2025)
        
        # --- INTERFACE EM ABAS (Variáveis Padronizadas para Evitar NameError) ---
        sidebar_abas = [
            "📊 Painel Operacional (2025)", 
            "📈 Produção (2021-2025)",
            "⏱️ Atendimento a Emergências",
            "🛢️ Consolidação por Produtos"
        ]
        tab_operacional, tab_comparativa, tab_atendimento, tab_produtos = st.tabs(sidebar_abas)
        
        # =========================================================================
        # ABA 1: PAINEL OPERACIONAL DETALHADO (2025)
        # =========================================================================
        with tab_operacional:
            st.subheader("Filtros do Dashboard")
            col_f1, col_f2 = st.columns(2)
            
            with col_f1:
                bacias_unicas = df_plataformas_2025['bacia_sedimentar'].dropna().unique() if 'bacia_sedimentar' in df_plataformas_2025.columns else []
                bacias_disponiveis = sorted([str(x) for x in bacias_unicas])
                bacias_selecionadas = st.multiselect("Filtrar por Bacias Sedimentares:", options=bacias_disponiveis, default=bacias_disponiveis, key="ms_b_aba1")
                
            with col_f2:
                empresas_unicas = df_plataformas_2025['empresa'].dropna().unique() if 'empresa' in df_plataformas_2025.columns else []
                empresas_disponiveis = sorted([str(x) for x in empresas_unicas])
                empresas_selecionadas = st.multiselect("Filtrar por Empresas:", options=empresas_disponiveis, default=empresas_disponiveis, key="ms_e_aba1")
                
            df_filtrado = df_plataformas_2025[df_plataformas_2025['bacia_sedimentar'].isin(bacias_selecionadas) & df_plataformas_2025['empresa'].isin(empresas_selecionadas)].copy()
            
            st.write("---")
            col1, col2, col3 = st.columns(3)
            tot_filt = len(df_filtrado)
            med_enc = df_filtrado['dias_encerramento'].mean() if 'dias_encerramento' in df_filtrado.columns and tot_filt > 0 else 0
            tot_pei = df_filtrado['acionamento_pei'].astype(str).str.strip().str.upper().isin(['SIM', 'S']).sum() if 'acionamento_pei' in df_filtrado.columns else 0
            perc_pei = (tot_pei / tot_filt * 100) if tot_filt > 0 else 0
            
            with col1: st.metric("Total de Acidentes Filtrados", f"{tot_filt}")
            with col2: st.metric("Média de Dias para Encerramento", f"{med_enc:.1f} dias")
            with col3: st.metric("Taxa de Acionamento de PEI", f"{perc_pei:.1f}%", f"{tot_pei} acionamentos")
            
            st.write("---")
            col_g_aba1, col_t_aba1 = st.columns([1.2, 1.8])
            with col_g_aba1:
                if not df_filtrado.empty and 'empresa' in df_filtrado.columns:
                    df_grafico = df_filtrado['empresa'].value_counts().reset_index()
                    df_grafico.columns = ['Empresa', 'Quantidade']
                    fig = px.bar(df_grafico, x='Quantidade', y='Empresa', orientation='h', color='Quantidade', color_continuous_scale='Reds')
                    fig.update_layout(
                        title=dict(text="<b>Volume de Ocorrências por Operadora (2025)</b>", x=0.5, font=dict(size=18, color='#1E4620')),
                        plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                        yaxis={'categoryorder':'total ascending'}, showlegend=False, margin=dict(t=80, b=40, l=40, r=40)
                    )
                    fig.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                    fig.update_yaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                    st.plotly_chart(fig, use_container_width=True)
            with col_t_aba1:
                st.subheader("Base Filtrada (Dados 2025)")
                st.dataframe(df_filtrado[['num_processo', 'instalacao', 'bacia_sedimentar', 'empresa', 'dias_encerramento']], use_container_width=True, height=350)

        # =========================================================================
        # ABA 2: RELATÓRIO COMPARATIVO E PRODUÇÃO (TODOS OS 4 GRÁFICOS RESTAURADOS)
        # =========================================================================
        with tab_comparativa:
            st.markdown("### 📈 Painel Analítico: Histórico e Performance de Incidentes Ambientais")
            st.write("Dados de produção unificados aos registros de incidentes das plataformas.")
            st.write("---")
            
            col_linha1_esq, col_linha1_dir = st.columns(2)
            
            # --- GRÁFICO 1: Histórico de Acidentes vs. Taxa por Produção (2021-2025) ---
            with col_linha1_esq:
                anos_g1 = [2021, 2022, 2023, 2024, 2025]
                acid_vals_g1 = [df_total_prod['Acid_2021'].iloc[0], df_total_prod['Acid_2022'].iloc[0], df_total_prod['Acid_2023'].iloc[0], df_total_prod['Acid_2024'].iloc[0], acid_total_2025]
                prod_vals_g1 = [df_total_prod['Prod_2021'].iloc[0], df_total_prod['Prod_2022'].iloc[0], df_total_prod['Prod_2023'].iloc[0], df_total_prod['Prod_2024'].iloc[0], df_total_prod['Prod_2025'].iloc[0]]
                taxas_g1 = [round(a / p, 1) for a, p in zip(acid_vals_g1, prod_vals_g1)]
                df_g1 = pd.DataFrame({'Ano': [str(x) for x in anos_g1], 'Acidentes': acid_vals_g1, 'Taxa': taxas_g1})
                
                limite_y_comum = max(acid_vals_g1) * 1.25 
                fig1 = make_subplots(specs=[[{"secondary_y": True}]])
                fig1.add_trace(go.Bar(x=df_g1['Ano'], y=df_g1['Acidentes'], name="Nº de Acidentes", marker_color='#3498db', text=df_g1['Acidentes'], textposition='outside', textfont=dict(color='black', size=13)), secondary_y=False)
                fig1.add_trace(go.Scatter(x=df_g1['Ano'], y=df_g1['Taxa'], name="Acidentes / Mboe/d", mode='lines+markers+text', line=dict(color='#2c3e50', width=3), marker=dict(size=8), text=df_g1['Taxa'], textposition='top center', textfont=dict(color='black', size=13)), secondary_y=True)
                
                fig1.update_layout(
                    title=dict(text="<b>Total de Acidentes por Ano e Taxa por Produção (2021-2025)</b>", x=0.5, font=dict(size=18, color='#1E4620')),
                    plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13), legend_title_text='',
                    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), margin=dict(t=100, b=50, l=50, r=50)
                )
                fig1.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                fig1.update_yaxes(title_text="Nº de Acidentes por Ano", secondary_y=False, range=[0, limite_y_comum], showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                fig1.update_yaxes(title_text="Acidentes / Mboe/d", secondary_y=True, range=[0, limite_y_comum], showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                st.plotly_chart(fig1, use_container_width=True)
                
            # --- GRÁFICO 2: Distribuição Anual de Ocorrências por Bacia Sedimentar (2023-2025) ---
            with col_linha1_dir:
                df_g2_clean = df_bacias_prod[(df_bacias_prod['Acid_2023'] > 0) | (df_bacias_prod['Acid_2024'].fillna(0) > 0) | (df_bacias_prod['Acid_2025'] > 0)].copy()
                df_g2_melted = df_g2_clean.melt(id_vars=['Bacia Sedimentar'], value_vars=['Acid_2023', 'Acid_2024', 'Acid_2025'], var_name='Ano', value_name='Acidentes')
                df_g2_melted['Ano'] = df_g2_melted['Ano'].str.replace('Acid_', '')
                
                bacia_ranking = df_g2_clean.set_index('Bacia Sedimentar')[['Acid_2023', 'Acid_2024', 'Acid_2025']].sum(axis=1).sort_values(ascending=False).index.tolist()
                df_g2_melted['Bacia Sedimentar'] = pd.Categorical(df_g2_melted['Bacia Sedimentar'], categories=bacia_ranking, ordered=True)
                df_g2_melted = df_g2_melted.sort_values('Bacia Sedimentar')
                
                fig2 = px.bar(df_g2_melted, x='Bacia Sedimentar', y='Acidentes', color='Ano', barmode='group', text='Acidentes', color_discrete_sequence=['#2ecc71', '#3498db', '#f39c12'], category_orders={"Ano": ["2023", "2024", "2025"]})
                fig2.update_traces(textposition='outside', textfont=dict(color='black', size=12))
                fig2.update_layout(
                    title=dict(text="<b>Distribuição de Ocorrências por Bacia Sedimentar (2023-2025)</b>", x=0.5, font=dict(size=18, color='#1E4620')),
                    xaxis_title="", yaxis_title="Nº de Acidentes", plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                    legend_title_text='', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), margin=dict(t=100, b=50, l=50, r=50)
                )
                fig2.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                fig2.update_yaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                st.plotly_chart(fig2, use_container_width=True)
                
            st.write("---")
            col_linha2_esq, col_linha2_dir = st.columns(2)
            
            # --- GRÁFICO 3: Percentual de Acidentes em Bacias (2025) ---
            with col_linha2_esq:
                df_g3 = df_bacias_prod[df_bacias_prod['Acid_2025'] > 0].copy()
                total_2025_bacias = df_g3['Acid_2025'].sum()
                df_g3['Percentual'] = (df_g3['Acid_2025'] / total_2025_bacias * 100) if total_2025_bacias > 0 else 0
                df_g3 = df_g3.sort_values(by='Percentual', ascending=True)
                
                fig3 = px.bar(df_g3, x='Percentual', y='Bacia Sedimentar', orientation='h', text='Percentual', color_discrete_sequence=['#3498db'])
                fig3.update_traces(texttemplate='%{text:.2f}%', textposition='outside', textfont=dict(color='black', size=12))
                fig3.update_layout(
                    title=dict(text="<b>Percentual de Comunicados por Bacia Sedimentar (2025)</b>", x=0.5, font=dict(size=18, color='#1E4620')),
                    xaxis_title="Percentual de Ocorrências (%)", yaxis_title="", plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                    xaxis=dict(range=[0, 115]), margin=dict(t=100, b=50, l=50, r=50)
                )
                fig3.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                fig3.update_yaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                st.plotly_chart(fig3, use_container_width=True)
                
            # --- GRÁFICO 4: Taxa de Acidentes por Production (Santos e Campos - 2023-2025) ---
            with col_linha2_dir:
                df_g4 = df_bacias_prod[df_bacias_prod['Bacia Sedimentar'].isin(['Campos', 'Santos'])].copy()
                df_g4['Rate_2023'] = df_g4['Acid_2023'] / df_g4['Prod_2023']
                df_g4['Rate_2024'] = df_g4['Acid_2024'] / df_g4['Prod_2024']
                df_g4['Rate_2025'] = df_g4['Acid_2025'] / df_g4['Prod_2025']
                
                df_g4_melted = df_g4.melt(id_vars=['Bacia Sedimentar'], value_vars=['Rate_2023', 'Rate_2024', 'Rate_2025'], var_name='Ano', value_name='Taxa')
                df_g4_melted['Ano'] = df_g4_melted['Ano'].str.replace('Rate_', '')
                
                fig4 = px.bar(df_g4_melted, x='Bacia Sedimentar', y='Taxa', color='Ano', barmode='group', text='Taxa', color_discrete_sequence=['#2ecc71', '#3498db', '#f39c12'], category_orders={"Ano": ["2023", "2024", "2025"]})
                fig4.update_traces(texttemplate='%{text:.1f}', textposition='outside', textfont=dict(color='black', size=12))
                fig4.update_layout(
                    title=dict(text="<b>Taxa de Acidentes por Produção Média Diária (2023-2025)</b>", x=0.5, font=dict(size=18, color='#1E4620')),
                    xaxis_title="", yaxis_title="Acidentes a cada Mboe/d", plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                    legend_title_text='', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), margin=dict(t=100, b=50, l=50, r=50)
                )
                fig4.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                fig4.update_yaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                st.plotly_chart(fig4, use_container_width=True)

        # =========================================================================
        # ABA 3: ATENDIMENTO A EMERGÊNCIAS (TRAMITAÇÃO + ENCERRAMENTOS)
        # =========================================================================
        with tab_atendimento:
            st.markdown("### ⏱️ Eficácia e Tempo de Resposta (Nupaem)")
            st.write("Análise de cumprimento de metas de tramitação e metodologias de resposta do Ibama.")
            st.write("---")
            
            col_atend_1, col_atend_2 = st.columns(2)
            
            # --- FIGURA 3.3.5: Evolução Temporal do Atendimento (2021-2025) ---
            with col_atend_1:
                df_g5 = df_atend_tot[df_atend_tot['Ano'].isin([2021, 2022, 2023, 2024])].copy()
                nova_linha_25 = {'Ano': 2025, 'Até 30 dias': t_ate30, 'Mais de 30 dias': t_mais30, 'Não Atendidos': t_nao, 'Tempo Médio até 1º Atendimento': t_med}
                df_g5 = pd.concat([df_g5, pd.DataFrame([nova_linha_25])], ignore_index=True)
                df_g5['Ano'] = df_g5['Ano'].astype(str)
                
                max_bar_height = (df_g5['Até 30 dias'] + df_g5['Mais de 30 dias'] + df_g5['Não Atendidos']).max()
                max_line_height = df_g5['Tempo Médio até 1º Atendimento'].max()
                limite_y_atend = max(max_bar_height, max_line_height) * 1.15
                
                fig5 = make_subplots(specs=[[{"secondary_y": True}]])
                fig5.add_trace(go.Bar(name='Até 30 dias', x=df_g5['Ano'], y=df_g5['Até 30 dias'], marker_color='#1FA1DD', text=df_g5['Até 30 dias'], textposition='inside', textfont=dict(color='black', size=13)), secondary_y=False)
                fig5.add_trace(go.Bar(name='Mais de 30 dias', x=df_g5['Ano'], y=df_g5['Mais de 30 dias'], marker_color='#FDBB2F', text=df_g5['Mais de 30 dias'], textposition='inside', textfont=dict(color='black', size=13)), secondary_y=False)
                fig5.add_trace(go.Bar(name='Não Atendidos', x=df_g5['Ano'], y=df_g5['Não Atendidos'], marker_color='#8BC53F', text=df_g5['Não Atendidos'], textposition='inside', textfont=dict(color='black', size=13)), secondary_y=False)
                fig5.add_trace(go.Scatter(name='Tempo Médio', x=df_g5['Ano'], y=df_g5['Tempo Médio até 1º Atendimento'], mode='lines+markers+text', line=dict(color='#727272', width=3), marker=dict(size=9, color='#727272'), text=df_g5['Tempo Médio até 1º Atendimento'].round(0), textposition='top center', textfont=dict(color='black', size=13)), secondary_y=True)
                
                fig5.update_layout(
                    barmode='stack', plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                    legend_title_text='', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), margin=dict(t=50, b=50, l=50, r=50)
                )
                fig5.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                fig5.update_yaxes(title_text="Número de Acidentes Atendidos", secondary_y=False, range=[0, limite_y_atend], showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                fig5.update_yaxes(title_text="Tempo Médio (Dias)", secondary_y=True, range=[0, limite_y_atend], showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                st.plotly_chart(fig5, use_container_width=True)
                
            # --- FIGURA 3.3.7: Forma de Atendimento (2025) ---
            with col_atend_2:
                if 'forma_atendimento' in df_plataformas_2025.columns and not df_plataformas_2025['forma_atendimento'].dropna().empty:
                    df_formas = df_plataformas_2025['forma_atendimento'].value_counts().reset_index()
                    df_formas.columns = ['Forma', 'Quantidade']
                    tot_formas = df_formas['Quantidade'].sum()
                    df_formas['Texto'] = df_formas.apply(lambda row: f"({(row['Quantidade']/tot_formas)*100:.1f}%); {row['Quantidade']}", axis=1)
                    
                    fig7 = px.bar(df_formas, x='Forma', y='Quantidade', text='Texto', color_discrete_sequence=['#1FA1DD'])
                    fig7.update_traces(textposition='outside', textfont=dict(color='black', size=12))
                    fig7.update_layout(
                        plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13), margin=dict(t=50, b=120, l=50, r=50)
                    )
                    fig7.update_xaxes(title="", showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                    fig7.update_yaxes(title="Volume de Documentos/Ações", showgrid=False, zeroline=False, linecolor='black', range=[0, df_formas['Quantidade'].max()*1.2], tickfont=dict(size=12))
                    st.plotly_chart(fig7, use_container_width=True)
                else:
                    st.info("Dados de 'Forma de atendimento' insuficientes na planilha de 2025.")
            
            st.write("---")
            
            # --- FIGURA 3.3.6: Fusão Avançada de Eixos Lado a Lado (2024-2025) ---
            ordem_bacias = ["Total", "Campos", "Santos", "Sergipe-Alagoas", "Espírito Santo", "Potiguar", "Ceará", "Camamu-Almada"]
            
            df_b24_c = df_atend_b24.rename(columns={'Tempo Médio até 1º Atendimento': 'Tempo Médio'})
            tot24 = df_g5[df_g5['Ano'] == '2024'].iloc[0]
            linha_tot24 = {'Bacia': 'Total', 'Até 30 dias': tot24['Até 30 dias'], 'Mais de 30 dias': tot24['Mais de 30 dias'], 'Não Atendidos': tot24['Não Atendidos'], 'Tempo Médio': tot24['Tempo Médio até 1º Atendimento']}
            df_b24_c = pd.concat([pd.DataFrame([linha_tot24]), df_b24_c], ignore_index=True)
            df_b24_c['Bacia'] = pd.Categorical(df_b24_c['Bacia'], categories=ordem_bacias, ordered=True)
            df_b24_c = df_b24_c.sort_values('Bacia')
            
            df_b25_c = df_atend_b25.copy()
            df_b25_c['Bacia'] = pd.Categorical(df_b25_c['Bacia'], categories=ordem_bacias, ordered=True)
            df_b25_c = df_b25_c.sort_values('Bacia').dropna(subset=['Bacia'])
            
            fig6 = make_subplots(rows=1, cols=2, subplot_titles=("<b>2024</b>", "<b>2025</b>"), specs=[[{"secondary_y": True}, {"secondary_y": True}]], horizontal_spacing=0.00)
            
            max_b_h = max((df_b24_c['Até 30 dias'] + df_b24_c['Mais de 30 dias'] + df_b24_c['Não Atendidos']).max(), (df_b25_c['Até 30 dias'] + df_b25_c['Mais de 30 dias'] + df_b25_c['Não Atendidos']).max())
            max_l_h = max(df_b24_c['Tempo Médio'].max(), df_b25_c['Tempo Médio'].max())
            limite_y_facets = max(max_b_h, max_l_h) * 1.15
            
            def add_bars_dots(fig, df_data, col_idx):
                fig.add_trace(go.Bar(name='Até 30 dias', x=df_data['Bacia'], y=df_data['Até 30 dias'], marker_color='#1FA1DD', text=df_data['Até 30 dias'], textposition='inside', textfont=dict(color='black', size=11), showlegend=(col_idx==1)), row=1, col=col_idx, secondary_y=False)
                fig.add_trace(go.Bar(name='Mais de 30 dias', x=df_data['Bacia'], y=df_data['Mais de 30 dias'], marker_color='#FDBB2F', text=df_data['Mais de 30 dias'], textposition='inside', textfont=dict(color='black', size=11), showlegend=(col_idx==1)), row=1, col=col_idx, secondary_y=False)
                fig.add_trace(go.Bar(name='Não Atendidos', x=df_data['Bacia'], y=df_data['Não Atendidos'], marker_color='#8BC53F', text=df_data['Não Atendidos'], textposition='inside', textfont=dict(color='black', size=11), showlegend=(col_idx==1)), row=1, col=col_idx, secondary_y=False)
                fig.add_trace(go.Scatter(name='Tempo Médio por Bacia', x=df_data['Bacia'], y=df_data['Tempo Médio'], mode='markers', marker=dict(size=8, color='black'), showlegend=(col_idx==1)), row=1, col=col_idx, secondary_y=True)
            
            add_bars_dots(fig6, df_b24_c, 1)
            add_bars_dots(fig6, df_b25_c, 2)
            
            media_total_2024 = df_b24_c[df_b24_c['Bacia'] == 'Total']['Tempo Médio'].values[0] if not df_b24_c.empty else 0
            media_total_2025 = df_b25_c[df_b25_c['Bacia'] == 'Total']['Tempo Médio'].values[0] if not df_b25_c.empty else 0
            fig6.add_hline(y=media_total_2024, line_dash="dash", line_color="black", row=1, col=1, secondary_y=True)
            fig6.add_hline(y=media_total_2025, line_dash="dash", line_color="black", row=1, col=2, secondary_y=True)
            
            fig6.update_layout(
                barmode='stack', plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                legend_title_text='', legend=dict(orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5), margin=dict(t=80, b=60, l=60, r=60)
            )
            
            fig6.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickangle=45, row=1, col=1, tickfont=dict(size=12))
            fig6.update_yaxes(title_text="Número de Acidentes Atendidos", secondary_y=False, range=[0, limite_y_facets], showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
            fig6.update_yaxes(visible=False, secondary_y=True, row=1, col=1) 
            
            fig6.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickangle=45, row=1, col=2, tickfont=dict(size=12))
            fig6.update_yaxes(visible=False, secondary_y=False, row=1, col=2) 
            fig6.update_yaxes(title_text="Tempo Médio até 1º Atendimento (Dias)", secondary_y=True, range=[0, limite_y_facets], showgrid=False, zeroline=False, linecolor='black', row=1, col=2, tickfont=dict(size=12))
            st.plotly_chart(fig6, use_container_width=True)
            
            # =========================================================================
            # NOVAS VISUALIZAÇÕES: PRAZOS DE ENCERRAMENTO DA INVESTIGAÇÃO (FIG 3.3.8 e 3.3.9)
            # =========================================================================
            st.write("---")
            col_enc_1, col_enc_2 = st.columns(2)
            
            # --- CLASSIFICAÇÃO DOS DADOS DE 2025 ---
            df_plataformas_2025['cat_enc'] = df_plataformas_2025['dias_encerramento'].apply(
                lambda x: 'Investigação em Andamento' if x <= 0 else ('<=180' if x <= 180 else '>180')
            )
            
            # --- FIGURA 3.3.8: Evolução Temporal de Encerramento (2023-2025) ---
            with col_enc_1:
                df_g8 = df_enc_hist[df_enc_hist['Ano'].isin([2023, 2024])].copy()
                
                enc_ate180_25 = (df_plataformas_2025['cat_enc'] == '<=180').sum()
                enc_mais180_25 = (df_plataformas_2025['cat_enc'] == '>180').sum()
                enc_andamento_25 = (df_plataformas_2025['cat_enc'] == 'Investigação em Andamento').sum()
                
                linha_25 = {
                    'Ano': 2025, 
                    '<=180': enc_ate180_25, 
                    '>180': enc_mais180_25, 
                    'Investigação em Andamento': enc_andamento_25
                }
                df_g8 = pd.concat([df_g8, pd.DataFrame([linha_25])], ignore_index=True)
                df_g8['Ano'] = df_g8['Ano'].astype(str)
                
                fig8 = go.Figure()
                fig8.add_trace(go.Bar(name='Até 180 dias (6 meses)', x=df_g8['Ano'], y=df_g8['<=180'], marker_color='#1FA1DD', text=df_g8['<=180'], textposition='inside', textfont=dict(color='black', size=13)))
                fig8.add_trace(go.Bar(name='Mais de 180 dias', x=df_g8['Ano'], y=df_g8['>180'], marker_color='#FDBB2F', text=df_g8['>180'], textposition='inside', textfont=dict(color='black', size=13)))
                fig8.add_trace(go.Bar(name='Em Andamento', x=df_g8['Ano'], y=df_g8['Investigação em Andamento'], marker_color='#8BC53F', text=df_g8['Investigação em Andamento'], textposition='inside', textfont=dict(color='black', size=13)))
                
                fig8.update_layout(
                    barmode='stack', plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                    legend_title_text='', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                    margin=dict(t=50, b=50, l=50, r=50)
                )
                fig8.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                fig8.update_yaxes(title_text="Número de Processos", showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12), title_font=dict(size=14))
                st.plotly_chart(fig8, use_container_width=True)
                
            # --- FIGURA 3.3.9: Encerramento por Bacia Sedimentar (2025) ---
            with col_enc_2:
                bacia_stats_9 = [{
                    'Bacia': 'Total', 
                    '<=180': enc_ate180_25, 
                    '>180': enc_mais180_25, 
                    'Investigação em Andamento': enc_andamento_25
                }]
                
                for bacia, grupo in df_plataformas_2025.groupby('bacia_sedimentar'):
                    bacia_stats_9.append({
                        'Bacia': bacia.strip(),
                        '<=180': (grupo['cat_enc'] == '<=180').sum(),
                        '>180': (grupo['cat_enc'] == '>180').sum(),
                        'Investigação em Andamento': (grupo['cat_enc'] == 'Investigação em Andamento').sum()
                    })
                df_g9 = pd.DataFrame(bacia_stats_9)
                
                ordem_bacias_enc = ["Total", "Campos", "Santos", "Sergipe-Alagoas", "Espírito Santo", "Potiguar", "Ceará", "Camamu-Almada"]
                df_g9['Bacia'] = pd.Categorical(df_g9['Bacia'], categories=ordem_bacias_enc, ordered=True)
                df_g9 = df_g9.sort_values('Bacia').dropna(subset=['Bacia'])
                
                fig9 = go.Figure()
                fig9.add_trace(go.Bar(name='Até 180 dias (6 meses)', x=df_g9['Bacia'], y=df_g9['<=180'], marker_color='#1FA1DD', text=df_g9['<=180'], textposition='inside', textfont=dict(color='black', size=11), showlegend=True))
                fig9.add_trace(go.Bar(name='Mais de 180 dias', x=df_g9['Bacia'], y=df_g9['>180'], marker_color='#FDBB2F', text=df_g9['>180'], textposition='inside', textfont=dict(color='black', size=11), showlegend=True))
                fig9.add_trace(go.Bar(name='Em Andamento', x=df_g9['Bacia'], y=df_g9['Investigação em Andamento'], marker_color='#8BC53F', text=df_g9['Investigação em Andamento'], textposition='inside', textfont=dict(color='black', size=11), showlegend=True))
                
                fig9.update_layout(
                    barmode='stack', plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                    legend_title_text='', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
                    margin=dict(t=50, b=50, l=50, r=50)
                )
                fig9.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickangle=45, tickfont=dict(size=12))
                fig9.update_yaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                st.plotly_chart(fig9, use_container_width=True)

        # =========================================================================
        # ABA 4: CONSOLIDAÇÃO POR PRODUTO (PIPELINE OPTIMIZADO + DOWNLOAD DE-PARA)
        # =========================================================================
        with tab_produtos:
            # --- 1. FUNÇÕES INTERNAS DE HIGIENIZAÇÃO, PADRONIZAÇÃO E FORMATAÇÃO ---
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
                if 'fcba' in n_lower or 'completação aquoso' in n_lower or 'completação base água' in n_lower or 'completação base agua' in n_lower: return "FCBA (Fluido de Completação de Base Aquosa)"
                if 'fpba' in n_lower or ('perfuração' in n_lower and 'base aquosa' in n_lower): return "FPBA (Fluido de Perfuração de Base Aquosa)"
                if 'produto oleoso' in n_lower or n_lower in ['óleo lubrificante', 'oleo lubrificante']: return "Produto Oleoso Genérico"
                return n

            def formatar_volume_br(val):
                s = f"{val:,.8f}"
                parts = s.split('.')
                return f"{parts[0].replace(',', '.')},{parts[1]}"

            def get_cor_risco(classe, fig_num):
                c = str(classe).strip()
                if c == 'A': return '#1FA1DD'
                if c == 'B': return '#8BC53F'
                if c == 'D': return '#FDBB2F'
                if c == 'Não Classificado': return '#8e44ad'
                if c == 'Não Avaliado': return '#F37021' if fig_num in [11, 13] else '#E74C3C'
                return '#BDC3C7'

            # --- 2. EXTRAÇÃO GLOBAL DA BASE UNIFICADA COM PRESERVAÇÃO DO NOME ORIGINAL ---
            registros_brutos = []
            for _, row in df_plataformas_2025.iterrows():
                eq_atual = str(row.get('equipment', 'Não Informado')).strip()
                id_proc = str(row.get('num_processo', 'S/N'))
                
                for p in ['1', '2', '3']:
                    marca_original = str(row.get(f'marca_p{p}')).strip() if pd.notna(row.get(f'marca_p{p}')) else ''
                    if marca_original != '' and marca_original.upper() != 'PREENCHER' and marca_original.lower() != 'nan':
                        vol = limpar_volume_safely(row.get(f'qtd_p{p}'))
                        registros_brutos.append({
                            'Produto_Original': marca_original, # Mapeia o nome original para auditoria
                            'Produto': padronizar_nome_produto(marca_original),
                            'Volume': vol,
                            'Equipamento': eq_atual,
                            'Processo': id_proc
                        })
            
            df_todas_liberacoes = pd.DataFrame(registros_brutos) if registros_brutos else pd.DataFrame(columns=['Produto_Original', 'Produto', 'Volume', 'Equipamento', 'Processo'])
            
            if not df_todas_liberacoes.empty:
                # Enriquecimento com as colunas reais da planilha produtos_consolidados.xlsx
                df_todas_liberacoes = df_todas_liberacoes.merge(df_map_prod, how='left', left_on='Produto', right_on='Nome do Produto')
                df_todas_liberacoes['Classe de Risco'] = df_todas_liberacoes['Classe de Risco'].fillna('Não Avaliado')
                df_todas_liberacoes['Tipo'] = df_todas_liberacoes['Tipo'].fillna('Sem Informação')
                
                # ISOLAMENTO DE LÍQUIDOS NOCIVOS: Remove registros marcados como "Não se Aplica"
                df_liquidos = df_todas_liberacoes[
                    (df_todas_liberacoes['Classe de Risco'] != 'Não se Aplica') & 
                    (df_todas_liberacoes['Tipo'] != 'Não se Aplica')
                ].copy()

                # --- 3. MAPEAMENTO REATIVO DE EQUIPAMENTOS ATÔMICOS ---
                equipamentos_unicos_filtro = set()
                for eq_row in df_liquidos['Equipamento']:
                    eq_str = str(eq_row).strip()
                    if pd.isna(eq_row) or eq_str == '' or eq_str.lower() in ['nan', 'não informado']:
                        equipamentos_unicos_filtro.add('Não Informado')
                    else:
                        for item in eq_str.split(','):
                            item_clean = item.strip()
                            if item_clean: equipamentos_unicos_filtro.add(item_clean)
                list_equip_ordenada = sorted(list(equipamentos_unicos_filtro))
                
                list_classe = sorted(list(df_liquidos['Classe de Risco'].unique()))
                list_tipo = sorted(list(df_liquidos['Tipo'].unique()))

                # --- 4. SEÇÃO DE FILTROS AVANÇADOS DA ABA ---
                st.subheader("Filtros Analíticos de Líquidos Nocivos")
                col_p1, col_p2, col_p3 = st.columns(3)
                
                with col_p1:
                    equip_selecionados = st.multiselect("Filtrar por Equipamento Envolvido:", options=list_equip_ordenada, default=list_equip_ordenada, key="ms_eq_p4")
                with col_p2:
                    classes_selecionadas = st.multiselect("Filtrar por Classe de Risco:", options=list_classe, default=list_classe, key="ms_cl_p4")
                with col_p3:
                    tipos_selecionados = st.multiselect("Filtrar por Tipo de Produto:", options=list_tipo, default=list_tipo, key="ms_tp_p4")
                
                def verificar_aderencia(eq_val, selecionados):
                    eq_str = str(eq_val).strip()
                    if pd.isna(eq_val) or eq_str == '' or eq_str.lower() in ['nan', 'não informado']: return 'Não Informado' in selecionados
                    items = [item.strip() for item in eq_str.split(',') if item.strip()]
                    return any(item in selecionados for item in items)

                # Filtragem direta via Pandas preservando os metadados da planilha
                df_prod_filtrado = df_liquidos[
                    (df_liquidos['Equipamento'].apply(lambda x: verificar_aderencia(x, equip_selecionados))) &
                    (df_liquidos['Classe de Risco'].isin(classes_selecionadas)) &
                    (df_liquidos['Tipo'].isin(tipos_selecionados))
                ].copy()

                if not df_prod_filtrado.empty:
                    # Métricas de texto para a Figura 3.3.10
                    tot_acid_25 = len(df_plataformas_2025)
                    df_plataformas_2025['cont_prods'] = df_plataformas_2025.apply(lambda r: sum([1 for i in ['1','2','3'] if pd.notna(r.get(f'marca_p{i}')) and str(r.get(f'marca_p{i}')).strip().upper() not in ['','PREENCHER', 'NAN']]), axis=1)
                    acid_2_simultaneos = (df_plataformas_2025['cont_prods'] >= 2).sum()
                    
                    df_gas = df_todas_liberacoes[df_todas_liberacoes['Produto'].str.contains('Gás natural', case=False, na=False)]
                    acid_gas = df_gas['Processo'].nunique()
                    vol_gas = df_gas['Volume'].sum()
                    
                    st.write("---")
                    
                    # --- 5. INDICADORES EXECUTIVOS FLUIDOS ---
                    col_m1, col_m2, col_m3 = st.columns(3)
                    with col_m1: st.metric("Total de Produtos Distintos", f"{df_prod_filtrado['Produto'].nunique()}")
                    with col_m2: st.metric("Total de Liberações de Produtos", f"{len(df_prod_filtrado)}")
                    with col_m3: st.metric("Soma do Volume Total Liberado", f"{formatar_volume_br(df_prod_filtrado['Volume'].sum())}")
                    
                    st.write("---")
                    st.markdown(f"***Texto Auxiliar:** Em **{acid_2_simultaneos}** dos **{tot_acid_25}** acidentes ocorridos em 2025 houve a liberação simultânea de dois ou mais produtos distintos. Houve **{acid_gas}** acidentes envolvendo a liberação de gás natural, num volume total de aproximadamente **{vol_gas:,.2f}** m3, que foram excluídos desta análise. Desta forma, o universo amostral nos gráficos e tabela a seguir é de **{len(df_prod_filtrado)}** produtos líquidos liberados em **{df_prod_filtrado['Processo'].nunique()}** processos.*")
                    
                    # --- NOVO REQUISITO: EXPORTADOR DE CORRELAÇÃO DE NOMES (DE-PARA) ---
                    import io
                    buffer_excel = io.BytesIO()
                    with pd.ExcelWriter(buffer_excel, engine='openpyxl') as writer:
                        # Extrai a correlação única do universo mapeado
                        df_correlacao = df_liquidos[['Produto_Original', 'Produto']].drop_duplicates().sort_values('Produto_Original')
                        df_correlacao.columns = ['Nome Original (Planilha)', 'Nome Harmonizado (Painel)']
                        df_correlacao.to_excel(writer, index=False, sheet_name="De-Para_Produtos")
                    buffer_excel.seek(0)
                    
                    st.download_button(
                        label="📥 Baixar Dicionário de Correlação de Nomes (De-Para Excel)",
                        data=buffer_excel,
                        file_name="dicionario_correlacao_produtos.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.write("---")
                    
                    # --- 6. PLOTAGEM DOS 5 GRÁFICOS DINÂMICOS (FIGURAS 3.3.10 a 3.3.14) ---
                    col_graf1, col_graf2 = st.columns(2)
                    
                    # FIGURA 3.3.10: Tipo (Oleosos x Não Oleosos)
                    with col_graf1:
                        df_g10 = df_prod_filtrado.groupby('Tipo').agg(Vol=('Volume','sum'), Acid=('Processo','nunique')).reset_index()
                        fig10 = make_subplots(specs=[[{"secondary_y": True}]])
                        cores_tp = {'Não Oleoso': '#1FA1DD', 'Oleoso': '#8BC53F', 'Sem Informação': '#BDC3C7'}
                        for tipo in df_g10['Tipo'].unique():
                            d = df_g10[df_g10['Tipo'] == tipo]
                            fig10.add_trace(go.Bar(name=f"Volume {tipo}", x=d['Tipo'], y=d['Vol'], marker_color=cores_tp.get(tipo, '#BDC3C7'), text=d['Vol'].apply(lambda x: f"{x:,.1f} m3".replace('.',',')), textposition='inside', showlegend=False), secondary_y=False)
                            fig10.add_trace(go.Scatter(name=f'Acidentes {tipo}', x=d['Tipo'], y=d['Acid'], mode='markers+text', marker=dict(color='black', size=12), text=d['Acid'], textposition='top center', textfont=dict(color='black', size=13), showlegend=False), secondary_y=True)
                        fig10.update_layout(plot_bgcolor='white', margin=dict(t=30, b=30, l=40, r=40))
                        fig10.update_xaxes(showgrid=False, linecolor='black')
                        fig10.update_yaxes(title_text="Volume de Produto Liberado (m3)", secondary_y=False, showgrid=False, linecolor='black')
                        fig10.update_yaxes(title_text="Número de Acidentes", secondary_y=True, showgrid=False, linecolor='black')
                        st.plotly_chart(fig10, use_container_width=True)

                    # FIGURA 3.3.11: Classe de Risco Geral
                    with col_graf2:
                        df_g11 = df_prod_filtrado.groupby('Classe de Risco').agg(Vol=('Volume','sum'), Acid=('Processo','nunique')).reset_index()
                        ordem_11 = ['A', 'B', 'D', 'Não Classificado', 'Não Avaliado']
                        df_g11['Classe de Risco'] = pd.Categorical(df_g11['Classe de Risco'], categories=ordem_11, ordered=True)
                        df_g11 = df_g11.sort_values('Classe de Risco').dropna(subset=['Classe de Risco'])
                        
                        fig11 = make_subplots(specs=[[{"secondary_y": True}]])
                        for _, r in df_g11.iterrows():
                            fig11.add_trace(go.Bar(name=r['Classe de Risco'], x=[r['Classe de Risco']], y=[r['Vol']], marker_color=get_cor_risco(r['Classe de Risco'], 11), text=f"{r['Vol']:,.2f} m3".replace('.',','), textposition='inside', showlegend=False), secondary_y=False)
                            fig11.add_trace(go.Scatter(name='Acidentes', x=[r['Classe de Risco']], y=[r['Acid']], mode='markers+text', marker=dict(color='black', size=10), text=str(r['Acid']), textposition='top center', textfont=dict(color='black', size=13), showlegend=False), secondary_y=True)
                        fig11.update_layout(plot_bgcolor='white', margin=dict(t=30, b=30, l=40, r=40))
                        fig11.update_xaxes(showgrid=False, linecolor='black')
                        fig11.update_yaxes(title_text="Volume Liberado (m3)", secondary_y=False, showgrid=False, linecolor='black')
                        fig11.update_yaxes(title_text="Número de Acidentes", secondary_y=True, showgrid=False, linecolor='black')
                        st.plotly_chart(fig11, use_container_width=True)

                    # FIGURA 3.3.12: Top 20 por Ocorrência (Barras Horizontais) - ORDENAÇÃO CORRIGIDA
                    df_g12 = df_prod_filtrado.groupby(['Produto', 'Classe de Risco']).agg(Acid=('Processo','nunique')).reset_index()
                    df_g12 = df_g12.sort_values(by='Acid', ascending=False).head(20)
                    df_g12['Rank'] = [f"{i}. {p}" for i, p in enumerate(df_g12['Produto'], 1)]
                    
                    fig12 = go.Figure()
                    for c in ordem_11:
                        d = df_g12[df_g12['Classe de Risco'] == c]
                        if not d.empty:
                            fig12.add_trace(go.Bar(name=f'Risco {c}' if c in ['A','B','D'] else c, y=d['Rank'], x=d['Acid'], orientation='h', marker_color=get_cor_risco(c, 12), text=d['Acid'], textposition='outside'))
                    
                    # Inverte a lista de ranking. Como o Plotly desenha de baixo para cima, o index 0 (maior) vai para o topo
                    lista_rank_eixo_y = df_g12['Rank'].tolist()[::-1]
                    
                    fig12.update_layout(barmode='stack', plot_bgcolor='white', margin=dict(t=30, b=30, l=40, r=40), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5))
                    fig12.update_xaxes(showgrid=False, zeroline=False, linecolor='black')
                    fig12.update_yaxes(showgrid=False, zeroline=False, linecolor='black', categoryorder='array', categoryarray=lista_rank_eixo_y)
                    st.plotly_chart(fig12, use_container_width=True)

                    # FIGURA 3.3.13: Acidentes por Faixa de Volume (BINS)
                    bins = [-float('inf'), 0.00001, 0.0001, 0.001, 0.01, 0.1, 1.0, 8.0, 200.0, float('inf')]
                    lbls = ['<= 10 mL', '10 mL < x <= 100 mL', '100 mL < x <= 1 L', '1 L < x <= 10 L', '10 L < x <= 100 L', '100 L < x <= 1 m3', '1 m3 < x <= 8 m3', '8 m3 < x <= 200 m3', 'x > 200 m3']
                    df_prod_filtrado['Faixa_Vol'] = pd.cut(df_prod_filtrado['Volume'], bins=bins, labels=lbls)
                    
                    df_g13 = df_prod_filtrado.groupby(['Faixa_Vol', 'Classe de Risco'], observed=False).agg(Acid=('Processo','nunique')).reset_index()
                    fig13 = go.Figure()
                    for c in ordem_11:
                        d = df_g13[df_g13['Classe de Risco'] == c].set_index('Faixa_Vol').reindex(lbls).reset_index().fillna({'Acid':0})
                        fig13.add_trace(go.Bar(name=f'Risco {c}' if c in ['A','B','D'] else c, x=d['Faixa_Vol'], y=d['Acid'], marker_color=get_cor_risco(c, 13), text=d['Acid'].replace(0, ''), textposition='inside'))
                    
                    df_g13_tot = df_prod_filtrado.groupby('Faixa_Vol', observed=False).agg(Acid=('Processo','nunique')).reset_index()
                    for _, r in df_g13_tot.iterrows():
                        fig13.add_annotation(x=r['Faixa_Vol'], y=r['Acid'], text=f"<b>{r['Acid']}</b>", showarrow=False, yshift=15, font=dict(color="white", size=12), bgcolor="black", bordercolor="black", borderpad=3)
                    
                    fig13.update_layout(barmode='stack', plot_bgcolor='white', margin=dict(t=30, b=30, l=40, r=40), legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5))
                    fig13.update_xaxes(tickangle=45, showgrid=False, linecolor='black')
                    fig13.update_yaxes(showgrid=False, linecolor='black')
                    st.plotly_chart(fig13, use_container_width=True)

                    # FIGURA 3.3.14: Top 20 por Volume + Demais - EXPANDIDA E INCLINADA
                    df_g14 = df_prod_filtrado.groupby(['Produto', 'Classe de Risco']).agg(Vol=('Volume','sum'), Acid=('Processo','nunique')).reset_index().sort_values(by='Vol', ascending=False)
                    top20 = df_g14.head(20).copy()
                    demais = df_g14.iloc[20:].copy()
                    if not demais.empty:
                        top20 = pd.concat([top20, pd.DataFrame([{'Produto':'Demais Produtos', 'Classe de Risco':'Não Avaliado', 'Vol':demais['Vol'].sum(), 'Acid':demais['Acid'].sum()}])], ignore_index=True)
                    
                    top20['Vol_Medio'] = top20['Vol'] / top20['Acid']
                    top20['Rank'] = [f"{i}. {p}" if p != 'Demais Produtos' else p for i, p in enumerate(top20['Produto'], 1)]
                    
                    fig14 = make_subplots(specs=[[{"secondary_y": True}]])
                    for c in ordem_11:
                        d = top20[top20['Classe de Risco'] == c]
                        if not d.empty:
                            fig14.add_trace(go.Bar(name=f'Risco {c}' if c in ['A','B','D'] else c, x=d['Rank'], y=d['Vol'], marker_color=get_cor_risco(c, 14), text=d['Vol'].apply(lambda x: f"{x:,.2f}".replace('.',',')), textposition='outside'), secondary_y=False)
                    
                    fig14.add_trace(go.Scatter(name='Volume Médio', x=top20['Rank'], y=top20['Vol_Medio'], mode='markers+text', marker=dict(color='grey', size=8), text=top20['Vol_Medio'].apply(lambda x: f"{x:,.2f}".replace('.',',')), textposition='top center', textfont=dict(color='grey'), showlegend=False), secondary_y=True)
                    
                    fig14.update_layout(
                        plot_bgcolor='white', 
                        height=650,  # <-- Estica a figura verticalmente eliminando o efeito espremido
                        margin=dict(t=50, b=180, l=40, r=40),  # <-- Amplia o respiro inferior para acomodar os textos longos
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5)
                    )
                    fig14.update_xaxes(tickangle=45, showgrid=False, linecolor='black')  # <-- Inclina os rótulos perfeitamente em 45°
                    fig14.update_yaxes(title_text="Volume Total Liberado (m3)", secondary_y=False, showgrid=False, linecolor='black')
                    fig14.update_yaxes(title_text="Volume Médio", secondary_y=True, showgrid=False, linecolor='black')
                    st.plotly_chart(fig14, use_container_width=True)

                    st.write("---")
                    
                    # --- 7. TABELA CORPORATIVA DE LÍQUIDOS NOCIVOS ---
                    def obter_unicos(series):
                        s = set()
                        for v in series:
                            if pd.notna(v):
                                for item in str(v).split(','):
                                    cl = item.strip()
                                    if cl and cl.lower() not in ['nan', 'não informado']: s.add(cl)
                        return ", ".join(sorted(s)) if s else "Não Informado"

                    df_resumo = df_prod_filtrado.groupby(['Produto', 'Classe de Risco', 'Tipo']).agg(Qtd_Acidentes=('Processo', 'nunique'), Vol_Total=('Volume', 'sum'), Equipamentos_Lista=('Equipamento', obter_unicos)).reset_index()
                    df_resumo.columns = ['Nome do Produto', 'Classe de Risco', 'Tipo', 'Quantidade de Acidentes', 'Soma dos Volumes', 'Equipamentos Envolvidos']
                    df_resumo = df_resumo[['Nome do Produto', 'Quantidade de Acidentes', 'Soma dos Volumes', 'Classe de Risco', 'Tipo', 'Equipamentos Envolvidos']].sort_values(by='Nome do Produto')
                    
                    st.dataframe(df_resumo.style.format({'Soma dos Volumes': '{:,.8f}'}, decimal=',', thousands='.'), use_container_width=True)
                else:
                    st.info("Nenhum produto corresponde aos critérios dos filtros selecionados.")
                
    except Exception as e:
        st.error("Erro interno ao consolidar os dados das abas.")
        st.code(str(e))
else:
    st.warning("⚠️ Arquivos necessários não encontrados!")
    st.markdown(
        f"""
        Para que o painel completo funcione, garanta que todos os arquivos estejam no seu repositório:
        1. **Planilha de Acidentes 2025:** `{NOME_ACIDENTES}` (Aba "Geral")
        2. **Histórico de Produção:** `{NOME_PRODUCAO}` (Abas "Total" e "Bacias")
        3. **Histórico de Tramitação:** `{NOME_ATENDIMENTO}` (Abas "Total" e "Bacias_2024")
        4. **Histórico de Encerramentos:** `{NOME_ENCERRAMENTO}` (Aba "Total")
        """
    )
