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


# --- VERIFICAÇÃO DE CONFIGURAÇÃO DE ARQUIVOS ---
if os.path.exists(NOME_ACIDENTES) and os.path.exists(NOME_PRODUCAO) and os.path.exists(NOME_ATENDIMENTO) and os.path.exists(NOME_ENCERRAMENTO):
    try:
        # Carga dos bancos de dados
        df_2025_bruto = carregar_dados_2025(NOME_ACIDENTES)
        df_total_prod, df_bacias_prod = carregar_producao_historica(NOME_PRODUCAO)
        df_atend_tot, df_atend_b24 = carregar_atendimento_historico(NOME_ATENDIMENTO)
        df_enc_hist = carregar_encerramento_historico(NOME_ENCERRAMENTO)
        
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
        # ABA 4: CONSOLIDAÇÃO DE DADOS POR PRODUTO (MARCAS 1, 2 E 3)
        # =========================================================================
        with tab_produtos:
            st.markdown("### 🛢️ Inventário Consolidado de Produtos e Marcas Comerciais")
            st.write("Agregação unificada de todos os produtos com vazamento registrado em plataformas (2025).")
            st.write("---")
            
            registros_produtos = []
            
            for _, row in df_plataformas_2025.iterrows():
                equipamento_atual = str(row.get('equipment', 'Não Informado')).strip()
                id_processo = str(row.get('num_processo', 'S/N'))
                
                if pd.notna(row.get('marca_p1')) and str(row.get('marca_p1')).strip().lower() != 'nan':
                    marca = str(row['marca_p1']).strip()
                    vol = pd.to_numeric(row.get('qtd_p1'), errors='coerce')
                    registros_produtos.append({'Produto': marca, 'Volume': vol, 'Equipamento': equipamento_atual, 'Processo': id_processo})
                    
                if pd.notna(row.get('marca_p2')) and str(row.get('marca_p2')).strip().lower() != 'nan':
                    marca = str(row['marca_p2']).strip()
                    vol = pd.to_numeric(row.get('qtd_p2'), errors='coerce')
                    registros_produtos.append({'Produto': marca, 'Volume': vol, 'Equipamento': equipamento_atual, 'Processo': id_processo})
                    
                if pd.notna(row.get('marca_p3')) and str(row.get('marca_p3')).strip().lower() != 'nan':
                    marca = str(row['marca_p3']).strip()
                    vol = pd.to_numeric(row.get('qtd_p3'), errors='coerce')
                    registros_produtos.append({'Produto': marca, 'Volume': vol, 'Equipamento': equipamento_atual, 'Processo': id_processo})
            
            if registros_produtos:
                df_prod_bruto = pd.DataFrame(registros_produtos)
                
                df_prod_summary = df_prod_bruto.groupby('Produto').agg(
                    Qtd_Acidentes=('Processo', 'nunique'),
                    Vol_Total=('Volume', 'sum'),
                    Equipamentos_Lista=('Equipamento', lambda x: ", ".join(sorted(set(x) - {'nan', 'Não Informado', ''})))
                ).reset_index()
                
                df_prod_summary['Classe de Risco'] = "A obter"
                df_prod_summary['Tipo'] = "A obter"
                
                df_prod_summary = df_prod_summary[[
                    'Produto', 'Qtd_Acidentes', 'Vol_Total', 'Classe de Risco', 'Tipo', 'Equipamentos_Lista'
                ]]
                
                df_prod_summary.columns = [
                    'Nome do Produto', 'Quantidade de Acidentes', 'Soma dos Volumes', 'Classe de Risco', 'Tipo', 'Equipamentos Envolvidos'
                ]
                
                df_prod_transposto = df_prod_summary.set_index('Nome do Produto').T
                st.dataframe(df_prod_transposto, use_container_width=True)
                
            else:
                st.info("Nenhum registro de vazamento de produto comercial detalhado foi encontrado para os filtros ativos.")
                
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
