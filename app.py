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
        
    return df_filtrado

@st.cache_data(ttl=1800)
def carregar_producao_historica(caminho):
    df_total = pd.read_excel(caminho, sheet_name="Total")
    df_bacias = pd.read_excel(caminho, sheet_name="Bacias")
    return df_total, df_bacias


# --- VERIFICAÇÃO DE ARQUIVOS ---
if os.path.exists(NOME_ACIDENTES) and os.path.exists(NOME_PRODUCAO):
    try:
        df_2025_bruto = carregar_dados_2025(NOME_ACIDENTES)
        df_total_prod, df_bacias_prod = carregar_producao_historica(NOME_PRODUCAO)
        
        if 'origem_acidente' in df_2025_bruto.columns:
            df_plataformas_2025 = df_2025_bruto[
                df_2025_bruto['origem_acidente'].astype(str).str.strip().str.lower() == 'plataforma'
            ].copy()
        else:
            df_plataformas_2025 = pd.DataFrame(columns=df_2025_bruto.columns)
            
        acid_total_2025 = len(df_plataformas_2025)
        
        df_plataformas_2025['bacia_clean'] = df_plataformas_2025['bacia_sedimentar'].astype(str).str.strip().str.lower()
        counts_2025_dict = df_plataformas_2025['bacia_clean'].value_counts().to_dict()
        
        df_bacias_prod['bacia_clean'] = df_bacias_prod['Bacia Sedimentar'].astype(str).str.strip().str.lower()
        df_bacias_prod['Acid_2025'] = df_bacias_prod['bacia_clean'].map(counts_2025_dict).fillna(0).astype(int)
        
        # --- INTERFACE EM ABAS ---
        tab_operacional, tab_comparativa = st.tabs([
            "📊 Painel Operacional (2025)", 
            "📈 Relatório Comparativo & Produção (2021-2025)"
        ])
        
        # ==========================================
        # ABA 1: PAINEL OPERACIONAL DETALHADO (2025)
        # ==========================================
        with tab_operacional:
            st.subheader("Filtros do Dashboard")
            col_f1, col_f2 = st.columns(2)
            
            with col_f1:
                bacias_unicas = df_plataformas_2025['bacia_sedimentar'].dropna().unique() if 'bacia_sedimentar' in df_plataformas_2025.columns else []
                bacias_disponiveis = sorted([str(x) for x in bacias_unicas])
                bacias_selecionadas = st.multiselect(
                    "Filtrar por Bacias Sedimentares:",
                    options=bacias_disponiveis,
                    default=bacias_disponiveis,
                    key="multiselect_bacias_aba1"
                )
                
            with col_f2:
                empresas_unicas = df_plataformas_2025['empresa'].dropna().unique() if 'empresa' in df_plataformas_2025.columns else []
                empresas_disponiveis = sorted([str(x) for x in empresas_unicas])
                empresas_selecionadas = st.multiselect(
                    "Filtrar por Empresas:",
                    options=empresas_disponiveis,
                    default=empresas_disponiveis,
                    key="multiselect_empresas_aba1"
                )
                
            df_filtrado = df_plataformas_2025.copy()
            if 'bacia_sedimentar' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['bacia_sedimentar'].isin(bacias_selecionadas)]
            if 'empresa' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['empresa'].isin(empresas_selecionadas)]
            
            st.write("---")
            
            # Indicadores (KPIs)
            col1, col2, col3 = st.columns(3)
            total_acidentes_filtrados = len(df_filtrado)
            
            if 'dias_encerramento' in df_filtrado.columns and total_acidentes_filtrados > 0:
                media_dias_fechamento = df_filtrado['dias_encerramento'].mean()
            else:
                media_dias_fechamento = 0
                
            if 'acionamento_pei' in df_filtrado.columns and total_acidentes_filtrados > 0:
                total_pei = df_filtrado['acionamento_pei'].astype(str).str.strip().str.upper().isin(['SIM', 'S']).sum()
                percentual_pei = (total_pei / total_acidentes_filtrados * 100)
            else:
                total_pei, percentual_pei = 0, 0
            
            with col1:
                st.metric("Total de Acidentes Filtrados", f"{total_acidentes_filtrados}")
            with col2:
                st.metric("Média de Dias para Encerramento", f"{media_dias_fechamento:.1f} dias")
            with col3:
                st.metric("Taxa de Acionamento de PEI", f"{percentual_pei:.1f}%", f"{total_pei} acionamentos")
                
            st.write("---")
            
            col_g_aba1, col_t_aba1 = st.columns([1.2, 1.8])
            
            with col_g_aba1:
                if not df_filtrado.empty and 'empresa' in df_filtrado.columns:
                    df_grafico = df_filtrado['empresa'].value_counts().reset_index()
                    df_grafico.columns = ['Empresa', 'Quantidade']
                    
                    fig = px.bar(
                        df_grafico, 
                        x='Quantidade', 
                        y='Empresa', 
                        orientation='h',
                        labels={'Quantidade': 'Número de Acidentes', 'Empresa': ''},
                        color='Quantidade',
                        color_continuous_scale='Reds'
                    )
                    
                    fig.update_layout(
                        title=dict(
                            text="<b>Volume de Ocorrências por Operadora (2025)</b>",
                            x=0.5,
                            xanchor='center',
                            font=dict(size=18, color='#1E4620')
                        ),
                        plot_bgcolor='white',
                        paper_bgcolor='white',
                        font=dict(color='black', size=13), 
                        yaxis={'categoryorder':'total ascending'}, 
                        showlegend=False,
                        margin=dict(t=80, b=40, l=40, r=40) 
                    )
                    fig.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                    fig.update_yaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                    fig.update_traces(textfont=dict(color='black', size=12))
                    fig.update_coloraxes(
                        colorbar=dict(
                            tickfont=dict(color='black', size=11),
                            title_font=dict(color='black', size=13)
                        )
                    )
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Filtros muito restritivos para gerar gráfico.")
                    
            with col_t_aba1:
                st.subheader("Base Filtrada (Dados 2025)")
                colunas_exibicao = [c for c in ['num_processo', 'instalacao', 'bacia_sedimentar', 'empresa', 'dias_encerramento'] if c in df_filtrado.columns]
                st.dataframe(df_filtrado[colunas_exibicao], use_container_width=True, height=350)
                
        # ==========================================
        # ABA 2: RELATÓRIO COMPARATIVO E PRODUÇÃO
        # ==========================================
        with tab_comparativa:
            st.markdown("### 📈 Painel Analítico: Histórico e Performance de Incidentes Ambientais")
            st.write("Dados de produção unificados aos registros de incidentes das plataformas.")
            st.write("---")
            
            col_linha1_esq, col_linha1_dir = st.columns(2)
            
            # --- GRÁFICO 1: Histórico de Acidentes vs. Taxa por Produção (2021-2025) ---
            with col_linha1_esq:
                anos_g1 = [2021, 2022, 2023, 2024, 2025]
                acid_vals_g1 = [
                    df_total_prod['Acid_2021'].iloc[0],
                    df_total_prod['Acid_2022'].iloc[0],
                    df_total_prod['Acid_2023'].iloc[0],
                    df_total_prod['Acid_2024'].iloc[0],
                    acid_total_2025
                ]
                prod_vals_g1 = [
                    df_total_prod['Prod_2021'].iloc[0],
                    df_total_prod['Prod_2022'].iloc[0],
                    df_total_prod['Prod_2023'].iloc[0],
                    df_total_prod['Prod_2024'].iloc[0],
                    df_total_prod['Prod_2025'].iloc[0]
                ]
                taxas_g1 = [round(a / p, 1) for a, p in zip(acid_vals_g1, prod_vals_g1)]
                
                df_g1 = pd.DataFrame({'Ano': [str(x) for x in anos_g1], 'Acidentes': acid_vals_g1, 'Taxa': taxas_g1})
                
                limite_y_comum = max(acid_vals_g1) * 1.25 
                
                fig1 = make_subplots(specs=[[{"secondary_y": True}]])
                fig1.add_trace(
                    go.Bar(
                        x=df_g1['Ano'], y=df_g1['Acidentes'], 
                        name="Nº de Acidentes", marker_color='#3498db',
                        text=df_g1['Acidentes'], textposition='outside',
                        textfont=dict(color='black', size=13) 
                    ),
                    secondary_y=False
                )
                fig1.add_trace(
                    go.Scatter(
                        x=df_g1['Ano'], y=df_g1['Taxa'], 
                        name="Acidentes / Mboe/d", mode='lines+markers+text',
                        line=dict(color='#2c3e50', width=3), marker=dict(size=8, symbol='circle'),
                        text=df_g1['Taxa'], textposition='top center',
                        textfont=dict(color='black', size=13) 
                    ),
                    secondary_y=True
                )
                
                fig1.update_layout(
                    title=dict(
                        text="<b>Total de Acidentes por Ano e Taxa por Produção (2021-2025)</b>",
                        x=0.5,
                        xanchor='center',
                        font=dict(size=18, color='#1E4620')
                    ),
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    font=dict(color='black', size=13), 
                    legend_title_text='', # <-- GARANTE QUE NÃO HÁ TÍTULO DE LEGENDA FLUTUANDO
                    legend=dict(
                        orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5,
                        font=dict(color='black', size=13) 
                    ),
                    margin=dict(t=100, b=50, l=50, r=50) 
                )
                fig1.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                fig1.update_yaxes(title_text="Nº de Acidentes por Ano (Barras)", secondary_y=False, range=[0, limite_y_comum], showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                fig1.update_yaxes(title_text="Acidentes / Mboe/d (Linha)", secondary_y=True, range=[0, limite_y_comum], showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                st.plotly_chart(fig1, use_container_width=True)
                
            # --- GRÁFICO 2: Acidentes por Bacia Sedimentar (2023-2025) ---
            with col_linha1_dir:
                df_g2_clean = df_bacias_prod[
                    (df_bacias_prod['Acid_2023'] > 0) | 
                    (df_bacias_prod['Acid_2024'].fillna(0) > 0) | 
                    (df_bacias_prod['Acid_2025'] > 0)
                ].copy()
                
                df_g2_melted = df_g2_clean.melt(
                    id_vars=['Bacia Sedimentar'],
                    value_vars=['Acid_2023', 'Acid_2024', 'Acid_2025'],
                    var_name='Ano', value_name='Acidentes'
                )
                df_g2_melted['Ano'] = df_g2_melted['Ano'].str.replace('Acid_', '')
                
                bacia_ranking = df_g2_clean.set_index('Bacia Sedimentar')[['Acid_2023', 'Acid_2024', 'Acid_2025']].sum(axis=1).sort_values(ascending=False).index.tolist()
                df_g2_melted['Bacia Sedimentar'] = pd.Categorical(df_g2_melted['Bacia Sedimentar'], categories=bacia_ranking, ordered=True)
                df_g2_melted = df_g2_melted.sort_values('Bacia Sedimentar')
                
                fig2 = px.bar(
                    df_g2_melted, x='Bacia Sedimentar', y='Acidentes', color='Ano', barmode='group',
                    text='Acidentes',
                    color_discrete_sequence=['#2ecc71', '#3498db', '#f39c12'], 
                    category_orders={"Ano": ["2023", "2024", "2025"]} 
                )
                fig2.update_traces(textposition='outside', textfont=dict(color='black', size=12))
                fig2.update_layout(
                    title=dict(
                        text="<b>Distribuição de Ocorrências por Bacia Sedimentar (2023-2025)</b>",
                        x=0.5,
                        xanchor='center',
                        font=dict(size=18, color='#1E4620')
                    ),
                    xaxis_title="", yaxis_title="Nº de Acidentes",
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    font=dict(color='black', size=13),
                    legend_title_text='', # <-- CORREÇÃO: REMOVE A PALAVRA "Ano" SOLTA NA LEGENDA
                    legend=dict(
                        orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5,
                        font=dict(color='black', size=13)
                    ),
                    margin=dict(t=100, b=50, l=50, r=50) 
                )
                fig2.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                fig2.update_yaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                st.plotly_chart(fig2, use_container_width=True)
                
            st.write("---")
            col_linha2_esq, col_linha2_dir = st.columns(2)
            
            # --- GRÁFICO 3: Percentual de Acidentes em 2025 ---
            with col_linha2_esq:
                df_g3 = df_bacias_prod[df_bacias_prod['Acid_2025'] > 0].copy()
                total_2025_bacias = df_g3['Acid_2025'].sum()
                
                if total_2025_bacias > 0:
                    df_g3['Percentual'] = (df_g3['Acid_2025'] / total_2025_bacias) * 100
                else:
                    df_g3['Percentual'] = 0
                    
                df_g3 = df_g3.sort_values(by='Percentual', ascending=True)
                
                fig3 = px.bar(
                    df_g3, x='Percentual', y='Bacia Sedimentar', orientation='h',
                    text='Percentual', color_discrete_sequence=['#3498db'] 
                )
                fig3.update_traces(texttemplate='%{text:.2f}%', textposition='outside', textfont=dict(color='black', size=12))
                fig3.update_layout(
                    title=dict(
                        text="<b>Percentual de Comunicados por Bacia Sedimentar (2025)</b>",
                        x=0.5,
                        xanchor='center',
                        font=dict(size=18, color='#1E4620')
                    ),
                    xaxis_title="Percentual de Ocorrências (%)", yaxis_title="",
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    font=dict(color='black', size=13),
                    margin=dict(t=100, b=50, l=50, r=50) 
                )
                fig3.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                fig3.update_yaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                st.plotly_chart(fig3, use_container_width=True)
                
            # --- GRÁFICO 4: Acidentes por Produção por Bacia (Santos e Campos - 2023-2025) ---
            with col_linha2_dir:
                df_g4 = df_bacias_prod[df_bacias_prod['Bacia Sedimentar'].isin(['Campos', 'Santos'])].copy()
                
                df_g4['Rate_2023'] = df_g4['Acid_2023'] / df_g4['Prod_2023']
                df_g4['Rate_2024'] = df_g4['Acid_2024'] / df_g4['Prod_2024']
                df_g4['Rate_2025'] = df_g4['Acid_2025'] / df_g4['Prod_2025']
                
                df_g4_melted = df_g4.melt(
                    id_vars=['Bacia Sedimentar'],
                    value_vars=['Rate_2023', 'Rate_2024', 'Rate_2025'],
                    var_name='Ano', value_name='Taxa'
                )
                df_g4_melted['Ano'] = df_g4_melted['Ano'].str.replace('Rate_', '')
                
                fig4 = px.bar(
                    df_g4_melted, x='Bacia Sedimentar', y='Taxa', color='Ano', barmode='group',
                    text='Taxa', color_discrete_sequence=['#2ecc71', '#3498db', '#f39c12'],
                    category_orders={"Ano": ["2023", "2024", "2025"]} 
                )
                fig4.update_traces(texttemplate='%{text:.1f}', textposition='outside', textfont=dict(color='black', size=12))
                fig4.update_layout(
                    title=dict(
                        text="<b>Taxa de Acidentes por Produção Média Diária (2023-2025)</b>",
                        x=0.5,
                        xanchor='center',
                        font=dict(size=18, color='#1E4620')
                    ),
                    xaxis_title="", yaxis_title="Acidentes a cada Mboe/d",
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    font=dict(color='black', size=13),
                    legend_title_text='', # <-- CORREÇÃO: REMOVE A PALAVRA "Ano" SOLTA NA LEGENDA
                    legend=dict(
                        orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5,
                        font=dict(color='black', size=13)
                    ),
                    margin=dict(t=100, b=50, l=50, r=50) 
                )
                fig4.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                fig4.update_yaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(color='black', size=12), title_font=dict(color='black', size=14))
                st.plotly_chart(fig4, use_container_width=True)
                
    except Exception as e:
        st.error("Erro interno ao processar dados unificados das planilhas.")
        st.code(str(e))
else:
    st.warning("⚠️ Arquivos necessários não encontrados!")
    st.markdown(
        f"""
        Para que este relatório unificado funcione, garanta que ambos os arquivos estejam presentes na mesma pasta:
        * **Planilha de Acidentes 2025:** `{NOME_ACIDENTES}` (Aba "Geral")
        * **Histórico de Produção do Ibama:** `{NOME_PRODUCAO}` (Com as abas "Total" e "Bacias")
        """
    )
