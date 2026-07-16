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
NOME_ATENDIMENTO = "Temo_Atendimento.xlsx"

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
        "Empresa": "empresa",
        "Acionado PEI ou similar?": "acionamento_pei",
        "Tempo de atendimento (em dias)": "tempo_atendimento",
        "Forma de atendimento (primeiro documento/ação)": "forma_atendimento",
        "Dias Até Encerramento da Investigação": "dias_encerramento",
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
def carregar_producao(caminho):
    return pd.read_excel(caminho, sheet_name="Total"), pd.read_excel(caminho, sheet_name="Bacias")

@st.cache_data(ttl=1800)
def carregar_atendimento(caminho):
    df_tot = pd.read_excel(caminho, sheet_name="Total")
    df_tot.rename(columns={"Mais do que 30 dias": "Mais de 30 dias"}, inplace=True)
    df_b24 = pd.read_excel(caminho, sheet_name="Bacias_2024")
    return df_tot, df_b24

# --- VERIFICAÇÃO DE ARQUIVOS ---
if os.path.exists(NOME_ACIDENTES) and os.path.exists(NOME_PRODUCAO) and os.path.exists(NOME_ATENDIMENTO):
    try:
        df_2025_bruto = carregar_dados_2025(NOME_ACIDENTES)
        df_total_prod, df_bacias_prod = carregar_producao(NOME_PRODUCAO)
        df_atend_tot, df_atend_b24 = carregar_atendimento(NOME_ATENDIMENTO)
        
        if 'origem_acidente' in df_2025_bruto.columns:
            df_25 = df_2025_bruto[df_2025_bruto['origem_acidente'].astype(str).str.strip().str.lower() == 'plataforma'].copy()
        else:
            df_25 = pd.DataFrame(columns=df_2025_bruto.columns)
            
        acid_total_2025 = len(df_25)
        
        df_25['bacia_clean'] = df_25['bacia_sedimentar'].astype(str).str.strip().str.lower()
        counts_2025_dict = df_25['bacia_clean'].value_counts().to_dict()
        df_bacias_prod['bacia_clean'] = df_bacias_prod['Bacia Sedimentar'].astype(str).str.strip().str.lower()
        df_bacias_prod['Acid_2025'] = df_bacias_prod['bacia_clean'].map(counts_2025_dict).fillna(0).astype(int)
        
        # --- ENGENHARIA DE DADOS: ATENDIMENTO 2025 ---
        def cat_tempo(t):
            if pd.isna(t): return 'Não Atendidos'
            elif t <= 30: return 'Até 30 dias'
            else: return 'Mais de 30 dias'
            
        df_25['cat_atendimento'] = df_25['tempo_atendimento'].apply(cat_tempo) if 'tempo_atendimento' in df_25.columns else 'Não Atendidos'
        
        t_ate30 = (df_25['cat_atendimento'] == 'Até 30 dias').sum()
        t_mais30 = (df_25['cat_atendimento'] == 'Mais de 30 dias').sum()
        t_nao = (df_25['cat_atendimento'] == 'Não Atendidos').sum()
        t_med = df_25['tempo_atendimento'].mean() if 'tempo_atendimento' in df_25.columns else 0
        
        bacia_stats_2025 = [{'Bacia': 'Total', 'Até 30 dias': t_ate30, 'Mais de 30 dias': t_mais30, 'Não Atendidos': t_nao, 'Tempo Médio': t_med}]
        for bacia, group in df_25.groupby('bacia_sedimentar'):
            ate = (group['cat_atendimento'] == 'Até 30 dias').sum()
            mais = (group['cat_atendimento'] == 'Mais de 30 dias').sum()
            nao = (group['cat_atendimento'] == 'Não Atendidos').sum()
            med = group['tempo_atendimento'].mean()
            bacia_stats_2025.append({'Bacia': bacia.strip(), 'Até 30 dias': ate, 'Mais de 30 dias': mais, 'Não Atendidos': nao, 'Tempo Médio': med})
        df_atend_b25 = pd.DataFrame(bacia_stats_2025)
        
        # --- INTERFACE EM ABAS ---
        tab_op, tab_comp, tab_atend = st.tabs([
            "📊 Painel Operacional (2025)", 
            "📈 Produção (2021-2025)",
            "⏱️ Atendimento a Emergências"
        ])
        
        # ==========================================
        # ABA 1: PAINEL OPERACIONAL DETALHADO (2025)
        # ==========================================
        with tab_op:
            st.subheader("Filtros do Dashboard")
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                bacias_unicas = df_25['bacia_sedimentar'].dropna().unique() if 'bacia_sedimentar' in df_25.columns else []
                bacias_disp = sorted([str(x) for x in bacias_unicas])
                b_selecionadas = st.multiselect("Filtrar por Bacias:", options=bacias_disp, default=bacias_disp, key="ms_b1")
            with col_f2:
                emp_unicas = df_25['empresa'].dropna().unique() if 'empresa' in df_25.columns else []
                emp_disp = sorted([str(x) for x in emp_unicas])
                e_selecionadas = st.multiselect("Filtrar por Empresas:", options=emp_disp, default=emp_disp, key="ms_e1")
                
            df_filtrado = df_25[df_25['bacia_sedimentar'].isin(b_selecionadas) & df_25['empresa'].isin(e_selecionadas)].copy()
            
            st.write("---")
            col1, col2, col3 = st.columns(3)
            tot_filtrados = len(df_filtrado)
            med_encerra = df_filtrado['dias_encerramento'].mean() if 'dias_encerramento' in df_filtrado.columns and tot_filtrados > 0 else 0
            tot_pei = df_filtrado['acionamento_pei'].astype(str).str.strip().str.upper().isin(['SIM', 'S']).sum() if 'acionamento_pei' in df_filtrado.columns else 0
            perc_pei = (tot_pei / tot_filtrados * 100) if tot_filtrados > 0 else 0
            
            with col1: st.metric("Total Filtrados", f"{tot_filtrados}")
            with col2: st.metric("Média Dias p/ Encerramento", f"{med_encerra:.1f} dias")
            with col3: st.metric("Taxa Acionamento PEI", f"{perc_pei:.1f}%", f"{tot_pei} acionamentos")
            
            st.write("---")
            col_g_aba1, col_t_aba1 = st.columns([1.2, 1.8])
            with col_g_aba1:
                if not df_filtrado.empty and 'empresa' in df_filtrado.columns:
                    df_graf = df_filtrado['empresa'].value_counts().reset_index()
                    df_graf.columns = ['Empresa', 'Quantidade']
                    fig = px.bar(df_graf, x='Quantidade', y='Empresa', orientation='h', color='Quantidade', color_continuous_scale='Reds')
                    fig.update_layout(
                        title=dict(text="<b>Ocorrências por Operadora</b>", x=0.5, font=dict(size=18, color='#1E4620')),
                        plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                        yaxis={'categoryorder':'total ascending'}, showlegend=False, margin=dict(t=80, b=40, l=40, r=40)
                    )
                    fig.update_xaxes(showgrid=False, zeroline=False, linecolor='black')
                    fig.update_yaxes(showgrid=False, zeroline=False, linecolor='black')
                    fig.update_traces(textfont=dict(color='black', size=12))
                    st.plotly_chart(fig, use_container_width=True)
            with col_t_aba1:
                st.subheader("Base Filtrada")
                st.dataframe(df_filtrado[['num_processo', 'instalacao', 'bacia_sedimentar', 'empresa', 'dias_encerramento']], use_container_width=True, height=350)
                
        # ==========================================
        # ABA 2: RELATÓRIO COMPARATIVO E PRODUÇÃO
        # ==========================================
        with tab_comp:
            st.markdown("### 📈 Histórico e Performance (Produção x Acidentes)")
            st.write("---")
            
            anos_g1 = [2021, 2022, 2023, 2024, 2025]
            acid_vals_g1 = [df_total_prod[f'Acid_{a}'].iloc[0] for a in anos_g1[:-1]] + [acid_total_2025]
            prod_vals_g1 = [df_total_prod[f'Prod_{a}'].iloc[0] for a in anos_g1]
            taxas_g1 = [round(a / p, 1) for a, p in zip(acid_vals_g1, prod_vals_g1)]
            df_g1 = pd.DataFrame({'Ano': [str(x) for x in anos_g1], 'Acidentes': acid_vals_g1, 'Taxa': taxas_g1})
            
            limite_y_comum = max(acid_vals_g1) * 1.25 
            fig1 = make_subplots(specs=[[{"secondary_y": True}]])
            fig1.add_trace(go.Bar(x=df_g1['Ano'], y=df_g1['Acidentes'], name="Nº Acidentes", marker_color='#3498db', text=df_g1['Acidentes'], textposition='outside', textfont=dict(color='black', size=13)), secondary_y=False)
            fig1.add_trace(go.Scatter(x=df_g1['Ano'], y=df_g1['Taxa'], name="Taxa (Mboe/d)", mode='lines+markers+text', line=dict(color='#2c3e50', width=3), marker=dict(size=8), text=df_g1['Taxa'], textposition='top center', textfont=dict(color='black', size=13)), secondary_y=True)
            
            fig1.update_layout(
                title=dict(text="<b>Acidentes por Ano e Taxa (2021-2025)</b>", x=0.5, font=dict(size=18, color='#1E4620')),
                plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                legend_title_text='', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5, font=dict(color='black', size=13)),
                margin=dict(t=100, b=50, l=50, r=50)
            )
            fig1.update_xaxes(showgrid=False, zeroline=False, linecolor='black')
            fig1.update_yaxes(title_text="Nº Acidentes", secondary_y=False, range=[0, limite_y_comum], showgrid=False, zeroline=False, linecolor='black')
            fig1.update_yaxes(title_text="Taxa", secondary_y=True, range=[0, limite_y_comum], showgrid=False, zeroline=False, linecolor='black')
            st.plotly_chart(fig1, use_container_width=True)

        # ==========================================
        # ABA 3: ATENDIMENTO A EMERGÊNCIAS
        # ==========================================
        with tab_atend:
            st.markdown("### ⏱️ Eficácia e Tempo de Resposta (Nupaem)")
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
                
                # TITULO PRINCIPAL REMOVIDO CONFORME SOLICITADO
                fig5.update_layout(
                    barmode='stack', plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                    legend_title_text='', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5), margin=dict(t=50, b=50, l=50, r=50)
                )
                fig5.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                fig5.update_yaxes(title_text="Número de Acidentes Atendidos", secondary_y=False, range=[0, limite_y_atend], showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12), title_font=dict(size=14))
                fig5.update_yaxes(title_text="Tempo Médio (Dias)", secondary_y=True, range=[0, limite_y_atend], showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12), title_font=dict(size=14))
                st.plotly_chart(fig5, use_container_width=True)
                
            # --- FIGURA 3.3.7: Forma de Atendimento (2025) ---
            with col_atend_2:
                if 'forma_atendimento' in df_25.columns and not df_25['forma_atendimento'].dropna().empty:
                    df_formas = df_25['forma_atendimento'].value_counts().reset_index()
                    df_formas.columns = ['Forma', 'Quantidade']
                    tot_formas = df_formas['Quantidade'].sum()
                    df_formas['Texto'] = df_formas.apply(lambda row: f"({(row['Quantidade']/tot_formas)*100:.1f}%); {row['Quantidade']}", axis=1)
                    
                    fig7 = px.bar(df_formas, x='Forma', y='Quantidade', text='Texto', color_discrete_sequence=['#1FA1DD'])
                    fig7.update_traces(textposition='outside', textfont=dict(color='black', size=12))
                    
                    # TITULO PRINCIPAL REMOVIDO CONFORME SOLICITADO
                    fig7.update_layout(
                        plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13), margin=dict(t=50, b=120, l=50, r=50)
                    )
                    fig7.update_xaxes(title="", showgrid=False, zeroline=False, linecolor='black', tickfont=dict(size=12))
                    fig7.update_yaxes(title="Volume de Documentos/Ações", showgrid=False, zeroline=False, linecolor='black', range=[0, df_formas['Quantidade'].max()*1.2], tickfont=dict(size=12), title_font=dict(size=14))
                    st.plotly_chart(fig7, use_container_width=True)
                else:
                    st.info("Dados sobre 'Forma de atendimento' insuficientes na planilha de 2025.")
            
            st.write("---")
            
            # --- FIGURA 3.3.6: Atendimento por Bacias (FUSÃO DOS ANOS LADO A LADO) ---
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
            
            # CORREÇÃO CRÍTICA DE DESIGN: horizontal_spacing=0.0 "gruda" fisicamente as duas colunas
            fig6 = make_subplots(
                rows=1, cols=2, 
                subplot_titles=("<b>2024</b>", "<b>2025</b>"), 
                specs=[[{"secondary_y": True}, {"secondary_y": True}]], 
                horizontal_spacing=0.00
            )
            
            max_b_h = max(
                (df_b24_c['Até 30 dias'] + df_b24_c['Mais de 30 dias'] + df_b24_c['Não Atendidos']).max(),
                (df_b25_c['Até 30 dias'] + df_b25_c['Mais de 30 dias'] + df_b25_c['Não Atendidos']).max()
            )
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
            
            # LAYOUT GERAL DA FIGURA 6 (TITULO REMOVIDO CONFORME PEDIDO)
            fig6.update_layout(
                barmode='stack', plot_bgcolor='white', paper_bgcolor='white', font=dict(color='black', size=13),
                legend_title_text='', legend=dict(orientation="h", yanchor="bottom", y=1.08, xanchor="center", x=0.5, font=dict(size=13)), 
                margin=dict(t=80, b=60, l=60, r=60)
            )
            
            # --- SOLUÇÃO CIRÚRGICA DE EIXOS PARA GRUDAR ---
            # COLUNA 1 (2024): Mantém o Eixo Principal (Esquerdo), REMOVE o Eixo Secundário (Direito)
            fig6.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickangle=45, row=1, col=1, tickfont=dict(size=12))
            fig6.update_yaxes(title_text="Número de Acidentes Atendidos", secondary_y=False, range=[0, limite_y_facets], showgrid=False, zeroline=False, linecolor='black', row=1, col=1, tickfont=dict(size=12), title_font=dict(size=14))
            fig6.update_yaxes(visible=False, secondary_y=True, row=1, col=1) # Oculta o centro-esquerda
            
            # COLUNA 2 (2025): REMOVE o Eixo Principal (Esquerdo), Mantém o Eixo Secundário (Direito)
            fig6.update_xaxes(showgrid=False, zeroline=False, linecolor='black', tickangle=45, row=1, col=2, tickfont=dict(size=12))
            fig6.update_yaxes(visible=False, secondary_y=False, row=1, col=2) # Oculta o centro-direita
            fig6.update_yaxes(title_text="Tempo Médio até 1º Atendimento (Dias)", secondary_y=True, range=[0, limite_y_facets], showgrid=False, zeroline=False, linecolor='black', row=1, col=2, tickfont=dict(size=12), title_font=dict(size=14))

            st.plotly_chart(fig6, use_container_width=True)

    except Exception as e:
        st.error("Erro interno ao processar dados de Atendimento.")
        st.code(str(e))
else:
    st.warning("⚠️ Arquivos necessários não encontrados!")
    st.markdown(
        f"Certifique-se de que `{NOME_ACIDENTES}`, `{NOME_PRODUCAO}` e `{NOME_ATENDIMENTO}` estão na mesma pasta."
    )
