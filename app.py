import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Configuração da página
st.set_page_config(
    page_title="Painel de Acidentes - Plataformas",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("⚓ Consolidação de Acidentes em Plataformas de Petróleo")
st.markdown("Análise interativa e monitoramento de emergências ambientais baseadas em dados consolidados.")

NOME_ARQUIVO = "acidentes_2025.xlsx"

@st.cache_data(ttl=1800)  # Mantém o cache por 30 minutos para performance
def carregar_e_limpar_dados(caminho):
    colunas_excel = "B,E,I,O,R:U,X:AA,AD:AG,AK:AM,BF,BI,BM,BO,BZ,CB:CE,CF"
    
    # Leitura direcionada explicitamente para a aba "Geral"
    df = pd.read_excel(
        caminho, 
        sheet_name="Geral", 
        usecols=colunas_excel
    )
    
    dicionario_colunas = {
        'Número do Processo': 'num_processo',
        'Instalação': 'instalacao',
        'Bacia Sedimentar': 'bacia_sedimentar',
        'Origem do Acidente': 'origem_acidente',
        'Empresa': 'empresa',
        'Acionamento do PEI': 'acionamento_pei',
        'Tempo de Atendimento': 'tempo_atendimento',
        'Forma de Atendimento': 'forma_atendimento',
        'Dias até o encerramento': 'dias_encerramento',
        'Caracteristica do Produto': 'caracteristica_produto',
        'Marca Comercial do Produto 1': 'marca_p1',
        'Produto 1': 'prod_1',
        'Quantidade do Produto 1': 'qtd_p1',
        'Unidade': 'uni_p1',
        'Marca Comercial Produto 2': 'marca_p2',
        'Produto 2': 'prod_2',
        'Quantidade Produt 2': 'qtd_p2',
        'Unidade Produt 2': 'uni_p2',
        'Marca Comercial Produto 3': 'marca_p3',
        'Produto 3': 'prod_3',
        'Quantidade Produt 3': 'qtd_p3',
        'Unidade Produt 3': 'uni_p3',
        'Causas 1': 'causa_1',
        'Causas 2': 'causa_2',
        'Causas 3': 'causa_3',
        'OperaçãoOperação/ocorrência/sistema': 'operacao',
        'Equipamento/sistema envolvido': 'equipamento',
        'Técnica de contenção/dispersão': 'tecnica_contencao',
        'Mobilização de embarcação de resposta': 'mobilizacao_barco'
    }
    
    df.columns = [dicionario_colunas.get(col, col) for col in df.columns]
    
    # --- Tratamento de Dados (Data Cleaning) ---
    df['empresa'] = df['empresa'].astype(str).str.strip().replace('nan', 'Não Informado')
    df['bacia_sedimentar'] = df['bacia_sedimentar'].astype(str).str.strip().replace('nan', 'Não Informada')
    df['dias_encerramento'] = pd.to_numeric(df['dias_encerramento'], errors='coerce').fillna(0).astype(int)
    
    return df

if os.path.exists(NOME_ARQUIVO):
    try:
        df_bruto = carregar_e_limpar_dados(NOME_ARQUIVO)
        
        # Filtro de Plataformas (busca inteligente na coluna de Origem ou Instalação)
        termos_plataforma = "Plataforma|FPSO|Sonda|FSO|Semi-submersível"
        df_plataformas = df_bruto[
            df_bruto['origem_acidente'].astype(str).str.contains(termos_plataforma, case=False, na=False) |
            df_bruto['instalacao'].astype(str).str.contains(termos_plataforma, case=False, na=False)
        ].copy()
        
        # --- PAINEL LATERAL (Filtros) ---
        st.sidebar.header("Filtros do Painel")
        
        bacias_disponiveis = sorted(df_plataformas['bacia_sedimentar'].unique())
        bacias_selecionadas = st.sidebar.multiselect(
            "Selecione as Bacias Sedimentares:",
            options=bacias_disponiveis,
            default=bacias_disponiveis
        )
        
        empresas_disponiveis = sorted(df_plataformas['empresa'].unique())
        empresas_selecionadas = st.sidebar.multiselect(
            "Selecione as Empresas:",
            options=empresas_disponiveis,
            default=empresas_disponiveis
        )
        
        df_filtrado = df_plataformas[
            (df_plataformas['bacia_sedimentar'].isin(bacias_selecionadas)) &
            (df_plataformas['empresa'].isin(empresas_selecionadas))
        ]
        
        # --- CORPO PRINCIPAL (Visualizações) ---
        col1, col2, col3 = st.columns(3)
        
        total_acidentes = len(df_filtrado)
        media_dias_fechamento = df_filtrado['dias_encerramento'].mean() if total_acidentes > 0 else 0
        
        total_pei = df_filtrado['acionamento_pei'].astype(str).str.strip().str.upper().isin(['SIM', 'S']).sum()
        percentual_pei = (total_pei / total_acidentes * 100) if total_acidentes > 0 else 0
        
        with col1:
            st.metric("Total de Acidentes", f"{total_acidentes}")
        with col2:
            st.metric("Média de Dias para Encerramento", f"{media_dias_fechamento:.1f} dias")
        with col3:
            st.metric("Taxa de Acionamento de PEI", f"{percentual_pei:.1f}%", f"{total_pei} acionamentos")
            
        st.write("---")
        
        col_grafico, col_tabela = st.columns([1.2, 1.8])
        
        with col_grafico:
            st.subheader("Acidentes por Empresa")
            if not df_filtrado.empty:
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
                fig.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Sem dados suficientes para gerar o gráfico.")
                
        with col_tabela:
            st.subheader("Base Filtrada")
            st.dataframe(
                df_filtrado[['num_processo', 'instalacao', 'bacia_sedimentar', 'empresa', 'dias_encerramento']], 
                use_container_width=True,
                height=400
            )
            
    except Exception as e:
        st.error("Erro ao processar as colunas da planilha.")
        st.code(str(e))
else:
    st.info(f"Aguardando o upload do arquivo `{NOME_ARQUIVO}` para o repositório.")
