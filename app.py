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
    # Carrega a aba inteira para evitar erros de índice de coluna
    df = pd.read_excel(caminho, sheet_name="Geral")
    
    # Remove espaços em branco extras dos nomes das colunas
    df.columns = df.columns.str.strip()
    
    # Mapeamento exato baseado na lista oficial fornecida
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
        "Equipamento/sistema envolvido": "equipamento",
        "Técnica de contenção/dispersão": "tecnica_contencao",
        "Mobilização de embarcação de resposta": "mobilizacao_barco",
        "Característica do produto": "caracteristica_produto"
    }
    
    # Filtra apenas as colunas mapeadas que de fato existem no Excel
    colunas_existentes = [col for col in dicionario_colunas.keys() if col in df.columns]
    df_filtrado = df[colunas_existentes].rename(columns=dicionario_colunas)
    
    # --- TRATAMENTO DE VALORES NULOS E ADAPTAÇÃO DE TIPOS ---
    if 'empresa' in df_filtrado.columns:
        df_filtrado['empresa'] = df_filtrado['empresa'].fillna('Não Informado').astype(str).str.strip()
        df_filtrado['empresa'] = df_filtrado['empresa'].replace({'nan': 'Não Informado', 'NaN': 'Não Informado'})
        
    if 'bacia_sedimentar' in df_filtrado.columns:
        df_filtrado['bacia_sedimentar'] = df_filtrado['bacia_sedimentar'].fillna('Não Informada').astype(str).str.strip()
        df_filtrado['bacia_sedimentar'] = df_filtrado['bacia_sedimentar'].replace({'nan': 'Não Informada', 'NaN': 'Não Informada'})
        
    if 'dias_encerramento' in df_filtrado.columns:
        df_filtrado['dias_encerramento'] = pd.to_numeric(df_filtrado['dias_encerramento'], errors='coerce').fillna(0).astype(int)
    
    return df_filtrado

if os.path.exists(NOME_ARQUIVO):
    try:
        df_bruto = carregar_e_limpar_dados(NOME_ARQUIVO)
        
        # Filtro de Plataformas (busca inteligente na coluna de Origem ou Instalação)
        termos_plataforma = "Plataforma|FPSO|Sonda|FSO|Semi-submersível"
        
        coluna_origem = df_bruto['origem_acidente'].astype(str) if 'origem_acidente' in df_bruto.columns else pd.Series(dtype=str)
        coluna_instalacao = df_bruto['instalacao'].astype(str) if 'instalacao' in df_bruto.columns else pd.Series(dtype=str)
        
        df_plataformas = df_bruto[
            coluna_origem.str.contains(termos_plataforma, case=False, na=False) |
            coluna_instalacao.str.contains(termos_plataforma, case=False, na=False)
        ].copy()
        
        # --- PAINEL LATERAL (Filtros) ---
        st.sidebar.header("Filtros do Painel")
        
        # Correção do erro '<' entre str e float: forçamos a conversão para string de todos os itens e removemos nulos antes do sorted
        if 'bacia_sedimentar' in df_plataformas.columns:
            bacias_unicas = df_plataformas['bacia_sedimentar'].dropna().unique()
            bacias_disponiveis = sorted([str(x) for x in bacias_unicas])
        else:
            bacias_disponiveis = []
            
        bacias_selecionadas = st.sidebar.multiselect(
            "Selecione as Bacias Sedimentares:",
            options=bacias_disponiveis,
            default=bacias_disponiveis
        )
        
        if 'empresa' in df_plataformas.columns:
            empresas_unicas = df_plataformas['empresa'].dropna().unique()
            empresas_disponiveis = sorted([str(x) for x in empresas_unicas])
        else:
            empresas_disponiveis = []
            
        empresas_selecionadas = st.sidebar.multiselect(
            "Selecione as Empresas:",
            options=empresas_disponiveis,
            default=empresas_disponiveis
        )
        
        # Filtra os dados com base na seleção do usuário
        df_filtrado = df_plataformas.copy()
        if 'bacia_sedimentar' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['bacia_sedimentar'].isin(bacias_selecionadas)]
        if 'empresa' in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado['empresa'].isin(empresas_selecionadas)]
        
        # --- CORPO PRINCIPAL (Visualizações) ---
        col1, col2, col3 = st.columns(3)
        
        total_acidentes = len(df_filtrado)
        
        if 'dias_encerramento' in df_filtrado.columns and total_acidentes > 0:
            media_dias_fechamento = df_filtrado['dias_encerramento'].mean()
        else:
            media_dias_fechamento = 0
            
        if 'acionamento_pei' in df_filtrado.columns and total_acidentes > 0:
            total_pei = df_filtrado['acionamento_pei'].astype(str).str.strip().str.upper().isin(['SIM', 'S']).sum()
            percentual_pei = (total_pei / total_acidentes * 100)
        else:
            total_pei, percentual_pei = 0, 0
        
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
                fig.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Sem dados suficientes para gerar o gráfico.")
                
        with col_tabela:
            st.subheader("Base Filtrada")
            colunas_exibicao = [c for c in ['num_processo', 'instalacao', 'bacia_sedimentar', 'empresa', 'dias_encerramento'] if c in df_filtrado.columns]
            st.dataframe(
                df_filtrado[colunas_exibicao], 
                use_container_width=True,
                height=400
            )
            
    except Exception as e:
        st.error("Erro interno no processamento de dados do dashboard.")
        st.code(str(e))
else:
    st.info(f"Aguardando o upload do arquivo `{NOME_ARQUIVO}` para o repositório.")
