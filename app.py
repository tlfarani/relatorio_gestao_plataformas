import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Painel de Acidentes - Plataformas",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("⚓ Consolidação de Acidentes em Plataformas de Petróleo")
st.markdown("Análise dinâmica de dados de emergências ambientais com foco em plataformas.")

# URL do SharePoint adaptada para download direto pelo Pandas
URL_SHAREPOINT = "https://ibamagovbr.sharepoint.com/:x:/s/EmergnciasAmbientais/IQBpuRol-DqdSZHcUIzACVUVAfca-KOFKSrLDeT6lLgEJUc?download=1"

@st.cache_data(ttl=3600)  # Atualiza o cache a cada 1 hora para buscar novos dados do SharePoint
def carregar_dados_nuvem(url):
    colunas_excel = "B,E,I,O,R:U,X:AA,AD:AG,AK:AM,BF,BI,BM,BO,BZ,CB:CE,CF"
    
    # O Pandas consegue ler diretamente de uma URL HTTP
    df = pd.read_excel(
        url, 
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
    return df

try:
    with st.spinner("Conectando ao SharePoint e carregando dados..."):
        df_bruto = carregar_dados_nuvem(URL_SHAREPOINT)
    
    # Filtro dinâmico para garantir que estamos olhando apenas para Plataformas
    termo_filtro = "Plataforma"
    df_plataformas = df_bruto[
        df_bruto['origem_acidente'].astype(str).str.contains(termo_filtro, case=False, na=False) |
        df_bruto['instalacao'].astype(str).str.contains(termo_filtro, case=False, na=False)
    ]
    
    st.success(f"Dados integrados com sucesso! {len(df_plataformas)} registros de plataformas mapeados.")
    
    # Exibindo os dados na tela
    st.subheader("Visualização dos Dados")
    st.dataframe(df_plataformas)

except Exception as e:
    st.error("Não foi possível carregar os dados diretamente do SharePoint.")
    st.info("Verifique se as permissões do link compartilhado permitem acesso de leitura sem login corporativo obrigatório.")
    st.code(str(e))
