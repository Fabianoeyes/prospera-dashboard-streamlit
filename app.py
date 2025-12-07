import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Dashboard Prospera",
    layout="wide"
)

st.title("Dashboard Prospera – Versão Web")
st.caption("Baseado na planilha 'Plano Estratégico - 2026.xlsm'")

@st.cache_data
def load_excel():
    xls = pd.read_excel(
        "Plano Estratégico - 2026.xlsm",
        sheet_name=None,
        engine="openpyxl"
    )
    return xls

try:
    dados = load_excel()
except Exception as e:
    st.error(f"Erro ao carregar a planilha: {e}")
    st.stop()

abas = list(dados.keys())

st.sidebar.header("Navegação")
aba_selecionada = st.sidebar.selectbox("Selecione a aba da planilha:", abas)

df = dados[aba_selecionada]

st.subheader(f"Aba selecionada: {aba_selecionada}")
st.dataframe(df, use_container_width=True)

st.info(
    "Esse é o app base. Depois transformaremos as fórmulas do Excel em cálculos Python "
    "e construiremos dashboards, KPIs e gráficos."
)
