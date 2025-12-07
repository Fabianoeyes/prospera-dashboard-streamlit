from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Dashboard Prospera",
    layout="wide"
)

st.title("Dashboard Prospera – Versão Web")
st.caption("Baseado na planilha 'Plano Estratégico - 2026.xlsm'")

def get_excel_path() -> Path:
    """Return the path to the strategic plan spreadsheet.

    The primary expected file name is ``Plano Estratégico - 2026.xlsm``. However,
    to make local and deployed environments more resilient to Unicode filename
    normalisation differences, we fallback to the first ``*.xlsm`` file that
    matches the "Plano Estrat" prefix.
    """

    base_dir = Path(__file__).resolve().parent
    default_path = base_dir / "Plano Estratégico - 2026.xlsm"

    if default_path.exists():
        return default_path

    alternatives = sorted(base_dir.glob("Plano Estrat*2026*.xlsm"))
    if alternatives:
        return alternatives[0]

    raise FileNotFoundError(
        "Não foi possível localizar a planilha 'Plano Estratégico - 2026.xlsm'."
    )


@st.cache_data
def load_excel():
    excel_path = get_excel_path()
    xls = pd.read_excel(
        excel_path,
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

# Inicializa o estado de edição para cada aba
if "abas_editaveis" not in st.session_state:
    st.session_state["abas_editaveis"] = {}

if aba_selecionada not in st.session_state["abas_editaveis"]:
    # Convertemos para ``object`` para permitir textos e fórmulas mesmo em colunas numéricas.
    st.session_state["abas_editaveis"][aba_selecionada] = dados[aba_selecionada].astype(object)

df_editavel = st.session_state["abas_editaveis"][aba_selecionada]

st.subheader(f"Aba selecionada: {aba_selecionada}")
df_editado = st.data_editor(
    df_editavel,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    key=f"editor_{aba_selecionada}",
)

st.caption(
    "Clique em uma célula para editar. É possível colar valores ou fórmulas (texto)."
)

# Atualiza o estado compartilhado com a versão editada para a aba atual
st.session_state["abas_editaveis"][aba_selecionada] = df_editado

st.info(
    "Esse é o app base. Depois transformaremos as fórmulas do Excel em cálculos Python "
    "e construiremos dashboards, KPIs e gráficos."
)
