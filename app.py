import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
from streamlit_autorefresh import st_autorefresh

st.set_page_config(
    page_title="Dashboard BESS",
    page_icon="⚡",
    layout="wide"
)

# Actualiza la pantalla cada 15 minutos
st_autorefresh(
    interval=15 * 60 * 1000,
    key="actualizacion_bess"
)

st.title("Dashboard de operación BESS")
st.caption("Seguimiento de energía inyectada, PPA y Spot")

@st.cache_data(ttl=900)
def cargar_datos():
    # Aquí se reemplaza el ejemplo por la consulta real a las API
    datos = {
        "Fecha": pd.date_range(
            start="2026-09-01",
            periods=10,
            freq="D"
        ),
        "Energia_BESS_MWh": [
            250, 270, 265, 280, 290,
            275, 300, 285, 295, 310
        ],
        "Energia_PPA_MWh": [
            180, 190, 185, 200, 205,
            190, 210, 195, 205, 220
        ]
    }

    df = pd.DataFrame(datos)

    df["Energia_Spot_MWh"] = (
        df["Energia_BESS_MWh"]
        - df["Energia_PPA_MWh"]
    )

    return df


df = cargar_datos()

energia_total = df["Energia_BESS_MWh"].sum()
energia_ppa = df["Energia_PPA_MWh"].sum()
energia_spot = df["Energia_Spot_MWh"].sum()

col1, col2, col3 = st.columns(3)

col1.metric(
    "Energía BESS",
    f"{energia_total:,.1f} MWh"
)

col2.metric(
    "Energía PPA",
    f"{energia_ppa:,.1f} MWh"
)

col3.metric(
    "Energía Spot",
    f"{energia_spot:,.1f} MWh"
)

figura = px.bar(
    df,
    x="Fecha",
    y=[
        "Energia_PPA_MWh",
        "Energia_Spot_MWh"
    ],
    title="Distribución diaria de energía",
    labels={
        "value": "Energía [MWh]",
        "variable": "Destino"
    },
    barmode="stack"
)

st.plotly_chart(
    figura,
    use_container_width=True
)

st.subheader("Detalle diario")

st.dataframe(
    df,
    use_container_width=True,
    hide_index=True
)

st.caption(
    f"Dashboard actualizado: "
    f"{datetime.now():%d-%m-%Y %H:%M:%S}"
)