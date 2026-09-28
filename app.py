import calendar
import os
import time
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from streamlit_autorefresh import st_autorefresh


# =========================================================
# CONFIGURACION DE STREAMLIT
# =========================================================

st.set_page_config(
    page_title="Dashboard BESS",
    page_icon="⚡",
    layout="wide"
)

st_autorefresh(
    interval=15 * 60 * 1000,
    key="actualizacion_bess"
)

st.title("Dashboard de operación BESS")
st.caption(
    "Seguimiento de energía inyectada, distribución PPA y Spot, "
    "valorización y cumplimiento de budget"
)


# =========================================================
# CREDENCIALES
# =========================================================

try:
    API_KEY_PRMTE = st.secrets["API_KEY_PRMTE"]
    API_KEY_CMG = st.secrets["API_KEY_CMG"]

except KeyError:
    st.error(
        "No se encontraron las claves API en Streamlit Secrets."
    )

    st.info(
        "Configura API_KEY_PRMTE y API_KEY_CMG desde "
        "Manage app → Settings → Secrets."
    )

    st.stop()


# =========================================================
# CONFIGURACION OPERACIONAL
# =========================================================

PRECIO_PPA = 81.35
TIPO_CMG = "PRELIMINAR"

ANIO_OPERACIONAL = 2026

FECHA_INICIO = "2026-01"

# Se descarga hasta el mes actual, limitado al año operacional.
fecha_actual = pd.Timestamp.now()

if fecha_actual.year == ANIO_OPERACIONAL:
    FECHA_FIN = fecha_actual.strftime("%Y-%m")
else:
    FECHA_FIN = f"{ANIO_OPERACIONAL}-12"

BARRA_CMG = "M.ELENA_______220"

BESS_MPIDS = [
    "MARELENA_023_E1_GSS",
    "MARELENA_023_E7_GSS",
    "MARELENA_023_E8_GSS",
    "MARELENA_023_E11_GSS"
]

ARCHIVO_BUDGETS = "budgets.xlsx"
HOJA_BUDGETS = "Budgets"

URL_PRMTE = (
    "https://medidas.api.coordinador.cl/"
    "medidas-v2/measurement"
)

URL_CMG = (
    "https://sipub.api.coordinador.cl/"
    "costo-marginal-real/v4/findByDate"
)

NOMBRES_MESES = {
    1: "Ene",
    2: "Feb",
    3: "Mar",
    4: "Abr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Ago",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dic"
}

MESES_COMPLETOS = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre"
}


# =========================================================
# FUNCIONES GENERALES
# =========================================================

def generar_periodos(fecha_inicio, fecha_fin):
    inicio = datetime.strptime(fecha_inicio, "%Y-%m")
    fin = datetime.strptime(fecha_fin, "%Y-%m")

    periodos = []
    actual = inicio

    while actual <= fin:
        periodos.append(actual.strftime("%Y%m"))

        if actual.month == 12:
            actual = actual.replace(
                year=actual.year + 1,
                month=1
            )
        else:
            actual = actual.replace(
                month=actual.month + 1
            )

    return periodos


def parse_fecha_safe(serie):
    return pd.to_datetime(
        serie.astype(str).str[:19],
        errors="coerce"
    )


def convertir_numero(serie):
    return pd.to_numeric(
        serie.astype(str).str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0)


# =========================================================
# CONSULTA API PRMTE
# =========================================================

def obtener_datos_prmte(periodo, mpid):
    params = {
        "channelId": 3,
        "measurePointId": mpid,
        "period": periodo,
        "user_key": API_KEY_PRMTE
    }

    ultimo_error = None

    for intento in range(3):
        try:
            respuesta = requests.get(
                URL_PRMTE,
                params=params,
                timeout=60
            )

            respuesta.raise_for_status()

            datos = respuesta.json()

            if isinstance(datos, list):
                return datos

            return []

        except (
            requests.RequestException,
            ValueError
        ) as error:
            ultimo_error = error

            if intento < 2:
                time.sleep(2)

    raise RuntimeError(
        f"No fue posible obtener PRMTE para {mpid}, "
        f"período {periodo}. Error: {ultimo_error}"
    )


@st.cache_data(ttl=900, show_spinner=False)
def descargar_bess(fecha_inicio, fecha_fin):
    periodos = generar_periodos(
        fecha_inicio,
        fecha_fin
    )

    dataframes_mpids = []

    for mpid in BESS_MPIDS:
        registros_mpid = []

        for periodo in periodos:
            datos = obtener_datos_prmte(
                periodo + "012345",
                mpid
            )

            for bloque in datos:
                mediciones = bloque.get(
                    "measurement",
                    []
                )

                for medida in mediciones:
                    registros_mpid.append({
                        "Fecha": medida.get("dateRange"),
                        "Energia_kWh": medida.get("channel3"),
                        "MPID": mpid
                    })

        if registros_mpid:
            df_mpid = pd.DataFrame(registros_mpid)

            df_mpid["Fecha"] = parse_fecha_safe(
                df_mpid["Fecha"]
            )

            df_mpid["Energia_kWh"] = convertir_numero(
                df_mpid["Energia_kWh"]
            )

            df_mpid = df_mpid.dropna(
                subset=["Fecha"]
            )

            dataframes_mpids.append(df_mpid)

    if not dataframes_mpids:
        return pd.DataFrame(
            columns=[
                "Fecha",
                "Energia_kWh"
            ]
        )

    df_todos_mpids = pd.concat(
        dataframes_mpids,
        ignore_index=True
    )

    # Suma de la energía de los cuatro MPID por intervalo.
    df_bess = (
        df_todos_mpids
        .groupby("Fecha", as_index=False)
        .agg(
            Energia_kWh=("Energia_kWh", "sum")
        )
        .sort_values("Fecha")
    )

    return df_bess


# =========================================================
# CONSULTA API CMG
# =========================================================

@st.cache_data(ttl=900, show_spinner=False)
def descargar_cmg(inicio, fin):
    registros = []
    pagina = 0

    while True:
        params = {
            "startDate": inicio,
            "endDate": fin,
            "page": pagina,
            "limit": 5000,
            "type": TIPO_CMG,
            "bar_transf": BARRA_CMG,
            "user_key": API_KEY_CMG
        }

        respuesta = requests.get(
            URL_CMG,
            params=params,
            timeout=60
        )

        respuesta.raise_for_status()

        respuesta_json = respuesta.json()

        datos_pagina = respuesta_json.get(
            "data",
            []
        )

        if not datos_pagina:
            break

        registros.extend(datos_pagina)

        if len(datos_pagina) < 5000:
            break

        pagina += 1

    if not registros:
        return pd.DataFrame(
            columns=[
                "Fecha",
                "CMG_USD_MWh"
            ]
        )

    df_cmg = pd.DataFrame(registros)

    columnas_requeridas = [
        "fecha",
        "hra",
        "min",
        "cmg_usd_mwh_"
    ]

    faltantes = [
        columna
        for columna in columnas_requeridas
        if columna not in df_cmg.columns
    ]

    if faltantes:
        raise ValueError(
            "La API CMG no entregó las columnas esperadas: "
            + ", ".join(faltantes)
        )

    df_cmg["hra"] = pd.to_numeric(
        df_cmg["hra"],
        errors="coerce"
    ).fillna(0)

    df_cmg["min"] = pd.to_numeric(
        df_cmg["min"],
        errors="coerce"
    ).fillna(0)

    df_cmg["CMG_USD_MWh"] = convertir_numero(
        df_cmg["cmg_usd_mwh_"]
    )

    df_cmg["Fecha"] = (
        pd.to_datetime(
            df_cmg["fecha"],
            errors="coerce"
        )
        + pd.to_timedelta(
            df_cmg["hra"],
            unit="h"
        )
        + pd.to_timedelta(
            df_cmg["min"],
            unit="m"
        )
    )

    df_cmg = (
        df_cmg[
            [
                "Fecha",
                "CMG_USD_MWh"
            ]
        ]
        .dropna(subset=["Fecha"])
        .groupby("Fecha", as_index=False)
        .agg(
            CMG_USD_MWh=("CMG_USD_MWh", "mean")
        )
        .sort_values("Fecha")
    )

    return df_cmg


# =========================================================
# CARGA DEL EXCEL DE BUDGETS
# =========================================================

@st.cache_data(ttl=900, show_spinner=False)
def cargar_budgets():
    if not os.path.exists(ARCHIVO_BUDGETS):
        return None

    df_budget = pd.read_excel(
        ARCHIVO_BUDGETS,
        sheet_name=HOJA_BUDGETS,
        engine="openpyxl"
    )

    columnas_requeridas = [
        "Anio",
        "Mes",
        "Budget_Generacion_MWh",
        "Budget_PPA_MWh"
    ]

    faltantes = [
        columna
        for columna in columnas_requeridas
        if columna not in df_budget.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas en budgets.xlsx: "
            + ", ".join(faltantes)
        )

    df_budget = df_budget[
        columnas_requeridas
    ].copy()

    for columna in columnas_requeridas:
        df_budget[columna] = pd.to_numeric(
            df_budget[columna],
            errors="coerce"
        )

    df_budget = df_budget.dropna(
        subset=[
            "Anio",
            "Mes"
        ]
    )

    df_budget["Anio"] = (
        df_budget["Anio"]
        .astype(int)
    )

    df_budget["Mes"] = (
        df_budget["Mes"]
        .astype(int)
    )

    df_budget = df_budget[
        df_budget["Mes"].between(1, 12)
    ]

    df_budget = (
        df_budget
        .groupby(
            [
                "Anio",
                "Mes"
            ],
            as_index=False
        )
        .agg({
            "Budget_Generacion_MWh": "sum",
            "Budget_PPA_MWh": "sum"
        })
    )

    return df_budget


# =========================================================
# DESCARGA Y TRATAMIENTO DE LOS DATOS
# =========================================================

with st.spinner(
    "Consultando datos PRMTE y costos marginales..."
):
    try:
        df_bess = descargar_bess(
            FECHA_INICIO,
            FECHA_FIN
        )

        if df_bess.empty:
            st.warning(
                "La API PRMTE no entregó mediciones "
                "para el período seleccionado."
            )
            st.stop()

        fecha_ini_cmg = f"{FECHA_INICIO}-01"

        ultimo_dia = calendar.monthrange(
            int(FECHA_FIN[:4]),
            int(FECHA_FIN[5:7])
        )[1]

        fecha_fin_cmg = (
            f"{FECHA_FIN}-{ultimo_dia:02d}"
        )

        df_cmg = descargar_cmg(
            fecha_ini_cmg,
            fecha_fin_cmg
        )

    except Exception as error:
        st.error(
            "No fue posible descargar los datos "
            "desde las API."
        )

        st.exception(error)
        st.stop()


# =========================================================
# UNION DE ENERGIA Y CMG
# =========================================================

df = pd.merge(
    df_bess,
    df_cmg,
    on="Fecha",
    how="left"
)

df["CMG_USD_MWh"] = (
    df["CMG_USD_MWh"]
    .fillna(0)
)

df["Energia_kWh"] = (
    pd.to_numeric(
        df["Energia_kWh"],
        errors="coerce"
    )
    .fillna(0)
)

# Si la API presenta valores negativos de inyección,
# se usa el valor absoluto como energía entregada.
df["Energia_MWh"] = (
    df["Energia_kWh"].abs() / 1000
)


# =========================================================
# DISTRIBUCION PPA Y SPOT
# =========================================================

# Se considera PPA desde las 21:00 hasta antes de las 06:00.
# Se considera Spot desde las 06:00 hasta antes de las 21:00.

df["Tipo_Contrato"] = "PPA"

df.loc[
    (
        df["Fecha"].dt.hour >= 6
    )
    & (
        df["Fecha"].dt.hour < 21
    ),
    "Tipo_Contrato"
] = "SPOT"

df["Precio_Aplicado_USD_MWh"] = PRECIO_PPA

df.loc[
    df["Tipo_Contrato"] == "SPOT",
    "Precio_Aplicado_USD_MWh"
] = df["CMG_USD_MWh"]

df["Ingreso_USD"] = (
    df["Energia_MWh"]
    * df["Precio_Aplicado_USD_MWh"]
)

df["Fecha_Dia"] = (
    df["Fecha"]
    .dt.normalize()
)

df["Anio"] = df["Fecha"].dt.year
df["Mes_Numero"] = df["Fecha"].dt.month


# =========================================================
# RESUMEN DIARIO
# =========================================================

resumen_diario_largo = (
    df.groupby(
        [
            "Fecha_Dia",
            "Tipo_Contrato"
        ],
        as_index=False
    )
    .agg(
        Energia_MWh=("Energia_MWh", "sum"),
        Ingreso_USD=("Ingreso_USD", "sum")
    )
)

energia_diaria = (
    resumen_diario_largo
    .pivot_table(
        index="Fecha_Dia",
        columns="Tipo_Contrato",
        values="Energia_MWh",
        aggfunc="sum",
        fill_value=0
    )
    .reset_index()
)

energia_diaria.columns.name = None

if "PPA" not in energia_diaria.columns:
    energia_diaria["PPA"] = 0

if "SPOT" not in energia_diaria.columns:
    energia_diaria["SPOT"] = 0

energia_diaria = energia_diaria.rename(
    columns={
        "PPA": "Energia_PPA_MWh",
        "SPOT": "Energia_Spot_MWh"
    }
)

energia_diaria["Energia_Total_MWh"] = (
    energia_diaria["Energia_PPA_MWh"]
    + energia_diaria["Energia_Spot_MWh"]
)


# =========================================================
# RESUMEN MENSUAL
# =========================================================

resumen_mensual_largo = (
    df.groupby(
        [
            "Anio",
            "Mes_Numero",
            "Tipo_Contrato"
        ],
        as_index=False
    )
    .agg(
        Energia_MWh=("Energia_MWh", "sum"),
        Ingreso_USD=("Ingreso_USD", "sum")
    )
)

energia_mensual = (
    resumen_mensual_largo
    .pivot_table(
        index=[
            "Anio",
            "Mes_Numero"
        ],
        columns="Tipo_Contrato",
        values="Energia_MWh",
        aggfunc="sum",
        fill_value=0
    )
    .reset_index()
)

energia_mensual.columns.name = None

if "PPA" not in energia_mensual.columns:
    energia_mensual["PPA"] = 0

if "SPOT" not in energia_mensual.columns:
    energia_mensual["SPOT"] = 0

energia_mensual = energia_mensual.rename(
    columns={
        "PPA": "Energia_PPA_MWh",
        "SPOT": "Energia_Spot_MWh"
    }
)

energia_mensual["Energia_Total_MWh"] = (
    energia_mensual["Energia_PPA_MWh"]
    + energia_mensual["Energia_Spot_MWh"]
)


ingreso_mensual = (
    resumen_mensual_largo
    .pivot_table(
        index=[
            "Anio",
            "Mes_Numero"
        ],
        columns="Tipo_Contrato",
        values="Ingreso_USD",
        aggfunc="sum",
        fill_value=0
    )
    .reset_index()
)

ingreso_mensual.columns.name = None

if "PPA" not in ingreso_mensual.columns:
    ingreso_mensual["PPA"] = 0

if "SPOT" not in ingreso_mensual.columns:
    ingreso_mensual["SPOT"] = 0

ingreso_mensual = ingreso_mensual.rename(
    columns={
        "PPA": "Ingreso_PPA_USD",
        "SPOT": "Ingreso_Spot_USD"
    }
)

resumen_mensual = pd.merge(
    energia_mensual,
    ingreso_mensual,
    on=[
        "Anio",
        "Mes_Numero"
    ],
    how="outer"
).fillna(0)

resumen_mensual["Ingreso_Total_USD"] = (
    resumen_mensual["Ingreso_PPA_USD"]
    + resumen_mensual["Ingreso_Spot_USD"]
)

resumen_mensual[
    "Precio_Promedio_USD_MWh"
] = resumen_mensual.apply(
    lambda fila: (
        fila["Ingreso_Total_USD"]
        / fila["Energia_Total_MWh"]
        if fila["Energia_Total_MWh"] > 0
        else 0
    ),
    axis=1
)


# =========================================================
# FILTROS
# =========================================================

st.sidebar.header("Filtros")

anios_disponibles = sorted(
    df["Anio"].unique().tolist(),
    reverse=True
)

anio_seleccionado = st.sidebar.selectbox(
    "Año",
    options=anios_disponibles,
    index=0
)

meses_disponibles = sorted(
    df.loc[
        df["Anio"] == anio_seleccionado,
        "Mes_Numero"
    ]
    .unique()
    .tolist(),
    reverse=True
)

mes_seleccionado = st.sidebar.selectbox(
    "Mes para gráfico diario",
    options=meses_disponibles,
    format_func=lambda mes: MESES_COMPLETOS[mes],
    index=0
)


# =========================================================
# KPIS
# =========================================================

df_anio = df[
    df["Anio"] == anio_seleccionado
].copy()

energia_total_anual = (
    df_anio["Energia_MWh"].sum()
)

energia_ppa_anual = (
    df_anio.loc[
        df_anio["Tipo_Contrato"] == "PPA",
        "Energia_MWh"
    ]
    .sum()
)

energia_spot_anual = (
    df_anio.loc[
        df_anio["Tipo_Contrato"] == "SPOT",
        "Energia_MWh"
    ]
    .sum()
)

ingreso_total_anual = (
    df_anio["Ingreso_USD"].sum()
)

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "Energía BESS anual",
    f"{energia_total_anual:,.1f} MWh"
)

col2.metric(
    "Energía PPA anual",
    f"{energia_ppa_anual:,.1f} MWh"
)

col3.metric(
    "Energía Spot anual",
    f"{energia_spot_anual:,.1f} MWh"
)

col4.metric(
    "Ingreso total anual",
    f"USD {ingreso_total_anual:,.0f}"
)


# =========================================================
# GRAFICO DIARIO
# =========================================================

st.subheader(
    "Distribución diaria de energía "
    f"{MESES_COMPLETOS[mes_seleccionado]} "
    f"{anio_seleccionado}"
)

diario_mes = energia_diaria[
    (
        energia_diaria["Fecha_Dia"].dt.year
        == anio_seleccionado
    )
    & (
        energia_diaria["Fecha_Dia"].dt.month
        == mes_seleccionado
    )
].copy()

primer_dia_mes = pd.Timestamp(
    year=anio_seleccionado,
    month=mes_seleccionado,
    day=1
)

if (
    anio_seleccionado == fecha_actual.year
    and mes_seleccionado == fecha_actual.month
):
    ultimo_dia_grafico = fecha_actual.normalize()

else:
    cantidad_dias = calendar.monthrange(
        anio_seleccionado,
        mes_seleccionado
    )[1]

    ultimo_dia_grafico = pd.Timestamp(
        year=anio_seleccionado,
        month=mes_seleccionado,
        day=cantidad_dias
    )

calendario_mes = pd.DataFrame({
    "Fecha_Dia": pd.date_range(
        start=primer_dia_mes,
        end=ultimo_dia_grafico,
        freq="D"
    )
})

diario_mes = calendario_mes.merge(
    diario_mes,
    on="Fecha_Dia",
    how="left"
)

columnas_diarias = [
    "Energia_PPA_MWh",
    "Energia_Spot_MWh",
    "Energia_Total_MWh"
]

diario_mes[columnas_diarias] = (
    diario_mes[columnas_diarias]
    .fillna(0)
)

ancho_barra_diaria = (
    0.42 * 24 * 60 * 60 * 1000
)

fig_diario = go.Figure()

fig_diario.add_trace(
    go.Bar(
        x=diario_mes["Fecha_Dia"],
        y=diario_mes["Energia_PPA_MWh"],
        name="Energía PPA",
        marker_color="#7DBCF0",
        marker_line_color="#1683DB",
        marker_line_width=0.5,
        width=ancho_barra_diaria,
        hovertemplate=(
            "<b>%{x|%d-%m-%Y}</b><br>"
            "PPA: %{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

fig_diario.add_trace(
    go.Bar(
        x=diario_mes["Fecha_Dia"],
        y=diario_mes["Energia_Spot_MWh"],
        name="Energía Spot",
        marker_color="#0A69C7",
        marker_line_color="#1683DB",
        marker_line_width=0.5,
        width=ancho_barra_diaria,
        hovertemplate=(
            "<b>%{x|%d-%m-%Y}</b><br>"
            "Spot: %{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

fig_diario.update_layout(
    barmode="stack",
    xaxis_title="Día del mes",
    yaxis_title="Energía [MWh]",
    legend_title="Destino",
    hovermode="x unified",
    height=500,
    bargap=0.45,
    margin=dict(
        l=30,
        r=30,
        t=30,
        b=30
    )
)

fig_diario.update_xaxes(
    tickmode="linear",
    dtick=24 * 60 * 60 * 1000,
    tickformat="%d",
    range=[
        primer_dia_mes - pd.Timedelta(hours=12),
        ultimo_dia_grafico + pd.Timedelta(hours=12)
    ],
    showgrid=False
)

fig_diario.update_yaxes(
    rangemode="tozero",
    gridcolor="rgba(140, 140, 140, 0.25)"
)

st.plotly_chart(
    fig_diario,
    use_container_width=True
)


# =========================================================
# CONSOLIDADO ANUAL
# =========================================================

st.subheader(
    f"Consolidado anual {anio_seleccionado}"
)

consolidado_anual = pd.DataFrame({
    "Mes_Numero": range(1, 13)
})

datos_anuales = resumen_mensual[
    resumen_mensual["Anio"]
    == anio_seleccionado
].copy()

consolidado_anual = consolidado_anual.merge(
    datos_anuales,
    on="Mes_Numero",
    how="left"
)

columnas_consolidado = [
    "Energia_PPA_MWh",
    "Energia_Spot_MWh",
    "Energia_Total_MWh",
    "Ingreso_PPA_USD",
    "Ingreso_Spot_USD",
    "Ingreso_Total_USD"
]

for columna in columnas_consolidado:
    if columna not in consolidado_anual.columns:
        consolidado_anual[columna] = 0

consolidado_anual[columnas_consolidado] = (
    consolidado_anual[columnas_consolidado]
    .fillna(0)
)

try:
    df_budgets = cargar_budgets()

except Exception as error:
    st.error(
        "No fue posible leer budgets.xlsx."
    )
    st.exception(error)
    df_budgets = None

if df_budgets is not None:
    budgets_anio = df_budgets[
        df_budgets["Anio"]
        == anio_seleccionado
    ].copy()

    consolidado_anual = consolidado_anual.merge(
        budgets_anio[
            [
                "Mes",
                "Budget_Generacion_MWh",
                "Budget_PPA_MWh"
            ]
        ],
        left_on="Mes_Numero",
        right_on="Mes",
        how="left"
    )

else:
    consolidado_anual[
        "Budget_Generacion_MWh"
    ] = pd.NA

    consolidado_anual[
        "Budget_PPA_MWh"
    ] = pd.NA

    st.info(
        "Carga budgets.xlsx en la carpeta principal "
        "para mostrar las líneas de budget."
    )

consolidado_anual["Nombre_Mes"] = (
    consolidado_anual["Mes_Numero"]
    .map(NOMBRES_MESES)
)

fig_anual = go.Figure()

fig_anual.add_trace(
    go.Bar(
        x=consolidado_anual["Nombre_Mes"],
        y=consolidado_anual["Energia_PPA_MWh"],
        name="Energía PPA real",
        marker_color="#149447",
        marker_line_color="#1559A6",
        marker_line_width=1.5,
        width=0.58,
        hovertemplate=(
            "<b>%{x}</b><br>"
            "PPA real: %{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

fig_anual.add_trace(
    go.Bar(
        x=consolidado_anual["Nombre_Mes"],
        y=consolidado_anual["Energia_Spot_MWh"],
        name="Energía Spot real",
        marker_color="#D94141",
        marker_line_color="#1559A6",
        marker_line_width=1.5,
        width=0.58,
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Spot real: %{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

fig_anual.add_trace(
    go.Scatter(
        x=consolidado_anual["Nombre_Mes"],
        y=consolidado_anual[
            "Budget_Generacion_MWh"
        ],
        name="Budget de generación",
        mode="lines+markers",
        line=dict(
            color="#FF3333",
            width=3
        ),
        marker=dict(
            color="#FF3333",
            size=8
        ),
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Budget generación: %{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

fig_anual.add_trace(
    go.Scatter(
        x=consolidado_anual["Nombre_Mes"],
        y=consolidado_anual["Budget_PPA_MWh"],
        name="Budget contrato PPA",
        mode="lines+markers",
        line=dict(
            color="#111111",
            width=3
        ),
        marker=dict(
            color="#111111",
            size=8
        ),
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Budget PPA: %{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

fig_anual.update_layout(
    barmode="stack",
    xaxis_title="Mes",
    yaxis_title="Energía [MWh]",
    legend_title="Serie",
    hovermode="x unified",
    height=600,
    bargap=0.35,
    margin=dict(
        l=30,
        r=30,
        t=30,
        b=30
    )
)

fig_anual.update_xaxes(
    categoryorder="array",
    categoryarray=[
        "Ene",
        "Feb",
        "Mar",
        "Abr",
        "May",
        "Jun",
        "Jul",
        "Ago",
        "Sep",
        "Oct",
        "Nov",
        "Dic"
    ],
    showgrid=False
)

fig_anual.update_yaxes(
    rangemode="tozero",
    gridcolor="rgba(140, 140, 140, 0.25)"
)

st.plotly_chart(
    fig_anual,
    use_container_width=True
)


# =========================================================
# CUMPLIMIENTO
# =========================================================

if df_budgets is not None:
    meses_con_datos = consolidado_anual[
        "Energia_Total_MWh"
    ] > 0

    budget_generacion_acumulado = (
        consolidado_anual.loc[
            meses_con_datos,
            "Budget_Generacion_MWh"
        ]
        .fillna(0)
        .sum()
    )

    budget_ppa_acumulado = (
        consolidado_anual.loc[
            meses_con_datos,
            "Budget_PPA_MWh"
        ]
        .fillna(0)
        .sum()
    )

    cumplimiento_generacion = (
        energia_total_anual
        / budget_generacion_acumulado
        * 100
        if budget_generacion_acumulado > 0
        else 0
    )

    cumplimiento_ppa = (
        energia_ppa_anual
        / budget_ppa_acumulado
        * 100
        if budget_ppa_acumulado > 0
        else 0
    )

    col_budget_1, col_budget_2 = st.columns(2)

    col_budget_1.metric(
        "Cumplimiento acumulado generación",
        f"{cumplimiento_generacion:,.1f} %"
    )

    col_budget_2.metric(
        "Cumplimiento acumulado PPA",
        f"{cumplimiento_ppa:,.1f} %"
    )


# =========================================================
# INFORMACION FINAL
# =========================================================

st.caption(
    "Última medición disponible: "
    f"{df['Fecha'].max():%d-%m-%Y %H:%M}"
)

st.caption(
    "Dashboard actualizado: "
    f"{datetime.now():%d-%m-%Y %H:%M:%S}"
)