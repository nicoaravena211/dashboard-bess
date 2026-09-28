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

# Actualiza la pantalla cada 12 horas.
DOCE_HORAS_MS = 12 * 60 * 60 * 1000
DOCE_HORAS_SEGUNDOS = 12 * 60 * 60

st_autorefresh(
    interval=DOCE_HORAS_MS,
    key="actualizacion_bess_12h"
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
        "No se encontraron las credenciales API en "
        "los Secrets de Streamlit."
    )

    st.info(
        "Configura API_KEY_PRMTE y API_KEY_CMG desde "
        "Manage app > Settings > Secrets."
    )

    st.stop()


# =========================================================
# CONFIGURACION OPERACIONAL
# =========================================================

PRECIO_PPA = 83.49
TIPO_CMG = "PRELIMINAR"

ANIO_OPERACIONAL = 2026
FECHA_INICIO = "2026-01"

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

ARCHIVO_EXCEL = "budgets.xlsx"
HOJA_BUDGETS = "Budgets"
HOJA_CONSOLIDADO = "Consolidado"

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
        serie.astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False),
        errors="coerce"
    ).fillna(0)


def periodo_a_texto(periodo):
    anio = int(periodo[:4])
    mes = int(periodo[4:6])

    return f"{MESES_COMPLETOS[mes]} {anio}"


# =========================================================
# LECTURA DE BUDGETS
# =========================================================

@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def cargar_budgets():
    if not os.path.exists(ARCHIVO_EXCEL):
        raise FileNotFoundError(
            f"No se encontró {ARCHIVO_EXCEL}."
        )

    df_budget = pd.read_excel(
        ARCHIVO_EXCEL,
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
            "Faltan columnas en la hoja Budgets: "
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
        subset=["Anio", "Mes"]
    )

    df_budget["Anio"] = (
        df_budget["Anio"].astype(int)
    )

    df_budget["Mes"] = (
        df_budget["Mes"].astype(int)
    )

    df_budget = df_budget[
        df_budget["Mes"].between(1, 12)
    ].copy()

    duplicados = df_budget.duplicated(
        subset=["Anio", "Mes"],
        keep=False
    )

    if duplicados.any():
        meses_duplicados = (
            df_budget.loc[
                duplicados,
                ["Anio", "Mes"]
            ]
            .drop_duplicates()
            .astype(str)
            .agg("-".join, axis=1)
            .tolist()
        )

        raise ValueError(
            "Hay meses duplicados en la hoja Budgets: "
            + ", ".join(meses_duplicados)
        )

    return df_budget


# =========================================================
# LECTURA DEL CONSOLIDADO MANUAL
# =========================================================

@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def cargar_consolidado_excel():
    if not os.path.exists(ARCHIVO_EXCEL):
        raise FileNotFoundError(
            f"No se encontró {ARCHIVO_EXCEL}."
        )

    df_consolidado = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name=HOJA_CONSOLIDADO,
        engine="openpyxl"
    )

    columnas_requeridas = [
        "Anio",
        "Mes",
        "Energia_PPA_MWh",
        "Energia_Spot_MWh",
        "Ingreso_PPA_USD",
        "Ingreso_Spot_USD"
    ]

    faltantes = [
        columna
        for columna in columnas_requeridas
        if columna not in df_consolidado.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas en la hoja Consolidado: "
            + ", ".join(faltantes)
        )

    df_consolidado = df_consolidado.copy()

    columnas_numericas = [
        "Anio",
        "Mes",
        "Energia_PPA_MWh",
        "Energia_Spot_MWh",
        "Ingreso_PPA_USD",
        "Ingreso_Spot_USD"
    ]

    for columna in columnas_numericas:
        df_consolidado[columna] = pd.to_numeric(
            df_consolidado[columna],
            errors="coerce"
        )

    df_consolidado = df_consolidado.dropna(
        subset=["Anio", "Mes"]
    )

    df_consolidado["Anio"] = (
        df_consolidado["Anio"].astype(int)
    )

    df_consolidado["Mes"] = (
        df_consolidado["Mes"].astype(int)
    )

    df_consolidado = df_consolidado[
        df_consolidado["Mes"].between(1, 12)
    ].copy()

    duplicados = df_consolidado.duplicated(
        subset=["Anio", "Mes"],
        keep=False
    )

    if duplicados.any():
        meses_duplicados = (
            df_consolidado.loc[
                duplicados,
                ["Anio", "Mes"]
            ]
            .drop_duplicates()
            .astype(str)
            .agg("-".join, axis=1)
            .tolist()
        )

        raise ValueError(
            "Hay meses duplicados en la hoja Consolidado: "
            + ", ".join(meses_duplicados)
        )

    columnas_valores = [
        "Energia_PPA_MWh",
        "Energia_Spot_MWh",
        "Ingreso_PPA_USD",
        "Ingreso_Spot_USD"
    ]

    if df_consolidado[columnas_valores].isna().any().any():
        raise ValueError(
            "Hay valores vacíos o no numéricos en la hoja "
            "Consolidado."
        )

    if (
        df_consolidado[columnas_valores] < 0
    ).any().any():
        raise ValueError(
            "La hoja Consolidado contiene valores negativos. "
            "Revisa las energías e ingresos."
        )

    df_consolidado["Energia_Total_MWh"] = (
        df_consolidado["Energia_PPA_MWh"]
        + df_consolidado["Energia_Spot_MWh"]
    )

    df_consolidado["Ingreso_Total_USD"] = (
        df_consolidado["Ingreso_PPA_USD"]
        + df_consolidado["Ingreso_Spot_USD"]
    )

    df_consolidado["Precio_Promedio_USD_MWh"] = (
        df_consolidado["Ingreso_Total_USD"]
        / df_consolidado["Energia_Total_MWh"]
        .replace(0, pd.NA)
    ).fillna(0)

    if "Fuente" not in df_consolidado.columns:
        df_consolidado["Fuente"] = "Excel consolidado"
    else:
        df_consolidado["Fuente"] = (
            df_consolidado["Fuente"]
            .fillna("Excel consolidado")
            .astype(str)
        )

    return df_consolidado


# =========================================================
# PERIODOS PENDIENTES
# =========================================================

def obtener_periodos_pendientes(
    fecha_inicio,
    fecha_fin,
    df_consolidado
):
    periodos_totales = generar_periodos(
        fecha_inicio,
        fecha_fin
    )

    periodos_excel = set()

    for _, fila in df_consolidado.iterrows():
        periodo = (
            f"{int(fila['Anio']):04d}"
            f"{int(fila['Mes']):02d}"
        )

        periodos_excel.add(periodo)

    periodos_pendientes = [
        periodo
        for periodo in periodos_totales
        if periodo not in periodos_excel
    ]

    return periodos_totales, periodos_pendientes


# =========================================================
# API PRMTE
# =========================================================

def obtener_datos_prmte(periodo, mpid):
    params = {
        "channelId": 3,
        "measurePointId": mpid,
        "period": periodo,
        "user_key": API_KEY_PRMTE
    }

    for intento in range(3):
        try:
            respuesta = requests.get(
                URL_PRMTE,
                params=params,
                timeout=60
            )

            if respuesta.status_code == 200:
                datos = respuesta.json()

                if isinstance(datos, list):
                    return datos

                return []

        except (
            requests.RequestException,
            ValueError
        ):
            pass

        if intento < 2:
            time.sleep(2)

    return None


@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def descargar_bess_periodos(periodos):
    if not periodos:
        return (
            pd.DataFrame(
                columns=["Fecha", "Energia_kWh"]
            ),
            []
        )

    registros_totales = []
    errores = []

    for periodo in periodos:
        periodo_valido = True
        registros_periodo = []

        for mpid in BESS_MPIDS:
            datos = obtener_datos_prmte(
                periodo + "012345",
                mpid
            )

            if datos is None:
                periodo_valido = False

                errores.append(
                    f"{periodo_a_texto(periodo)}: "
                    f"PRMTE no disponible para {mpid}."
                )

                continue

            for bloque in datos:
                for medida in bloque.get(
                    "measurement",
                    []
                ):
                    registros_periodo.append({
                        "Fecha": medida.get("dateRange"),
                        "Energia_kWh": medida.get("channel3"),
                        "MPID": mpid,
                        "Periodo": periodo
                    })

        if periodo_valido and registros_periodo:
            registros_totales.extend(
                registros_periodo
            )

        elif not registros_periodo:
            errores.append(
                f"{periodo_a_texto(periodo)}: "
                "el período no entregó mediciones."
            )

    if not registros_totales:
        return (
            pd.DataFrame(
                columns=["Fecha", "Energia_kWh"]
            ),
            sorted(set(errores))
        )

    df_mediciones = pd.DataFrame(
        registros_totales
    )

    df_mediciones["Fecha"] = parse_fecha_safe(
        df_mediciones["Fecha"]
    )

    df_mediciones["Energia_kWh"] = convertir_numero(
        df_mediciones["Energia_kWh"]
    )

    df_mediciones = df_mediciones.dropna(
        subset=["Fecha"]
    )

    df_bess = (
        df_mediciones
        .groupby("Fecha", as_index=False)
        .agg(
            Energia_kWh=("Energia_kWh", "sum")
        )
        .sort_values("Fecha")
    )

    return df_bess, sorted(set(errores))


# =========================================================
# API CMG
# =========================================================

def descargar_cmg_mes(periodo):
    anio = int(periodo[:4])
    mes = int(periodo[4:6])

    primer_dia = pd.Timestamp(
        year=anio,
        month=mes,
        day=1
    )

    ultimo_dia_numero = calendar.monthrange(
        anio,
        mes
    )[1]

    ultimo_dia = pd.Timestamp(
        year=anio,
        month=mes,
        day=ultimo_dia_numero
    )

    hoy = pd.Timestamp.now().normalize()

    if (
        anio == hoy.year
        and mes == hoy.month
    ):
        ultimo_dia = min(
            ultimo_dia,
            hoy
        )

    registros = []
    pagina = 0

    while True:
        params = {
            "startDate": primer_dia.strftime("%Y-%m-%d"),
            "endDate": ultimo_dia.strftime("%Y-%m-%d"),
            "page": pagina,
            "limit": 5000,
            "type": TIPO_CMG,
            "bar_transf": BARRA_CMG,
            "user_key": API_KEY_CMG
        }

        try:
            respuesta = requests.get(
                URL_CMG,
                params=params,
                timeout=60
            )

            if respuesta.status_code != 200:
                return None

            respuesta_json = respuesta.json()

        except (
            requests.RequestException,
            ValueError
        ):
            return None

        datos = respuesta_json.get(
            "data",
            []
        )

        if not datos:
            break

        registros.extend(datos)

        if len(datos) < 5000:
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

    if not all(
        columna in df_cmg.columns
        for columna in columnas_requeridas
    ):
        return None

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
            CMG_USD_MWh=(
                "CMG_USD_MWh",
                "mean"
            )
        )
    )

    return df_cmg


@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def descargar_cmg_periodos(periodos):
    dataframes = []
    errores = []

    for periodo in periodos:
        df_mes = descargar_cmg_mes(periodo)

        if df_mes is None:
            errores.append(
                f"{periodo_a_texto(periodo)}: "
                "CMG no disponible."
            )

            continue

        if not df_mes.empty:
            dataframes.append(df_mes)

    if not dataframes:
        return (
            pd.DataFrame(
                columns=[
                    "Fecha",
                    "CMG_USD_MWh"
                ]
            ),
            sorted(set(errores))
        )

    df_cmg = (
        pd.concat(
            dataframes,
            ignore_index=True
        )
        .groupby("Fecha", as_index=False)
        .agg(
            CMG_USD_MWh=(
                "CMG_USD_MWh",
                "mean"
            )
        )
        .sort_values("Fecha")
    )

    return df_cmg, sorted(set(errores))


# =========================================================
# VALIDACION DEL EXCEL
# =========================================================

try:
    df_budgets = cargar_budgets()
    df_consolidado_excel = cargar_consolidado_excel()

except Exception as error:
    st.error(
        "No fue posible validar budgets.xlsx."
    )

    st.info(str(error))

    st.stop()


periodos_totales, periodos_pendientes = (
    obtener_periodos_pendientes(
        FECHA_INICIO,
        FECHA_FIN,
        df_consolidado_excel
    )
)

periodos_consolidados = [
    periodo
    for periodo in periodos_totales
    if periodo not in periodos_pendientes
]


# =========================================================
# INFORMACION DE LAS FUENTES
# =========================================================

with st.expander(
    "Estado de las fuentes de información",
    expanded=False
):
    st.write(
        "**Meses cargados desde el Excel:**"
    )

    if periodos_consolidados:
        st.write(
            ", ".join(
                periodo_a_texto(periodo)
                for periodo in periodos_consolidados
            )
        )
    else:
        st.write("Ninguno")

    st.write(
        "**Meses que se consultarán desde las API:**"
    )

    if periodos_pendientes:
        st.write(
            ", ".join(
                periodo_a_texto(periodo)
                for periodo in periodos_pendientes
            )
        )
    else:
        st.write("Ninguno")


# =========================================================
# DESCARGA DE PERIODOS NO CONSOLIDADOS
# =========================================================

with st.spinner(
    "Consultando períodos pendientes en las API..."
):
    df_bess, errores_prmte = (
        descargar_bess_periodos(
            periodos_pendientes
        )
    )

    df_cmg, errores_cmg = (
        descargar_cmg_periodos(
            periodos_pendientes
        )
    )


errores_api = sorted(
    set(errores_prmte + errores_cmg)
)

if errores_api:
    st.warning(
        "Algunos períodos no estuvieron disponibles en las API. "
        "El dashboard continuará con la información disponible."
    )

    with st.expander(
        "Ver advertencias de las API"
    ):
        for error in errores_api:
            st.write(f"- {error}")


# =========================================================
# PROCESAMIENTO DE LA INFORMACION API
# =========================================================

df_api = pd.DataFrame()
resumen_diario_api = pd.DataFrame()
resumen_mensual_api = pd.DataFrame()

if not df_bess.empty:
    df_api = pd.merge(
        df_bess,
        df_cmg,
        on="Fecha",
        how="left"
    )

    df_api["CMG_USD_MWh"] = (
        df_api["CMG_USD_MWh"]
        .fillna(0)
    )

    df_api["Energia_MWh"] = (
        pd.to_numeric(
            df_api["Energia_kWh"],
            errors="coerce"
        )
        .fillna(0)
        .abs()
        / 1000
    )

    # PPA: 21:00 a 05:59
    # Spot: 06:00 a 20:59

    df_api["Tipo_Contrato"] = "PPA"

    df_api.loc[
        (
            df_api["Fecha"].dt.hour >= 6
        )
        & (
            df_api["Fecha"].dt.hour < 21
        ),
        "Tipo_Contrato"
    ] = "SPOT"

    df_api["Precio_Aplicado_USD_MWh"] = (
        PRECIO_PPA
    )

    df_api.loc[
        df_api["Tipo_Contrato"] == "SPOT",
        "Precio_Aplicado_USD_MWh"
    ] = df_api["CMG_USD_MWh"]

    df_api["Ingreso_USD"] = (
        df_api["Energia_MWh"]
        * df_api["Precio_Aplicado_USD_MWh"]
    )

    df_api["Fecha_Dia"] = (
        df_api["Fecha"].dt.normalize()
    )

    df_api["Anio"] = (
        df_api["Fecha"].dt.year
    )

    df_api["Mes"] = (
        df_api["Fecha"].dt.month
    )

    # -----------------------------------------------
    # RESUMEN DIARIO API
    # -----------------------------------------------

    diario_largo = (
        df_api
        .groupby(
            [
                "Fecha_Dia",
                "Tipo_Contrato"
            ],
            as_index=False
        )
        .agg(
            Energia_MWh=(
                "Energia_MWh",
                "sum"
            )
        )
    )

    resumen_diario_api = (
        diario_largo
        .pivot_table(
            index="Fecha_Dia",
            columns="Tipo_Contrato",
            values="Energia_MWh",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
    )

    resumen_diario_api.columns.name = None

    if "PPA" not in resumen_diario_api.columns:
        resumen_diario_api["PPA"] = 0

    if "SPOT" not in resumen_diario_api.columns:
        resumen_diario_api["SPOT"] = 0

    resumen_diario_api = resumen_diario_api.rename(
        columns={
            "PPA": "Energia_PPA_MWh",
            "SPOT": "Energia_Spot_MWh"
        }
    )

    resumen_diario_api["Energia_Total_MWh"] = (
        resumen_diario_api["Energia_PPA_MWh"]
        + resumen_diario_api["Energia_Spot_MWh"]
    )

    # -----------------------------------------------
    # RESUMEN MENSUAL API
    # -----------------------------------------------

    mensual_largo = (
        df_api
        .groupby(
            [
                "Anio",
                "Mes",
                "Tipo_Contrato"
            ],
            as_index=False
        )
        .agg(
            Energia_MWh=(
                "Energia_MWh",
                "sum"
            ),
            Ingreso_USD=(
                "Ingreso_USD",
                "sum"
            )
        )
    )

    energia_mensual = (
        mensual_largo
        .pivot_table(
            index=["Anio", "Mes"],
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

    ingreso_mensual = (
        mensual_largo
        .pivot_table(
            index=["Anio", "Mes"],
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

    resumen_mensual_api = pd.merge(
        energia_mensual,
        ingreso_mensual,
        on=["Anio", "Mes"],
        how="outer"
    ).fillna(0)

    resumen_mensual_api["Energia_Total_MWh"] = (
        resumen_mensual_api["Energia_PPA_MWh"]
        + resumen_mensual_api["Energia_Spot_MWh"]
    )

    resumen_mensual_api["Ingreso_Total_USD"] = (
        resumen_mensual_api["Ingreso_PPA_USD"]
        + resumen_mensual_api["Ingreso_Spot_USD"]
    )

    resumen_mensual_api[
        "Precio_Promedio_USD_MWh"
    ] = (
        resumen_mensual_api["Ingreso_Total_USD"]
        / resumen_mensual_api["Energia_Total_MWh"]
        .replace(0, pd.NA)
    ).fillna(0)

    resumen_mensual_api["Fuente"] = "API"


# =========================================================
# CONSOLIDADO FINAL: EXCEL + API
# =========================================================

columnas_finales = [
    "Anio",
    "Mes",
    "Energia_PPA_MWh",
    "Energia_Spot_MWh",
    "Energia_Total_MWh",
    "Ingreso_PPA_USD",
    "Ingreso_Spot_USD",
    "Ingreso_Total_USD",
    "Precio_Promedio_USD_MWh",
    "Fuente"
]

fuentes_mensuales = [
    df_consolidado_excel[columnas_finales]
]

if not resumen_mensual_api.empty:
    fuentes_mensuales.append(
        resumen_mensual_api[columnas_finales]
    )

resumen_mensual_final = pd.concat(
    fuentes_mensuales,
    ignore_index=True
)

# Excel siempre tiene mayor prioridad que la API.
resumen_mensual_final["Prioridad"] = (
    resumen_mensual_final["Fuente"]
    .apply(
        lambda fuente: (
            2
            if str(fuente).lower() != "api"
            else 1
        )
    )
)

resumen_mensual_final = (
    resumen_mensual_final
    .sort_values("Prioridad")
    .drop_duplicates(
        subset=["Anio", "Mes"],
        keep="last"
    )
    .drop(columns=["Prioridad"])
    .sort_values(["Anio", "Mes"])
    .reset_index(drop=True)
)


# =========================================================
# FILTROS
# =========================================================

st.sidebar.header("Filtros")

anios_disponibles = sorted(
    resumen_mensual_final["Anio"]
    .unique()
    .tolist(),
    reverse=True
)

anio_seleccionado = st.sidebar.selectbox(
    "Año",
    options=anios_disponibles,
    index=0
)


# =========================================================
# INDICADORES ANUALES
# =========================================================

resumen_anio = resumen_mensual_final[
    resumen_mensual_final["Anio"]
    == anio_seleccionado
].copy()

energia_total_anual = (
    resumen_anio["Energia_Total_MWh"]
    .sum()
)

energia_ppa_anual = (
    resumen_anio["Energia_PPA_MWh"]
    .sum()
)

energia_spot_anual = (
    resumen_anio["Energia_Spot_MWh"]
    .sum()
)

ingreso_total_anual = (
    resumen_anio["Ingreso_Total_USD"]
    .sum()
)

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "Energía BESS acumulada",
    f"{energia_total_anual:,.1f} MWh"
)

col2.metric(
    "Energía PPA acumulada",
    f"{energia_ppa_anual:,.1f} MWh"
)

col3.metric(
    "Energía Spot acumulada",
    f"{energia_spot_anual:,.1f} MWh"
)

col4.metric(
    "Ingreso acumulado",
    f"USD {ingreso_total_anual:,.0f}"
)


# =========================================================
# GRAFICO DIARIO
# =========================================================

st.subheader("Distribución diaria de energía")

if resumen_diario_api.empty:
    st.info(
        "No hay información diaria disponible desde las API. "
        "Los meses consolidados manualmente solo contienen "
        "totales mensuales."
    )

else:
    meses_diarios = (
        resumen_diario_api.assign(
            Anio=resumen_diario_api[
                "Fecha_Dia"
            ].dt.year,
            Mes=resumen_diario_api[
                "Fecha_Dia"
            ].dt.month
        )
        [["Anio", "Mes"]]
        .drop_duplicates()
        .sort_values(
            ["Anio", "Mes"],
            ascending=False
        )
    )

    opciones_meses = [
        (int(fila.Anio), int(fila.Mes))
        for fila in meses_diarios.itertuples()
    ]

    seleccion_diaria = st.sidebar.selectbox(
        "Mes para gráfico diario",
        options=opciones_meses,
        format_func=lambda valor: (
            f"{MESES_COMPLETOS[valor[1]]} "
            f"{valor[0]}"
        )
    )

    anio_diario = seleccion_diaria[0]
    mes_diario = seleccion_diaria[1]

    diario_mes = resumen_diario_api[
        (
            resumen_diario_api[
                "Fecha_Dia"
            ].dt.year == anio_diario
        )
        & (
            resumen_diario_api[
                "Fecha_Dia"
            ].dt.month == mes_diario
        )
    ].copy()

    primer_dia_mes = pd.Timestamp(
        year=anio_diario,
        month=mes_diario,
        day=1
    )

    cantidad_dias = calendar.monthrange(
        anio_diario,
        mes_diario
    )[1]

    ultimo_dia_mes = pd.Timestamp(
        year=anio_diario,
        month=mes_diario,
        day=cantidad_dias
    )

    hoy = pd.Timestamp.now().normalize()

    if (
        anio_diario == hoy.year
        and mes_diario == hoy.month
    ):
        ultimo_dia_grafico = min(
            ultimo_dia_mes,
            hoy
        )
    else:
        ultimo_dia_grafico = ultimo_dia_mes

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

    ancho_barra = (
        0.40 * 24 * 60 * 60 * 1000
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
            width=ancho_barra,
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
            width=ancho_barra,
            hovertemplate=(
                "<b>%{x|%d-%m-%Y}</b><br>"
                "Spot: %{y:,.2f} MWh"
                "<extra></extra>"
            )
        )
    )

    fig_diario.update_layout(
        barmode="stack",
        title=(
            f"{MESES_COMPLETOS[mes_diario]} "
            f"{anio_diario}"
        ),
        xaxis_title="Día del mes",
        yaxis_title="Energía [MWh]",
        legend_title="Destino",
        hovermode="x unified",
        height=500,
        bargap=0.50,
        margin=dict(
            l=30,
            r=30,
            t=50,
            b=30
        )
    )

    fig_diario.update_xaxes(
        tickmode="linear",
        dtick=24 * 60 * 60 * 1000,
        tickformat="%d",
        range=[
            primer_dia_mes
            - pd.Timedelta(hours=12),
            ultimo_dia_grafico
            + pd.Timedelta(hours=12)
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
    "Mes": range(1, 13)
})

consolidado_anual = consolidado_anual.merge(
    resumen_anio,
    on="Mes",
    how="left"
)

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
    on="Mes",
    how="left"
)

columnas_cero = [
    "Energia_PPA_MWh",
    "Energia_Spot_MWh",
    "Energia_Total_MWh",
    "Ingreso_PPA_USD",
    "Ingreso_Spot_USD",
    "Ingreso_Total_USD"
]

for columna in columnas_cero:
    if columna not in consolidado_anual.columns:
        consolidado_anual[columna] = 0

consolidado_anual[columnas_cero] = (
    consolidado_anual[columnas_cero]
    .fillna(0)
)

consolidado_anual["Nombre_Mes"] = (
    consolidado_anual["Mes"]
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
# CUMPLIMIENTO ACUMULADO
# =========================================================

meses_con_resultados = (
    consolidado_anual["Energia_Total_MWh"] > 0
)

budget_generacion_acumulado = (
    consolidado_anual.loc[
        meses_con_resultados,
        "Budget_Generacion_MWh"
    ]
    .fillna(0)
    .sum()
)

budget_ppa_acumulado = (
    consolidado_anual.loc[
        meses_con_resultados,
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
# FUENTE DE CADA MES
# =========================================================

with st.expander(
    "Fuente utilizada por mes",
    expanded=False
):
    tabla_fuentes = resumen_anio[
        [
            "Anio",
            "Mes",
            "Fuente"
        ]
    ].copy()

    tabla_fuentes["Mes"] = (
        tabla_fuentes["Mes"]
        .map(MESES_COMPLETOS)
    )

    tabla_fuentes = tabla_fuentes.rename(
        columns={
            "Anio": "Año"
        }
    )

    st.dataframe(
        tabla_fuentes,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# INFORMACION FINAL
# =========================================================

if not df_api.empty:
    st.caption(
        "Última medición disponible desde la API: "
        f"{df_api['Fecha'].max():%d-%m-%Y %H:%M}"
    )

st.caption(
    "Dashboard procesado: "
    f"{datetime.now():%d-%m-%Y %H:%M:%S}"
)

st.caption(
    "Las consultas API y la caché se renuevan "
    "cada 12 horas."
)