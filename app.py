import calendar
import os
import time
import unicodedata
from datetime import datetime

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
from streamlit_autorefresh import st_autorefresh


# =========================================================
# CONFIGURACION GENERAL
# =========================================================

st.set_page_config(
    page_title="Dashboard de Operación",
    page_icon="⚡",
    layout="wide"
)

DOCE_HORAS_SEGUNDOS = 12 * 60 * 60
DOCE_HORAS_MS = 12 * 60 * 60 * 1000

st_autorefresh(
    interval=DOCE_HORAS_MS,
    key="actualizacion_dashboard_12h"
)

st.title("Dashboard de operación")
st.caption(
    "Seguimiento del BESS María Elena y generación mensual "
    "de proyectos en operación"
)


# =========================================================
# CREDENCIALES STREAMLIT
# =========================================================

try:
    API_KEY_PRMTE = st.secrets["API_KEY_PRMTE"]
    API_KEY_CMG = st.secrets["API_KEY_CMG"]

except KeyError:
    st.error(
        "No se encontraron las credenciales API en "
        "Streamlit Secrets."
    )
    st.info(
        "Configura API_KEY_PRMTE y API_KEY_CMG desde "
        "Manage app > Settings > Secrets."
    )
    st.stop()


# =========================================================
# CONFIGURACION OPERACIONAL
# =========================================================

ANIO_OPERACIONAL = 2026
FECHA_INICIO = f"{ANIO_OPERACIONAL}-01"

hoy = pd.Timestamp.now()

if hoy.year == ANIO_OPERACIONAL:
    FECHA_FIN = hoy.strftime("%Y-%m")
elif hoy.year > ANIO_OPERACIONAL:
    FECHA_FIN = f"{ANIO_OPERACIONAL}-12"
else:
    FECHA_FIN = f"{ANIO_OPERACIONAL}-01"

PRECIO_PPA = 83.29
TIPO_CMG = "PRELIMINAR"
BARRA_CMG_BESS = "M.ELENA_______220"

ARCHIVO_EXCEL = "budgets.xlsx"

HOJA_BUDGETS_BESS = "Budgets"
HOJA_CONSOLIDADO_BESS = "Consolidado"
HOJA_BUDGETS_PROYECTOS = "Budgets_Proyectos"
HOJA_CONSOLIDADO_PROYECTOS = "Consolidado_Proyectos"

URL_PRMTE = (
    "https://medidas.api.coordinador.cl/"
    "medidas-v2/measurement"
)

URL_CMG = (
    "https://sipub.api.coordinador.cl/"
    "costo-marginal-real/v4/findByDate"
)


# =========================================================
# CONFIGURACION DE PROYECTOS
# =========================================================

PROYECTOS = {
    "MARIA ELENA PFV": [
        {
            "mpid": "MARELENA_220_JT1_GSS",
            "canal": 3
        }
    ],
    "SAN PEDRO III": [
        {
            "mpid": "SOLRJAMA_220_JT1_RUC",
            "canal": 3
        },
        {
            "mpid": "SOLRJAMA_220_JT2_RUC",
            "canal": 3
        }
    ],
    "DOÑA CARMEN": [
        {
            "mpid": "CDNCARMN_220_J1_ECM",
            "canal": 3
        }
    ],
    "LA QUINTA": [
        {
            "mpid": "CABRERO_023_PMGD7_QTC",
            "canal": 3
        }
    ],
    "LA PERLA": [
        {
            "mpid": "LSANGLES_013_PMGD7_LAP",
            "canal": 3
        }
    ],
    "LA HUERTA": [
        {
            "mpid": "PARRONAL_013_PMGD2_HUE",
            "canal": 3
        }
    ],
    "CHACAICO": [
        {
            "mpid": "LSANGLES_013_PMGD8_CCC",
            "canal": 3
        }
    ],
    "SANCLEMENTE": [
        {
            "mpid": "SCLMENTE_013_PMGD4_CFL",
            "canal": 3
        }
    ],
    "TRILALEO": [
        {
            "mpid": "CHOLGUAN_015_PMGD6_YTL",
            "canal": 3
        }
    ],
    "COLLANCO": [
        {
            "mpid": "CNSTUCON_023_PMGD4_OAO",
            "canal": 3
        }
    ]
}

BESS_MPIDS = [
    "MARELENA_023_E1_GSS",
    "MARELENA_023_E7_GSS",
    "MARELENA_023_E8_GSS",
    "MARELENA_023_E11_GSS"
]

ORDEN_PROYECTOS = [
    "MARIA ELENA PFV",
    "SAN PEDRO III",
    "DOÑA CARMEN",
    "LA QUINTA",
    "LA PERLA",
    "LA HUERTA",
    "CHACAICO",
    "SANCLEMENTE",
    "TRILALEO",
    "COLLANCO"
]

MESES_CORTOS = {
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

def quitar_tildes(texto):
    return "".join(
        caracter
        for caracter in unicodedata.normalize(
            "NFD",
            str(texto)
        )
        if unicodedata.category(caracter) != "Mn"
    )


def normalizar_nombre_columna(nombre):
    nombre = quitar_tildes(nombre)
    nombre = nombre.strip()
    nombre = nombre.replace(" ", "_")
    nombre = nombre.replace("-", "_")

    while "__" in nombre:
        nombre = nombre.replace("__", "_")

    return nombre.lower()


def normalizar_columnas_excel(df):
    df = df.copy()

    equivalencias = {
        "ano": "Anio",
        "anio": "Anio",
        "mes": "Mes",
        "proyecto": "Proyecto",
        "budget_generacion_mwh":
            "Budget_Generacion_MWh",
        "budget_ppa_mwh":
            "Budget_PPA_MWh",
        "energia_ppa_mwh":
            "Energia_PPA_MWh",
        "energia_spot_mwh":
            "Energia_Spot_MWh",
        "ingreso_ppa_usd":
            "Ingreso_PPA_USD",
        "ingreso_spot_usd":
            "Ingreso_Spot_USD",
        "generacion_mwh":
            "Generacion_MWh"
    }

    nuevos_nombres = {}

    for columna in df.columns:
        normalizado = normalizar_nombre_columna(
            columna
        )

        nuevos_nombres[columna] = (
            equivalencias.get(
                normalizado,
                str(columna).strip()
            )
        )

    return df.rename(columns=nuevos_nombres)


def convertir_numerico(serie):
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(
            serie,
            errors="coerce"
        )

    texto = (
        serie.astype(str)
        .str.strip()
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    return pd.to_numeric(
        texto,
        errors="coerce"
    )


def parse_fecha_safe(serie):
    return pd.to_datetime(
        serie.astype(str).str[:19],
        errors="coerce"
    )


def generar_periodos(fecha_inicio, fecha_fin):
    inicio = datetime.strptime(
        fecha_inicio,
        "%Y-%m"
    )

    fin = datetime.strptime(
        fecha_fin,
        "%Y-%m"
    )

    periodos = []
    actual = inicio

    while actual <= fin:
        periodos.append(
            actual.strftime("%Y%m")
        )

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


def periodo_a_texto(periodo):
    anio = int(periodo[:4])
    mes = int(periodo[4:6])

    return f"{MESES_COMPLETOS[mes]} {anio}"


PERIODOS_DISPONIBLES = generar_periodos(
    FECHA_INICIO,
    FECHA_FIN
)


# =========================================================
# CONSULTA PRMTE
# =========================================================

def obtener_datos_prmte(
    periodo,
    mpid,
    canal
):
    params = {
        "channelId": canal,
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


def extraer_mediciones(datos, canal):
    registros = []
    nombre_canal = f"channel{canal}"

    if datos is None:
        return registros

    for bloque in datos:
        for medicion in bloque.get(
            "measurement",
            []
        ):
            registros.append({
                "Fecha": medicion.get("dateRange"),
                "Energia_kWh": medicion.get(
                    nombre_canal
                )
            })

    return registros


# =========================================================
# LECTURA DE EXCEL
# =========================================================

def leer_hoja_excel(
    hoja,
    columnas_requeridas
):
    if not os.path.exists(ARCHIVO_EXCEL):
        raise FileNotFoundError(
            f"No se encontró {ARCHIVO_EXCEL}."
        )

    df = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name=hoja,
        engine="openpyxl"
    )

    df = normalizar_columnas_excel(df)

    faltantes = [
        columna
        for columna in columnas_requeridas
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            f"Faltan columnas en {hoja}: "
            + ", ".join(faltantes)
        )

    return df[columnas_requeridas].copy()


@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def cargar_budgets_bess():
    requeridas = [
        "Anio",
        "Mes",
        "Budget_Generacion_MWh",
        "Budget_PPA_MWh"
    ]

    df = leer_hoja_excel(
        HOJA_BUDGETS_BESS,
        requeridas
    )

    for columna in requeridas:
        df[columna] = convertir_numerico(
            df[columna]
        )

    df = df.dropna(
        subset=["Anio", "Mes"]
    )

    df["Anio"] = df["Anio"].astype(int)
    df["Mes"] = df["Mes"].astype(int)

    if df.duplicated(
        subset=["Anio", "Mes"]
    ).any():
        raise ValueError(
            "Hay meses duplicados en la hoja Budgets."
        )

    return df


@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def cargar_consolidado_bess():
    requeridas = [
        "Anio",
        "Mes",
        "Energia_PPA_MWh",
        "Energia_Spot_MWh",
        "Ingreso_PPA_USD",
        "Ingreso_Spot_USD"
    ]

    df = leer_hoja_excel(
        HOJA_CONSOLIDADO_BESS,
        requeridas
    )

    for columna in requeridas:
        df[columna] = convertir_numerico(
            df[columna]
        )

    df = df.dropna(
        subset=["Anio", "Mes"]
    )

    df["Anio"] = df["Anio"].astype(int)
    df["Mes"] = df["Mes"].astype(int)

    if df.duplicated(
        subset=["Anio", "Mes"]
    ).any():
        raise ValueError(
            "Hay meses duplicados en Consolidado."
        )

    df["Energia_Total_MWh"] = (
        df["Energia_PPA_MWh"]
        + df["Energia_Spot_MWh"]
    )

    df["Ingreso_Total_USD"] = (
        df["Ingreso_PPA_USD"]
        + df["Ingreso_Spot_USD"]
    )

    df["Fuente"] = "Excel consolidado"

    return df


@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def cargar_budgets_proyectos():
    requeridas = [
        "Anio",
        "Mes",
        "Proyecto",
        "Budget_Generacion_MWh"
    ]

    df = leer_hoja_excel(
        HOJA_BUDGETS_PROYECTOS,
        requeridas
    )

    df["Anio"] = convertir_numerico(
        df["Anio"]
    )

    df["Mes"] = convertir_numerico(
        df["Mes"]
    )

    df["Budget_Generacion_MWh"] = (
        convertir_numerico(
            df["Budget_Generacion_MWh"]
        )
    )

    df["Proyecto"] = (
        df["Proyecto"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df = df.dropna(
        subset=[
            "Anio",
            "Mes",
            "Proyecto",
            "Budget_Generacion_MWh"
        ]
    )

    df["Anio"] = df["Anio"].astype(int)
    df["Mes"] = df["Mes"].astype(int)

    desconocidos = sorted(
        set(df["Proyecto"])
        - set(PROYECTOS.keys())
    )

    if desconocidos:
        raise ValueError(
            "Proyectos no reconocidos en "
            "Budgets_Proyectos: "
            + ", ".join(desconocidos)
        )

    if df.duplicated(
        subset=["Anio", "Mes", "Proyecto"]
    ).any():
        raise ValueError(
            "Hay budgets duplicados en "
            "Budgets_Proyectos."
        )

    return df


@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def cargar_consolidado_proyectos():
    requeridas = [
        "Anio",
        "Mes",
        "Proyecto",
        "Generacion_MWh"
    ]

    df = leer_hoja_excel(
        HOJA_CONSOLIDADO_PROYECTOS,
        requeridas
    )

    df["Anio"] = convertir_numerico(
        df["Anio"]
    )

    df["Mes"] = convertir_numerico(
        df["Mes"]
    )

    df["Generacion_MWh"] = convertir_numerico(
        df["Generacion_MWh"]
    )

    df["Proyecto"] = (
        df["Proyecto"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    df = df.dropna(
        subset=[
            "Anio",
            "Mes",
            "Proyecto",
            "Generacion_MWh"
        ]
    )

    df["Anio"] = df["Anio"].astype(int)
    df["Mes"] = df["Mes"].astype(int)

    desconocidos = sorted(
        set(df["Proyecto"])
        - set(PROYECTOS.keys())
    )

    if desconocidos:
        raise ValueError(
            "Proyectos no reconocidos en "
            "Consolidado_Proyectos: "
            + ", ".join(desconocidos)
        )

    if df.duplicated(
        subset=["Anio", "Mes", "Proyecto"]
    ).any():
        raise ValueError(
            "Hay registros duplicados en "
            "Consolidado_Proyectos."
        )

    if (df["Generacion_MWh"] < 0).any():
        raise ValueError(
            "Consolidado_Proyectos contiene "
            "generación negativa."
        )

    df["Fuente"] = "Excel consolidado"

    return df


# =========================================================
# PERIODOS PENDIENTES POR PROYECTO
# =========================================================

def obtener_pendientes_proyectos(
    df_consolidado
):
    pendientes = {}

    existentes = {
        (
            str(fila.Proyecto),
            f"{int(fila.Anio):04d}"
            f"{int(fila.Mes):02d}"
        )
        for fila in df_consolidado.itertuples()
    }

    for proyecto in PROYECTOS:
        pendientes[proyecto] = [
            periodo
            for periodo in PERIODOS_DISPONIBLES
            if (
                proyecto,
                periodo
            ) not in existentes
        ]

    return pendientes


# =========================================================
# DESCARGA DE PROYECTOS PENDIENTES
# =========================================================

@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def descargar_proyectos_pendientes(
    pendientes
):
    resultados = []
    advertencias = []

    for proyecto, periodos in pendientes.items():
        configuraciones = PROYECTOS[proyecto]

        for periodo in periodos:
            registros_periodo = []
            medidores_correctos = 0

            for configuracion in configuraciones:
                mpid = configuracion["mpid"]
                canal = configuracion["canal"]

                datos = obtener_datos_prmte(
                    periodo + "012345",
                    mpid,
                    canal
                )

                if datos is None:
                    advertencias.append(
                        f"{proyecto}, "
                        f"{periodo_a_texto(periodo)}: "
                        f"sin respuesta para {mpid}."
                    )
                    continue

                registros = extraer_mediciones(
                    datos,
                    canal
                )

                if registros:
                    medidores_correctos += 1
                    registros_periodo.extend(
                        registros
                    )
                else:
                    advertencias.append(
                        f"{proyecto}, "
                        f"{periodo_a_texto(periodo)}: "
                        f"sin mediciones para {mpid}."
                    )

            if (
                medidores_correctos
                != len(configuraciones)
            ):
                advertencias.append(
                    f"{proyecto}, "
                    f"{periodo_a_texto(periodo)} "
                    "no se incorporó porque faltan medidores."
                )
                continue

            if not registros_periodo:
                continue

            df_periodo = pd.DataFrame(
                registros_periodo
            )

            df_periodo["Fecha"] = (
                parse_fecha_safe(
                    df_periodo["Fecha"]
                )
            )

            df_periodo["Energia_kWh"] = (
                convertir_numerico(
                    df_periodo["Energia_kWh"]
                )
                .fillna(0)
            )

            df_periodo = (
                df_periodo
                .dropna(subset=["Fecha"])
                .groupby(
                    "Fecha",
                    as_index=False
                )
                .agg(
                    Energia_kWh=(
                        "Energia_kWh",
                        "sum"
                    )
                )
            )

            generacion_mwh = (
                df_periodo["Energia_kWh"]
                .abs()
                .sum()
                / 1000
            )

            resultados.append({
                "Anio": int(periodo[:4]),
                "Mes": int(periodo[4:6]),
                "Proyecto": proyecto,
                "Generacion_MWh":
                    generacion_mwh,
                "Fuente": "API"
            })

    return (
        pd.DataFrame(resultados),
        sorted(set(advertencias))
    )


# =========================================================
# CMG BESS
# =========================================================

def descargar_cmg_mes(periodo):
    anio = int(periodo[:4])
    mes = int(periodo[4:6])

    ultimo_dia = calendar.monthrange(
        anio,
        mes
    )[1]

    inicio = pd.Timestamp(
        year=anio,
        month=mes,
        day=1
    )

    fin = pd.Timestamp(
        year=anio,
        month=mes,
        day=ultimo_dia
    )

    hoy_normalizado = pd.Timestamp.now().normalize()

    if (
        anio == hoy_normalizado.year
        and mes == hoy_normalizado.month
    ):
        fin = min(fin, hoy_normalizado)

    registros = []
    pagina = 0

    while True:
        params = {
            "startDate": inicio.strftime(
                "%Y-%m-%d"
            ),
            "endDate": fin.strftime(
                "%Y-%m-%d"
            ),
            "page": pagina,
            "limit": 5000,
            "type": TIPO_CMG,
            "bar_transf": BARRA_CMG_BESS,
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

            datos = respuesta.json().get(
                "data",
                []
            )

        except (
            requests.RequestException,
            ValueError
        ):
            return None

        if not datos:
            break

        registros.extend(datos)

        if len(datos) < 5000:
            break

        pagina += 1

    if not registros:
        return None

    df = pd.DataFrame(registros)

    requeridas = [
        "fecha",
        "hra",
        "min",
        "cmg_usd_mwh_"
    ]

    if not all(
        columna in df.columns
        for columna in requeridas
    ):
        return None

    df["hra"] = convertir_numerico(
        df["hra"]
    ).fillna(0)

    df["min"] = convertir_numerico(
        df["min"]
    ).fillna(0)

    df["CMG_USD_MWh"] = convertir_numerico(
        df["cmg_usd_mwh_"]
    ).fillna(0)

    df["Fecha"] = (
        pd.to_datetime(
            df["fecha"],
            errors="coerce"
        )
        + pd.to_timedelta(
            df["hra"],
            unit="h"
        )
        + pd.to_timedelta(
            df["min"],
            unit="m"
        )
    )

    return (
        df[
            ["Fecha", "CMG_USD_MWh"]
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


# =========================================================
# PROCESAMIENTO BESS API
# =========================================================

@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def procesar_bess_api(periodos):
    resumenes = []
    advertencias = []

    for periodo in periodos:
        registros_totales = []
        medidores_correctos = 0

        for mpid in BESS_MPIDS:
            datos = obtener_datos_prmte(
                periodo + "012345",
                mpid,
                3
            )

            if datos is None:
                advertencias.append(
                    f"BESS, {periodo_a_texto(periodo)}: "
                    f"sin respuesta para {mpid}."
                )
                continue

            registros = extraer_mediciones(
                datos,
                3
            )

            if registros:
                medidores_correctos += 1
                registros_totales.extend(
                    registros
                )

        if medidores_correctos != len(BESS_MPIDS):
            advertencias.append(
                f"BESS, {periodo_a_texto(periodo)} "
                "no se incorporó porque faltan medidores."
            )
            continue

        if not registros_totales:
            continue

        df = pd.DataFrame(registros_totales)

        df["Fecha"] = parse_fecha_safe(
            df["Fecha"]
        )

        df["Energia_kWh"] = (
            convertir_numerico(
                df["Energia_kWh"]
            )
            .fillna(0)
        )

        df = (
            df
            .dropna(subset=["Fecha"])
            .groupby("Fecha", as_index=False)
            .agg(
                Energia_kWh=(
                    "Energia_kWh",
                    "sum"
                )
            )
        )

        df_cmg = descargar_cmg_mes(periodo)

        if df_cmg is None or df_cmg.empty:
            advertencias.append(
                f"BESS, {periodo_a_texto(periodo)}: "
                "CMG no disponible."
            )
            continue

        df = df.merge(
            df_cmg,
            on="Fecha",
            how="left"
        )

        df["CMG_USD_MWh"] = (
            df["CMG_USD_MWh"]
            .fillna(0)
        )

        df["Energia_MWh"] = (
            df["Energia_kWh"].abs()
            / 1000
        )

        df["Tipo"] = "PPA"

        df.loc[
            (
                df["Fecha"].dt.hour >= 6
            )
            & (
                df["Fecha"].dt.hour < 21
            ),
            "Tipo"
        ] = "SPOT"

        df["Precio"] = PRECIO_PPA

        df.loc[
            df["Tipo"] == "SPOT",
            "Precio"
        ] = df["CMG_USD_MWh"]

        df["Ingreso_USD"] = (
            df["Energia_MWh"]
            * df["Precio"]
        )

        energia_ppa = df.loc[
            df["Tipo"] == "PPA",
            "Energia_MWh"
        ].sum()

        energia_spot = df.loc[
            df["Tipo"] == "SPOT",
            "Energia_MWh"
        ].sum()

        ingreso_ppa = df.loc[
            df["Tipo"] == "PPA",
            "Ingreso_USD"
        ].sum()

        ingreso_spot = df.loc[
            df["Tipo"] == "SPOT",
            "Ingreso_USD"
        ].sum()

        resumenes.append({
            "Anio": int(periodo[:4]),
            "Mes": int(periodo[4:6]),
            "Energia_PPA_MWh":
                energia_ppa,
            "Energia_Spot_MWh":
                energia_spot,
            "Energia_Total_MWh":
                energia_ppa + energia_spot,
            "Ingreso_PPA_USD":
                ingreso_ppa,
            "Ingreso_Spot_USD":
                ingreso_spot,
            "Ingreso_Total_USD":
                ingreso_ppa + ingreso_spot,
            "Fuente": "API"
        })

    return (
        pd.DataFrame(resumenes),
        sorted(set(advertencias))
    )


# =========================================================
# VALIDAR Y CARGAR EXCEL
# =========================================================

try:
    df_budgets_bess = cargar_budgets_bess()
    df_consolidado_bess = (
        cargar_consolidado_bess()
    )
    df_budgets_proyectos = (
        cargar_budgets_proyectos()
    )
    df_consolidado_proyectos = (
        cargar_consolidado_proyectos()
    )

except Exception as error:
    st.error(
        "No fue posible validar budgets.xlsx."
    )
    st.info(str(error))
    st.stop()


# =========================================================
# CONSULTAR PROYECTOS PENDIENTES
# =========================================================

pendientes_proyectos = (
    obtener_pendientes_proyectos(
        df_consolidado_proyectos
    )
)

with st.spinner(
    "Consultando meses pendientes de proyectos..."
):
    (
        df_proyectos_api,
        advertencias_proyectos
    ) = descargar_proyectos_pendientes(
        pendientes_proyectos
    )

fuentes_proyectos = [
    df_consolidado_proyectos
]

if not df_proyectos_api.empty:
    fuentes_proyectos.append(
        df_proyectos_api
    )

df_proyectos_final = pd.concat(
    fuentes_proyectos,
    ignore_index=True
)

df_proyectos_final["Prioridad"] = (
    df_proyectos_final["Fuente"]
    .map({
        "API": 1,
        "Excel consolidado": 2
    })
    .fillna(2)
)

df_proyectos_final = (
    df_proyectos_final
    .sort_values("Prioridad")
    .drop_duplicates(
        subset=[
            "Anio",
            "Mes",
            "Proyecto"
        ],
        keep="last"
    )
    .drop(columns=["Prioridad"])
    .sort_values(
        ["Proyecto", "Anio", "Mes"]
    )
)


# =========================================================
# CONSULTAR BESS PENDIENTE
# =========================================================

periodos_bess_excel = {
    f"{int(fila.Anio):04d}"
    f"{int(fila.Mes):02d}"
    for fila in df_consolidado_bess.itertuples()
}

periodos_bess_pendientes = [
    periodo
    for periodo in PERIODOS_DISPONIBLES
    if periodo not in periodos_bess_excel
]

with st.spinner(
    "Consultando meses pendientes del BESS..."
):
    (
        df_bess_api,
        advertencias_bess
    ) = procesar_bess_api(
        periodos_bess_pendientes
    )

fuentes_bess = [
    df_consolidado_bess
]

if not df_bess_api.empty:
    fuentes_bess.append(
        df_bess_api
    )

df_bess_final = pd.concat(
    fuentes_bess,
    ignore_index=True
)

df_bess_final["Prioridad"] = (
    df_bess_final["Fuente"]
    .map({
        "API": 1,
        "Excel consolidado": 2
    })
    .fillna(2)
)

df_bess_final = (
    df_bess_final
    .sort_values("Prioridad")
    .drop_duplicates(
        subset=["Anio", "Mes"],
        keep="last"
    )
    .drop(columns=["Prioridad"])
    .sort_values(["Anio", "Mes"])
)


# =========================================================
# ADVERTENCIAS
# =========================================================

todas_advertencias = sorted(
    set(
        advertencias_proyectos
        + advertencias_bess
    )
)

if todas_advertencias:
    st.warning(
        "Algunas mediciones no estuvieron disponibles. "
        "El dashboard continuará con la información válida."
    )

    with st.expander(
        "Ver advertencias de las API"
    ):
        for advertencia in todas_advertencias:
            st.write(f"- {advertencia}")


# =========================================================
# FILTRO GENERAL
# =========================================================

st.sidebar.header("Filtros")

anio_seleccionado = st.sidebar.selectbox(
    "Año",
    options=[ANIO_OPERACIONAL],
    index=0
)


# =========================================================
# SECCION BESS PRIMERO
# =========================================================

st.header("BESS María Elena")

bess_anio = df_bess_final[
    df_bess_final["Anio"]
    == anio_seleccionado
].copy()

if bess_anio.empty:
    st.info(
        "No hay información disponible del BESS."
    )

else:
    meses_bess = sorted(
        bess_anio["Mes"]
        .astype(int)
        .unique()
        .tolist(),
        reverse=True
    )

    mes_bess = st.selectbox(
        "Mes para indicadores BESS",
        options=meses_bess,
        format_func=lambda mes: (
            MESES_COMPLETOS[mes]
        )
    )

    fila_bess = bess_anio[
        bess_anio["Mes"] == mes_bess
    ].iloc[0]

    col_b1, col_b2, col_b3, col_b4 = (
        st.columns(4)
    )

    col_b1.metric(
        "Inyección total del mes",
        (
            f"{fila_bess['Energia_Total_MWh']:,.1f} "
            "MWh"
        )
    )

    col_b2.metric(
        "Energía PPA del mes",
        (
            f"{fila_bess['Energia_PPA_MWh']:,.1f} "
            "MWh"
        )
    )

    col_b3.metric(
        "Energía Spot del mes",
        (
            f"{fila_bess['Energia_Spot_MWh']:,.1f} "
            "MWh"
        )
    )

    col_b4.metric(
        "Ingreso del mes",
        (
            f"USD "
            f"{fila_bess['Ingreso_Total_USD']:,.0f}"
        )
    )

    st.caption(
        f"Fuente: {fila_bess['Fuente']}. "
        f"Precio PPA: USD {PRECIO_PPA:,.2f}/MWh."
    )


# =========================================================
# GRAFICO BESS
# =========================================================

consolidado_bess = pd.DataFrame({
    "Mes": range(1, 13)
})

consolidado_bess = consolidado_bess.merge(
    bess_anio,
    on="Mes",
    how="left"
)

budgets_bess_anio = df_budgets_bess[
    df_budgets_bess["Anio"]
    == anio_seleccionado
]

consolidado_bess = consolidado_bess.merge(
    budgets_bess_anio[
        [
            "Mes",
            "Budget_Generacion_MWh",
            "Budget_PPA_MWh"
        ]
    ],
    on="Mes",
    how="left"
)

for columna in [
    "Energia_PPA_MWh",
    "Energia_Spot_MWh",
    "Energia_Total_MWh",
    "Ingreso_Total_USD"
]:
    if columna not in consolidado_bess.columns:
        consolidado_bess[columna] = 0

    consolidado_bess[columna] = (
        consolidado_bess[columna]
        .fillna(0)
    )

consolidado_bess["Nombre_Mes"] = (
    consolidado_bess["Mes"]
    .map(MESES_CORTOS)
)

fig_bess = go.Figure()

# PPA real azul
fig_bess.add_trace(
    go.Bar(
        x=consolidado_bess["Nombre_Mes"],
        y=consolidado_bess[
            "Energia_PPA_MWh"
        ],
        name="PPA real",
        marker_color="#1565C0",
        width=0.58,
        hovertemplate=(
            "<b>%{x}</b><br>"
            "PPA real: %{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

# Spot real celeste
fig_bess.add_trace(
    go.Bar(
        x=consolidado_bess["Nombre_Mes"],
        y=consolidado_bess[
            "Energia_Spot_MWh"
        ],
        name="Spot real",
        marker_color="#65BDEB",
        width=0.58,
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Spot real: %{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

# Budget generación azul oscuro
fig_bess.add_trace(
    go.Scatter(
        x=consolidado_bess["Nombre_Mes"],
        y=consolidado_bess[
            "Budget_Generacion_MWh"
        ],
        name="Budget generación",
        mode="lines+markers",
        line=dict(
            color="#0B3D91",
            width=3
        ),
        marker=dict(
            color="#0B3D91",
            size=8
        )
    )
)

# Budget PPA verde
fig_bess.add_trace(
    go.Scatter(
        x=consolidado_bess["Nombre_Mes"],
        y=consolidado_bess[
            "Budget_PPA_MWh"
        ],
        name="Budget PPA",
        mode="lines+markers",
        line=dict(
            color="#2EAD5B",
            width=3
        ),
        marker=dict(
            color="#2EAD5B",
            size=8
        )
    )
)

fig_bess.update_layout(
    title="Consolidado anual BESS",
    barmode="stack",
    xaxis_title="Mes",
    yaxis_title="Energía [MWh]",
    hovermode="x unified",
    height=560,
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="left",
        x=0
    )
)

fig_bess.update_yaxes(
    rangemode="tozero",
    gridcolor="rgba(140,140,140,0.22)"
)

st.plotly_chart(
    fig_bess,
    use_container_width=True
)


# =========================================================
# ACUMULADOS BESS
# =========================================================

with st.expander(
    "Ver acumulados del BESS",
    expanded=False
):
    col_a1, col_a2, col_a3, col_a4 = (
        st.columns(4)
    )

    col_a1.metric(
        "Inyección acumulada",
        (
            f"{bess_anio['Energia_Total_MWh'].sum():,.1f} "
            "MWh"
        )
    )

    col_a2.metric(
        "PPA acumulado",
        (
            f"{bess_anio['Energia_PPA_MWh'].sum():,.1f} "
            "MWh"
        )
    )

    col_a3.metric(
        "Spot acumulado",
        (
            f"{bess_anio['Energia_Spot_MWh'].sum():,.1f} "
            "MWh"
        )
    )

    col_a4.metric(
        "Ingreso acumulado",
        (
            f"USD "
            f"{bess_anio['Ingreso_Total_USD'].sum():,.0f}"
        )
    )


# =========================================================
# PROYECTOS
# =========================================================

st.header("Generación mensual de proyectos")

st.caption(
    "Cada gráfico utiliza el Excel para los meses "
    "consolidados y la API para los meses pendientes."
)

proyectos_anio = df_proyectos_final[
    df_proyectos_final["Anio"]
    == anio_seleccionado
].copy()

budgets_proyectos_anio = (
    df_budgets_proyectos[
        df_budgets_proyectos["Anio"]
        == anio_seleccionado
    ]
    .copy()
)


# =========================================================
# GRAFICOS PEQUEÑOS DE CADA PROYECTO
# =========================================================

columnas_graficos = st.columns(2)

for indice, proyecto in enumerate(
    ORDEN_PROYECTOS
):
    datos_proyecto = proyectos_anio[
        proyectos_anio["Proyecto"]
        == proyecto
    ][
        [
            "Mes",
            "Generacion_MWh",
            "Fuente"
        ]
    ].copy()

    budget_proyecto = (
        budgets_proyectos_anio[
            budgets_proyectos_anio[
                "Proyecto"
            ] == proyecto
        ][
            [
                "Mes",
                "Budget_Generacion_MWh"
            ]
        ]
        .copy()
    )

    consolidado = pd.DataFrame({
        "Mes": range(1, 13)
    })

    consolidado = (
        consolidado
        .merge(
            datos_proyecto,
            on="Mes",
            how="left"
        )
        .merge(
            budget_proyecto,
            on="Mes",
            how="left"
        )
    )

    consolidado["Generacion_MWh"] = (
        consolidado["Generacion_MWh"]
        .fillna(0)
    )

    consolidado["Nombre_Mes"] = (
        consolidado["Mes"]
        .map(MESES_CORTOS)
    )

    fig = go.Figure()

    fig.add_trace(
        go.Bar(
            x=consolidado["Nombre_Mes"],
            y=consolidado["Generacion_MWh"],
            name="Generación",
            marker_color="#2176C7",
            width=0.52,
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Generación: %{y:,.1f} MWh"
                "<extra></extra>"
            )
        )
    )

    fig.add_trace(
        go.Scatter(
            x=consolidado["Nombre_Mes"],
            y=consolidado[
                "Budget_Generacion_MWh"
            ],
            name="Budget",
            mode="lines+markers",
            line=dict(
                color="#F05A28",
                width=2
            ),
            marker=dict(
                size=5
            ),
            hovertemplate=(
                "<b>%{x}</b><br>"
                "Budget: %{y:,.1f} MWh"
                "<extra></extra>"
            )
        )
    )

    fig.update_layout(
        title=dict(
            text=proyecto,
            font=dict(size=16)
        ),
        height=330,
        margin=dict(
            l=35,
            r=15,
            t=50,
            b=35
        ),
        showlegend=(
            indice == 0
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0
        ),
        hovermode="x unified",
        xaxis_title=None,
        yaxis_title="MWh",
        bargap=0.40
    )

    fig.update_xaxes(
        tickfont=dict(size=10),
        showgrid=False
    )

    fig.update_yaxes(
        rangemode="tozero",
        tickfont=dict(size=10),
        gridcolor="rgba(140,140,140,0.18)"
    )

    with columnas_graficos[
        indice % 2
    ]:
        st.plotly_chart(
            fig,
            use_container_width=True,
            key=f"grafico_{proyecto}"
        )


# =========================================================
# GENERACION TOTAL VS BUDGET TOTAL
# =========================================================

st.subheader(
    "Generación mensual total de proyectos"
)

generacion_total = (
    proyectos_anio
    .groupby(
        "Mes",
        as_index=False
    )
    .agg(
        Generacion_Total_MWh=(
            "Generacion_MWh",
            "sum"
        )
    )
)

budget_total = (
    budgets_proyectos_anio
    .groupby(
        "Mes",
        as_index=False
    )
    .agg(
        Budget_Total_MWh=(
            "Budget_Generacion_MWh",
            "sum"
        )
    )
)

total_mensual = pd.DataFrame({
    "Mes": range(1, 13)
})

total_mensual = (
    total_mensual
    .merge(
        generacion_total,
        on="Mes",
        how="left"
    )
    .merge(
        budget_total,
        on="Mes",
        how="left"
    )
)

total_mensual[
    "Generacion_Total_MWh"
] = (
    total_mensual[
        "Generacion_Total_MWh"
    ]
    .fillna(0)
)

total_mensual["Nombre_Mes"] = (
    total_mensual["Mes"]
    .map(MESES_CORTOS)
)

fig_total = go.Figure()

fig_total.add_trace(
    go.Bar(
        x=total_mensual["Nombre_Mes"],
        y=total_mensual[
            "Generacion_Total_MWh"
        ],
        name="Generación real total",
        marker_color="#1565C0",
        width=0.38,
        offsetgroup="real",
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Generación total: %{y:,.1f} MWh"
            "<extra></extra>"
        )
    )
)

fig_total.add_trace(
    go.Bar(
        x=total_mensual["Nombre_Mes"],
        y=total_mensual[
            "Budget_Total_MWh"
        ],
        name="Budget total",
        marker_color="#E58A21",
        width=0.38,
        offsetgroup="budget",
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Budget total: %{y:,.1f} MWh"
            "<extra></extra>"
        )
    )
)

fig_total.update_layout(
    barmode="group",
    xaxis_title="Mes",
    yaxis_title="Generación [MWh]",
    height=520,
    hovermode="x unified",
    legend=dict(
        orientation="h",
        yanchor="bottom",
        y=1.02,
        xanchor="left",
        x=0
    )
)

fig_total.update_yaxes(
    rangemode="tozero",
    gridcolor="rgba(140,140,140,0.22)"
)

st.plotly_chart(
    fig_total,
    use_container_width=True
)


# =========================================================
# INDICADORES TOTALES DE PROYECTOS
# =========================================================

meses_con_generacion = total_mensual[
    total_mensual["Generacion_Total_MWh"] > 0
]

if not meses_con_generacion.empty:
    ultimo_mes = int(
        meses_con_generacion["Mes"].max()
    )

    fila_mes = total_mensual[
        total_mensual["Mes"] == ultimo_mes
    ].iloc[0]

    generacion_acumulada = (
        total_mensual.loc[
            total_mensual["Mes"] <= ultimo_mes,
            "Generacion_Total_MWh"
        ]
        .sum()
    )

    budget_acumulado = (
        total_mensual.loc[
            total_mensual["Mes"] <= ultimo_mes,
            "Budget_Total_MWh"
        ]
        .fillna(0)
        .sum()
    )

    diferencia_acumulada = (
        generacion_acumulada
        - budget_acumulado
    )

    col_t1, col_t2, col_t3, col_t4 = (
        st.columns(4)
    )

    col_t1.metric(
        f"Generación total {MESES_COMPLETOS[ultimo_mes]}",
        (
            f"{fila_mes['Generacion_Total_MWh']:,.1f} "
            "MWh"
        )
    )

    col_t2.metric(
        f"Budget total {MESES_COMPLETOS[ultimo_mes]}",
        (
            f"{fila_mes['Budget_Total_MWh']:,.1f} "
            "MWh"
            if pd.notna(
                fila_mes["Budget_Total_MWh"]
            )
            else "Sin budget"
        )
    )

    col_t3.metric(
        "Generación acumulada",
        f"{generacion_acumulada:,.1f} MWh"
    )

    col_t4.metric(
        "Diferencia acumulada vs budget",
        f"{diferencia_acumulada:,.1f} MWh"
    )


# =========================================================
# FUENTES UTILIZADAS
# =========================================================

with st.expander(
    "Ver fuentes utilizadas por proyecto",
    expanded=False
):
    tabla_fuentes = (
        proyectos_anio[
            [
                "Proyecto",
                "Mes",
                "Fuente"
            ]
        ]
        .copy()
        .sort_values(
            ["Proyecto", "Mes"]
        )
    )

    tabla_fuentes["Mes"] = (
        tabla_fuentes["Mes"]
        .map(MESES_COMPLETOS)
    )

    st.dataframe(
        tabla_fuentes,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# INFORMACION FINAL
# =========================================================

st.caption(
    "Dashboard procesado: "
    f"{datetime.now():%d-%m-%Y %H:%M:%S}"
)

st.caption(
    "Las consultas a las API y la caché se renuevan "
    "cada 12 horas."
)