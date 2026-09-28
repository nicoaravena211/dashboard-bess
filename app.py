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
DOCE_HORAS_MILISEGUNDOS = 12 * 60 * 60 * 1000

st_autorefresh(
    interval=DOCE_HORAS_MILISEGUNDOS,
    key="actualizacion_dashboard_12h"
)

st.title("Dashboard de operación")
st.caption(
    "Seguimiento mensual de generación de proyectos "
    "y operación del BESS María Elena"
)


# =========================================================
# CREDENCIALES
# =========================================================

try:
    API_KEY_PRMTE = st.secrets["API_KEY_PRMTE"]
    API_KEY_CMG = st.secrets["API_KEY_CMG"]

except KeyError:
    st.error(
        "No se encontraron las credenciales API "
        "en los Secrets de Streamlit."
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


# Precio PPA actualizado
PRECIO_PPA = 83.29

TIPO_CMG = "PRELIMINAR"
BARRA_CMG_BESS = "M.ELENA_______220"

ARCHIVO_EXCEL = "budgets.xlsx"

HOJA_BUDGETS_BESS = "Budgets"
HOJA_CONSOLIDADO_BESS = "Consolidado"
HOJA_BUDGETS_PROYECTOS = "Budgets_Proyectos"

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

PROYECTOS_GENERACION = {
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
            "Ingreso_Spot_USD"
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

    return (
        f"{MESES_COMPLETOS[mes]} "
        f"{anio}"
    )


PERIODOS_ANIO = generar_periodos(
    FECHA_INICIO,
    FECHA_FIN
)


# =========================================================
# CONSULTA GENERAL PRMTE
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


def extraer_mediciones(
    datos,
    canal
):
    registros = []
    nombre_canal = f"channel{canal}"

    if datos is None:
        return registros

    for bloque in datos:
        mediciones = bloque.get(
            "measurement",
            []
        )

        for medicion in mediciones:
            registros.append({
                "Fecha": medicion.get(
                    "dateRange"
                ),
                "Energia_kWh": medicion.get(
                    nombre_canal
                )
            })

    return registros


# =========================================================
# DESCARGA DE GENERACION DE OTROS PROYECTOS
# =========================================================

@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def descargar_generacion_proyectos(periodos):
    resultados = []
    advertencias = []

    for proyecto, configuraciones in (
        PROYECTOS_GENERACION.items()
    ):
        registros_proyecto = []

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
                        f"la API no respondió para {mpid}."
                    )
                    continue

                registros_mpid = extraer_mediciones(
                    datos,
                    canal
                )

                if registros_mpid:
                    medidores_correctos += 1
                    registros_periodo.extend(
                        registros_mpid
                    )
                else:
                    advertencias.append(
                        f"{proyecto}, "
                        f"{periodo_a_texto(periodo)}: "
                        f"sin mediciones para {mpid}."
                    )

            if (
                medidores_correctos
                == len(configuraciones)
            ):
                for registro in registros_periodo:
                    registro["Proyecto"] = proyecto
                    registro["Periodo"] = periodo

                registros_proyecto.extend(
                    registros_periodo
                )
            else:
                advertencias.append(
                    f"{proyecto}, "
                    f"{periodo_a_texto(periodo)} "
                    "no se incorporó porque faltan "
                    "medidores."
                )

        if registros_proyecto:
            df_proyecto = pd.DataFrame(
                registros_proyecto
            )

            df_proyecto["Fecha"] = (
                parse_fecha_safe(
                    df_proyecto["Fecha"]
                )
            )

            df_proyecto["Energia_kWh"] = (
                convertir_numerico(
                    df_proyecto["Energia_kWh"]
                )
                .fillna(0)
            )

            df_proyecto = (
                df_proyecto.dropna(
                    subset=["Fecha"]
                )
            )

            # Suma todos los MPID del proyecto por intervalo.
            df_proyecto = (
                df_proyecto
                .groupby(
                    [
                        "Proyecto",
                        "Fecha"
                    ],
                    as_index=False
                )
                .agg(
                    Energia_kWh=(
                        "Energia_kWh",
                        "sum"
                    )
                )
            )

            resultados.append(df_proyecto)

    if not resultados:
        return (
            pd.DataFrame(
                columns=[
                    "Proyecto",
                    "Fecha",
                    "Energia_kWh"
                ]
            ),
            sorted(set(advertencias))
        )

    df_generacion = pd.concat(
        resultados,
        ignore_index=True
    )

    # La API entrega kWh.
    # Para generación se usan valores absolutos.
    df_generacion["Generacion_MWh"] = (
        df_generacion["Energia_kWh"]
        .abs()
        / 1000
    )

    df_generacion["Anio"] = (
        df_generacion["Fecha"].dt.year
    )

    df_generacion["Mes"] = (
        df_generacion["Fecha"].dt.month
    )

    resumen_mensual = (
        df_generacion
        .groupby(
            [
                "Anio",
                "Mes",
                "Proyecto"
            ],
            as_index=False
        )
        .agg(
            Generacion_MWh=(
                "Generacion_MWh",
                "sum"
            )
        )
        .sort_values(
            [
                "Proyecto",
                "Anio",
                "Mes"
            ]
        )
    )

    return (
        resumen_mensual,
        sorted(set(advertencias))
    )


# =========================================================
# LECTURA DE BUDGETS DE PROYECTOS
# =========================================================

@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def cargar_budgets_proyectos():
    df = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name=HOJA_BUDGETS_PROYECTOS,
        engine="openpyxl"
    )

    df = normalizar_columnas_excel(df)

    requeridas = [
        "Anio",
        "Mes",
        "Proyecto",
        "Budget_Generacion_MWh"
    ]

    faltantes = [
        columna
        for columna in requeridas
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas en Budgets_Proyectos: "
            + ", ".join(faltantes)
        )

    df = df[requeridas].copy()

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
            "Proyecto"
        ]
    )

    df["Anio"] = df["Anio"].astype(int)
    df["Mes"] = df["Mes"].astype(int)

    if not df["Mes"].between(1, 12).all():
        raise ValueError(
            "Budgets_Proyectos contiene meses "
            "fuera del rango 1 a 12."
        )

    if df[
        "Budget_Generacion_MWh"
    ].isna().any():
        raise ValueError(
            "Budgets_Proyectos contiene budgets "
            "vacíos o no numéricos."
        )

    nombres_validos = set(
        PROYECTOS_GENERACION.keys()
    )

    nombres_excel = set(
        df["Proyecto"].unique()
    )

    desconocidos = sorted(
        nombres_excel - nombres_validos
    )

    if desconocidos:
        raise ValueError(
            "Hay proyectos no reconocidos en "
            "Budgets_Proyectos: "
            + ", ".join(desconocidos)
        )

    duplicados = df.duplicated(
        subset=[
            "Anio",
            "Mes",
            "Proyecto"
        ],
        keep=False
    )

    if duplicados.any():
        raise ValueError(
            "Existen budgets duplicados para el mismo "
            "año, mes y proyecto."
        )

    return df


# =========================================================
# BUDGETS Y CONSOLIDADO DEL BESS
# =========================================================

@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def cargar_budgets_bess():
    df = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name=HOJA_BUDGETS_BESS,
        engine="openpyxl"
    )

    df = normalizar_columnas_excel(df)

    requeridas = [
        "Anio",
        "Mes",
        "Budget_Generacion_MWh",
        "Budget_PPA_MWh"
    ]

    faltantes = [
        columna
        for columna in requeridas
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas en Budgets: "
            + ", ".join(faltantes)
        )

    df = df[requeridas].copy()

    for columna in requeridas:
        df[columna] = convertir_numerico(
            df[columna]
        )

    df = df.dropna(
        subset=["Anio", "Mes"]
    )

    df["Anio"] = df["Anio"].astype(int)
    df["Mes"] = df["Mes"].astype(int)

    return df


@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def cargar_consolidado_bess():
    df = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name=HOJA_CONSOLIDADO_BESS,
        engine="openpyxl"
    )

    df = normalizar_columnas_excel(df)

    requeridas = [
        "Anio",
        "Mes",
        "Energia_PPA_MWh",
        "Energia_Spot_MWh",
        "Ingreso_PPA_USD",
        "Ingreso_Spot_USD"
    ]

    faltantes = [
        columna
        for columna in requeridas
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas en Consolidado: "
            + ", ".join(faltantes)
        )

    df = df[requeridas].copy()

    for columna in requeridas:
        df[columna] = convertir_numerico(
            df[columna]
        )

    df = df.dropna(
        subset=["Anio", "Mes"]
    )

    df["Anio"] = df["Anio"].astype(int)
    df["Mes"] = df["Mes"].astype(int)

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


# =========================================================
# CMG PARA EL BESS
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
            "startDate": inicio.strftime("%Y-%m-%d"),
            "endDate": fin.strftime("%Y-%m-%d"),
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
        return pd.DataFrame(
            columns=["Fecha", "CMG_USD_MWh"]
        )

    df = pd.DataFrame(registros)

    necesarias = [
        "fecha",
        "hra",
        "min",
        "cmg_usd_mwh_"
    ]

    if not all(
        columna in df.columns
        for columna in necesarias
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
# DESCARGA Y PROCESAMIENTO DEL BESS PENDIENTE
# =========================================================

@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def procesar_bess_api(periodos):
    resumenes_mensuales = []
    resumenes_diarios = []
    advertencias = []

    for periodo in periodos:
        registros_bess = []
        mpids_correctos = 0

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
                mpids_correctos += 1
                registros_bess.extend(registros)

        if mpids_correctos != len(BESS_MPIDS):
            advertencias.append(
                f"BESS, {periodo_a_texto(periodo)} "
                "no se procesó porque faltan medidores."
            )
            continue

        if not registros_bess:
            continue

        df_bess = pd.DataFrame(registros_bess)

        df_bess["Fecha"] = parse_fecha_safe(
            df_bess["Fecha"]
        )

        df_bess["Energia_kWh"] = (
            convertir_numerico(
                df_bess["Energia_kWh"]
            )
            .fillna(0)
        )

        df_bess = (
            df_bess
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

        df_bess = df_bess.merge(
            df_cmg,
            on="Fecha",
            how="left"
        )

        df_bess["CMG_USD_MWh"] = (
            df_bess["CMG_USD_MWh"]
            .fillna(0)
        )

        df_bess["Energia_MWh"] = (
            df_bess["Energia_kWh"].abs()
            / 1000
        )

        df_bess["Tipo_Contrato"] = "PPA"

        df_bess.loc[
            (
                df_bess["Fecha"].dt.hour >= 6
            )
            & (
                df_bess["Fecha"].dt.hour < 21
            ),
            "Tipo_Contrato"
        ] = "SPOT"

        df_bess["Precio_Aplicado"] = (
            PRECIO_PPA
        )

        df_bess.loc[
            df_bess["Tipo_Contrato"] == "SPOT",
            "Precio_Aplicado"
        ] = df_bess["CMG_USD_MWh"]

        df_bess["Ingreso_USD"] = (
            df_bess["Energia_MWh"]
            * df_bess["Precio_Aplicado"]
        )

        df_bess["Fecha_Dia"] = (
            df_bess["Fecha"].dt.normalize()
        )

        df_bess["Anio"] = (
            df_bess["Fecha"].dt.year
        )

        df_bess["Mes"] = (
            df_bess["Fecha"].dt.month
        )

        diario = (
            df_bess
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

        diario = (
            diario
            .pivot_table(
                index="Fecha_Dia",
                columns="Tipo_Contrato",
                values="Energia_MWh",
                aggfunc="sum",
                fill_value=0
            )
            .reset_index()
        )

        diario.columns.name = None

        if "PPA" not in diario.columns:
            diario["PPA"] = 0

        if "SPOT" not in diario.columns:
            diario["SPOT"] = 0

        diario = diario.rename(
            columns={
                "PPA": "Energia_PPA_MWh",
                "SPOT": "Energia_Spot_MWh"
            }
        )

        resumenes_diarios.append(diario)

        mensual = (
            df_bess
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

        energia = (
            mensual
            .pivot_table(
                index=["Anio", "Mes"],
                columns="Tipo_Contrato",
                values="Energia_MWh",
                fill_value=0
            )
            .reset_index()
        )

        ingreso = (
            mensual
            .pivot_table(
                index=["Anio", "Mes"],
                columns="Tipo_Contrato",
                values="Ingreso_USD",
                fill_value=0
            )
            .reset_index()
        )

        energia.columns.name = None
        ingreso.columns.name = None

        for columna in ["PPA", "SPOT"]:
            if columna not in energia.columns:
                energia[columna] = 0

            if columna not in ingreso.columns:
                ingreso[columna] = 0

        energia = energia.rename(
            columns={
                "PPA": "Energia_PPA_MWh",
                "SPOT": "Energia_Spot_MWh"
            }
        )

        ingreso = ingreso.rename(
            columns={
                "PPA": "Ingreso_PPA_USD",
                "SPOT": "Ingreso_Spot_USD"
            }
        )

        resumen = energia.merge(
            ingreso,
            on=["Anio", "Mes"],
            how="outer"
        ).fillna(0)

        resumen["Energia_Total_MWh"] = (
            resumen["Energia_PPA_MWh"]
            + resumen["Energia_Spot_MWh"]
        )

        resumen["Ingreso_Total_USD"] = (
            resumen["Ingreso_PPA_USD"]
            + resumen["Ingreso_Spot_USD"]
        )

        resumen["Fuente"] = "API"

        resumenes_mensuales.append(resumen)

    mensual_final = (
        pd.concat(
            resumenes_mensuales,
            ignore_index=True
        )
        if resumenes_mensuales
        else pd.DataFrame()
    )

    diario_final = (
        pd.concat(
            resumenes_diarios,
            ignore_index=True
        )
        if resumenes_diarios
        else pd.DataFrame()
    )

    return (
        mensual_final,
        diario_final,
        sorted(set(advertencias))
    )


# =========================================================
# VALIDACION DEL ARCHIVO EXCEL
# =========================================================

try:
    df_budgets_bess = cargar_budgets_bess()
    df_consolidado_bess = cargar_consolidado_bess()
    df_budgets_proyectos = cargar_budgets_proyectos()

except Exception as error:
    st.error(
        "No fue posible validar budgets.xlsx."
    )
    st.info(str(error))
    st.stop()


# =========================================================
# DESCARGAR OTROS PROYECTOS
# =========================================================

with st.spinner(
    "Consultando generación de proyectos..."
):
    (
        df_generacion_proyectos,
        advertencias_proyectos
    ) = descargar_generacion_proyectos(
        PERIODOS_ANIO
    )


if advertencias_proyectos:
    st.warning(
        "Algunas mediciones de proyectos no estuvieron "
        "disponibles."
    )

    with st.expander(
        "Ver advertencias de proyectos"
    ):
        for advertencia in advertencias_proyectos:
            st.write(f"- {advertencia}")


# =========================================================
# DESCARGAR BESS SOLO PARA MESES NO CONSOLIDADOS
# =========================================================

periodos_excel_bess = {
    f"{int(fila.Anio):04d}{int(fila.Mes):02d}"
    for fila in df_consolidado_bess.itertuples()
}

periodos_pendientes_bess = [
    periodo
    for periodo in PERIODOS_ANIO
    if periodo not in periodos_excel_bess
]

with st.spinner(
    "Consultando períodos pendientes del BESS..."
):
    (
        df_bess_api,
        df_bess_diario,
        advertencias_bess
    ) = procesar_bess_api(
        periodos_pendientes_bess
    )


if advertencias_bess:
    st.warning(
        "Algunos períodos del BESS no estuvieron "
        "disponibles."
    )

    with st.expander(
        "Ver advertencias del BESS"
    ):
        for advertencia in advertencias_bess:
            st.write(f"- {advertencia}")


# =========================================================
# UNIR CONSOLIDADO BESS CON API
# =========================================================

columnas_bess = [
    "Anio",
    "Mes",
    "Energia_PPA_MWh",
    "Energia_Spot_MWh",
    "Energia_Total_MWh",
    "Ingreso_PPA_USD",
    "Ingreso_Spot_USD",
    "Ingreso_Total_USD",
    "Fuente"
]

fuentes_bess = [
    df_consolidado_bess[columnas_bess]
]

if not df_bess_api.empty:
    fuentes_bess.append(
        df_bess_api[columnas_bess]
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
# FILTROS
# =========================================================

st.sidebar.header("Filtros")

anio_seleccionado = st.sidebar.selectbox(
    "Año",
    options=[ANIO_OPERACIONAL],
    index=0
)


# =========================================================
# SECCION OTROS PROYECTOS
# =========================================================

st.header("Generación mensual de proyectos")

proyectos_disponibles = sorted(
    PROYECTOS_GENERACION.keys()
)

proyecto_seleccionado = st.selectbox(
    "Selecciona un proyecto",
    options=proyectos_disponibles
)

datos_proyecto = (
    df_generacion_proyectos[
        (
            df_generacion_proyectos["Proyecto"]
            == proyecto_seleccionado
        )
        & (
            df_generacion_proyectos["Anio"]
            == anio_seleccionado
        )
    ]
    .copy()
)

budget_proyecto = (
    df_budgets_proyectos[
        (
            df_budgets_proyectos["Proyecto"]
            == proyecto_seleccionado
        )
        & (
            df_budgets_proyectos["Anio"]
            == anio_seleccionado
        )
    ]
    .copy()
)

consolidado_proyecto = pd.DataFrame({
    "Mes": range(1, 13)
})

consolidado_proyecto = (
    consolidado_proyecto
    .merge(
        datos_proyecto[
            [
                "Mes",
                "Generacion_MWh"
            ]
        ],
        on="Mes",
        how="left"
    )
    .merge(
        budget_proyecto[
            [
                "Mes",
                "Budget_Generacion_MWh"
            ]
        ],
        on="Mes",
        how="left"
    )
)

consolidado_proyecto["Generacion_MWh"] = (
    consolidado_proyecto["Generacion_MWh"]
    .fillna(0)
)

consolidado_proyecto["Nombre_Mes"] = (
    consolidado_proyecto["Mes"]
    .map(NOMBRES_MESES)
)

meses_con_generacion = consolidado_proyecto[
    consolidado_proyecto["Generacion_MWh"] > 0
]

if not meses_con_generacion.empty:
    ultimo_mes = int(
        meses_con_generacion["Mes"].max()
    )

    fila_ultimo_mes = (
        consolidado_proyecto[
            consolidado_proyecto["Mes"]
            == ultimo_mes
        ]
        .iloc[0]
    )

    generacion_mes = (
        fila_ultimo_mes["Generacion_MWh"]
    )

    budget_mes = (
        fila_ultimo_mes[
            "Budget_Generacion_MWh"
        ]
    )

    cumplimiento_mes = (
        generacion_mes / budget_mes * 100
        if pd.notna(budget_mes)
        and budget_mes > 0
        else 0
    )

    generacion_acumulada = (
        consolidado_proyecto.loc[
            consolidado_proyecto["Mes"]
            <= ultimo_mes,
            "Generacion_MWh"
        ]
        .sum()
    )

    budget_acumulado = (
        consolidado_proyecto.loc[
            consolidado_proyecto["Mes"]
            <= ultimo_mes,
            "Budget_Generacion_MWh"
        ]
        .fillna(0)
        .sum()
    )

    col_p1, col_p2, col_p3, col_p4 = (
        st.columns(4)
    )

    col_p1.metric(
        f"Generación {MESES_COMPLETOS[ultimo_mes]}",
        f"{generacion_mes:,.1f} MWh"
    )

    col_p2.metric(
        "Budget del mes",
        (
            f"{budget_mes:,.1f} MWh"
            if pd.notna(budget_mes)
            else "Sin budget"
        )
    )

    col_p3.metric(
        "Cumplimiento del mes",
        f"{cumplimiento_mes:,.1f} %"
    )

    col_p4.metric(
        "Generación acumulada",
        f"{generacion_acumulada:,.1f} MWh"
    )


fig_proyecto = go.Figure()

fig_proyecto.add_trace(
    go.Bar(
        x=consolidado_proyecto["Nombre_Mes"],
        y=consolidado_proyecto[
            "Generacion_MWh"
        ],
        name="Generación real",
        marker_color="#1683DB",
        width=0.55,
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Generación: %{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

fig_proyecto.add_trace(
    go.Scatter(
        x=consolidado_proyecto["Nombre_Mes"],
        y=consolidado_proyecto[
            "Budget_Generacion_MWh"
        ],
        name="Budget generación",
        mode="lines+markers",
        line=dict(
            color="#FF3B30",
            width=3
        ),
        marker=dict(
            size=8
        ),
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Budget: %{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

fig_proyecto.update_layout(
    title=(
        f"Generación mensual "
        f"{proyecto_seleccionado}"
    ),
    xaxis_title="Mes",
    yaxis_title="Generación [MWh]",
    hovermode="x unified",
    height=540,
    bargap=0.35
)

fig_proyecto.update_yaxes(
    rangemode="tozero",
    gridcolor="rgba(140,140,140,0.25)"
)

st.plotly_chart(
    fig_proyecto,
    use_container_width=True
)


# =========================================================
# GRAFICO DE CUMPLIMIENTO DE PROYECTOS
# =========================================================

consolidado_proyecto[
    "Cumplimiento_Pct"
] = (
    consolidado_proyecto["Generacion_MWh"]
    / consolidado_proyecto[
        "Budget_Generacion_MWh"
    ].replace(0, pd.NA)
    * 100
)

fig_cumplimiento = go.Figure()

fig_cumplimiento.add_trace(
    go.Bar(
        x=consolidado_proyecto["Nombre_Mes"],
        y=consolidado_proyecto[
            "Cumplimiento_Pct"
        ],
        name="Cumplimiento",
        marker_color=[
            (
                "#149447"
                if pd.notna(valor)
                and valor >= 100
                else "#D94141"
            )
            for valor in consolidado_proyecto[
                "Cumplimiento_Pct"
            ]
        ],
        hovertemplate=(
            "<b>%{x}</b><br>"
            "Cumplimiento: %{y:,.1f}%"
            "<extra></extra>"
        )
    )
)

fig_cumplimiento.add_hline(
    y=100,
    line_dash="dash",
    line_color="#111111",
    annotation_text="Budget 100%"
)

fig_cumplimiento.update_layout(
    title="Cumplimiento mensual de generación",
    xaxis_title="Mes",
    yaxis_title="Cumplimiento [%]",
    height=430
)

fig_cumplimiento.update_yaxes(
    range=[0, 120],
    dtick=20,
    ticksuffix="%"
)

st.plotly_chart(
    fig_cumplimiento,
    use_container_width=True
)


# =========================================================
# SECCION BESS
# =========================================================

st.header("BESS María Elena")

bess_anio = df_bess_final[
    df_bess_final["Anio"]
    == anio_seleccionado
].copy()

if not bess_anio.empty:
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

    fila_bess = (
        bess_anio[
            bess_anio["Mes"] == mes_bess
        ]
        .iloc[0]
    )

    col_b1, col_b2, col_b3, col_b4 = (
        st.columns(4)
    )

    col_b1.metric(
        "Inyección total del mes",
        f"{fila_bess['Energia_Total_MWh']:,.1f} MWh"
    )

    col_b2.metric(
        "Energía PPA del mes",
        f"{fila_bess['Energia_PPA_MWh']:,.1f} MWh"
    )

    col_b3.metric(
        "Energía Spot del mes",
        f"{fila_bess['Energia_Spot_MWh']:,.1f} MWh"
    )

    col_b4.metric(
        "Ingreso del mes",
        f"USD {fila_bess['Ingreso_Total_USD']:,.0f}"
    )

    st.caption(
        f"Fuente del mes: {fila_bess['Fuente']}. "
        f"Precio PPA: USD {PRECIO_PPA:,.2f}/MWh."
    )


# =========================================================
# CONSOLIDADO ANUAL DEL BESS
# =========================================================

consolidado_bess = pd.DataFrame({
    "Mes": range(1, 13)
})

consolidado_bess = (
    consolidado_bess
    .merge(
        bess_anio,
        on="Mes",
        how="left"
    )
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
    consolidado_bess[columna] = (
        consolidado_bess[columna]
        .fillna(0)
    )

consolidado_bess["Nombre_Mes"] = (
    consolidado_bess["Mes"]
    .map(NOMBRES_MESES)
)

fig_bess = go.Figure()

fig_bess.add_trace(
    go.Bar(
        x=consolidado_bess["Nombre_Mes"],
        y=consolidado_bess["Energia_PPA_MWh"],
        name="PPA real",
        marker_color="#149447",
        width=0.58
    )
)

fig_bess.add_trace(
    go.Bar(
        x=consolidado_bess["Nombre_Mes"],
        y=consolidado_bess["Energia_Spot_MWh"],
        name="Spot real",
        marker_color="#D94141",
        width=0.58
    )
)

fig_bess.add_trace(
    go.Scatter(
        x=consolidado_bess["Nombre_Mes"],
        y=consolidado_bess[
            "Budget_Generacion_MWh"
        ],
        name="Budget generación",
        mode="lines+markers",
        line=dict(
            color="#FF3333",
            width=3
        )
    )
)

fig_bess.add_trace(
    go.Scatter(
        x=consolidado_bess["Nombre_Mes"],
        y=consolidado_bess[
            "Budget_PPA_MWh"
        ],
        name="Budget PPA",
        mode="lines+markers",
        line=dict(
            color="#111111",
            width=3
        )
    )
)

fig_bess.update_layout(
    title="Consolidado anual BESS",
    barmode="stack",
    xaxis_title="Mes",
    yaxis_title="Energía [MWh]",
    hovermode="x unified",
    height=580
)

st.plotly_chart(
    fig_bess,
    use_container_width=True
)


# =========================================================
# INGRESOS ACUMULADOS BESS
# =========================================================

ingreso_bess_acumulado = (
    bess_anio["Ingreso_Total_USD"]
    .sum()
)

energia_bess_acumulada = (
    bess_anio["Energia_Total_MWh"]
    .sum()
)

with st.expander(
    "Ver acumulados del BESS",
    expanded=False
):
    col_a1, col_a2 = st.columns(2)

    col_a1.metric(
        "Inyección acumulada BESS",
        f"{energia_bess_acumulada:,.1f} MWh"
    )

    col_a2.metric(
        "Ingreso acumulado BESS",
        f"USD {ingreso_bess_acumulado:,.0f}"
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