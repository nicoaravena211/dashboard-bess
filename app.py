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
# CONFIGURACION GENERAL DE STREAMLIT
# =========================================================

st.set_page_config(
    page_title="Dashboard BESS",
    page_icon="⚡",
    layout="wide"
)

DOCE_HORAS_SEGUNDOS = 12 * 60 * 60
DOCE_HORAS_MILISEGUNDOS = 12 * 60 * 60 * 1000

st_autorefresh(
    interval=DOCE_HORAS_MILISEGUNDOS,
    key="actualizacion_dashboard_12_horas"
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

elif fecha_actual.year > ANIO_OPERACIONAL:
    FECHA_FIN = f"{ANIO_OPERACIONAL}-12"

else:
    FECHA_FIN = f"{ANIO_OPERACIONAL}-01"


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

def quitar_tildes(texto):
    texto = str(texto)

    return "".join(
        caracter
        for caracter in unicodedata.normalize("NFD", texto)
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
    """
    Permite encabezados como Año, AÑO, Anio, ANIO,
    espacios adicionales y diferencias entre mayúsculas
    y minúsculas.
    """

    df = df.copy()

    equivalencias = {
        "ano": "Anio",
        "anio": "Anio",
        "mes": "Mes",
        "budget_generacion_mwh": "Budget_Generacion_MWh",
        "budget_ppa_mwh": "Budget_PPA_MWh",
        "energia_ppa_mwh": "Energia_PPA_MWh",
        "energia_spot_mwh": "Energia_Spot_MWh",
        "ingreso_ppa_usd": "Ingreso_PPA_USD",
        "ingreso_spot_usd": "Ingreso_Spot_USD",
        "fuente": "Fuente"
    }

    nuevos_nombres = {}

    for columna in df.columns:
        nombre_normalizado = normalizar_nombre_columna(
            columna
        )

        if nombre_normalizado in equivalencias:
            nuevos_nombres[columna] = equivalencias[
                nombre_normalizado
            ]
        else:
            nuevos_nombres[columna] = str(
                columna
            ).strip()

    return df.rename(columns=nuevos_nombres)


def convertir_numerico(serie):
    """
    Convierte números de Excel o API a formato numérico.
    Acepta números reales o texto con coma decimal.
    """

    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(
            serie,
            errors="coerce"
        )

    serie_limpia = (
        serie.astype(str)
        .str.strip()
        .str.replace(" ", "", regex=False)
    )

    return pd.to_numeric(
        serie_limpia.str.replace(",", ".", regex=False),
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


# =========================================================
# LECTURA Y VALIDACION DE BUDGETS
# =========================================================

@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def cargar_budgets():
    if not os.path.exists(ARCHIVO_EXCEL):
        raise FileNotFoundError(
            f"No se encontró {ARCHIVO_EXCEL} "
            "en la carpeta principal."
        )

    df_budget = pd.read_excel(
        ARCHIVO_EXCEL,
        sheet_name=HOJA_BUDGETS,
        engine="openpyxl"
    )

    df_budget = normalizar_columnas_excel(
        df_budget
    )

    columnas_requeridas = [
        "Anio",
        "Mes",
        "Budget_Generacion_MWh",
        "Budget_PPA_MWh"
    ]

    columnas_faltantes = [
        columna
        for columna in columnas_requeridas
        if columna not in df_budget.columns
    ]

    if columnas_faltantes:
        raise ValueError(
            "Faltan columnas en la hoja Budgets: "
            + ", ".join(columnas_faltantes)
            + ". Columnas encontradas: "
            + ", ".join(
                str(columna)
                for columna in df_budget.columns
            )
        )

    df_budget = df_budget[
        columnas_requeridas
    ].copy()

    for columna in columnas_requeridas:
        df_budget[columna] = convertir_numerico(
            df_budget[columna]
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

    if not df_budget["Mes"].between(
        1,
        12
    ).all():
        raise ValueError(
            "La hoja Budgets contiene meses fuera "
            "del rango 1 a 12."
        )

    columnas_valores = [
        "Budget_Generacion_MWh",
        "Budget_PPA_MWh"
    ]

    if df_budget[
        columnas_valores
    ].isna().any().any():
        raise ValueError(
            "La hoja Budgets contiene valores vacíos "
            "o no numéricos."
        )

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
            "Hay meses duplicados en Budgets: "
            + ", ".join(meses_duplicados)
        )

    return df_budget


# =========================================================
# LECTURA Y VALIDACION DEL CONSOLIDADO
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

    df_consolidado = normalizar_columnas_excel(
        df_consolidado
    )

    columnas_requeridas = [
        "Anio",
        "Mes",
        "Energia_PPA_MWh",
        "Energia_Spot_MWh",
        "Ingreso_PPA_USD",
        "Ingreso_Spot_USD"
    ]

    columnas_faltantes = [
        columna
        for columna in columnas_requeridas
        if columna not in df_consolidado.columns
    ]

    if columnas_faltantes:
        raise ValueError(
            "Faltan columnas en la hoja Consolidado: "
            + ", ".join(columnas_faltantes)
            + ". Columnas encontradas: "
            + ", ".join(
                str(columna)
                for columna in df_consolidado.columns
            )
        )

    columnas_a_conservar = (
        columnas_requeridas
        + (
            ["Fuente"]
            if "Fuente" in df_consolidado.columns
            else []
        )
    )

    df_consolidado = df_consolidado[
        columnas_a_conservar
    ].copy()

    columnas_numericas = [
        "Anio",
        "Mes",
        "Energia_PPA_MWh",
        "Energia_Spot_MWh",
        "Ingreso_PPA_USD",
        "Ingreso_Spot_USD"
    ]

    for columna in columnas_numericas:
        df_consolidado[columna] = convertir_numerico(
            df_consolidado[columna]
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

    if not df_consolidado["Mes"].between(
        1,
        12
    ).all():
        raise ValueError(
            "La hoja Consolidado contiene meses "
            "fuera del rango 1 a 12."
        )

    columnas_valores = [
        "Energia_PPA_MWh",
        "Energia_Spot_MWh",
        "Ingreso_PPA_USD",
        "Ingreso_Spot_USD"
    ]

    if df_consolidado[
        columnas_valores
    ].isna().any().any():
        raise ValueError(
            "La hoja Consolidado contiene valores "
            "vacíos o no numéricos."
        )

    if (
        df_consolidado[columnas_valores] < 0
    ).any().any():
        raise ValueError(
            "La hoja Consolidado contiene valores negativos."
        )

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
            "Hay meses duplicados en Consolidado: "
            + ", ".join(meses_duplicados)
        )

    df_consolidado["Energia_Total_MWh"] = (
        df_consolidado["Energia_PPA_MWh"]
        + df_consolidado["Energia_Spot_MWh"]
    )

    df_consolidado["Ingreso_Total_USD"] = (
        df_consolidado["Ingreso_PPA_USD"]
        + df_consolidado["Ingreso_Spot_USD"]
    )

    df_consolidado[
        "Precio_Promedio_USD_MWh"
    ] = (
        df_consolidado["Ingreso_Total_USD"]
        / df_consolidado[
            "Energia_Total_MWh"
        ].replace(0, pd.NA)
    ).fillna(0)

    # La fuente se crea automáticamente.
    # No es necesario incluirla en el Excel.
    df_consolidado["Fuente"] = (
        "Excel consolidado"
    )

    return df_consolidado


# =========================================================
# DETERMINAR MESES QUE DEBE CONSULTAR LA API
# =========================================================

def determinar_periodos(
    fecha_inicio,
    fecha_fin,
    df_consolidado
):
    periodos_totales = generar_periodos(
        fecha_inicio,
        fecha_fin
    )

    periodos_excel = set()

    for fila in df_consolidado.itertuples():
        periodo = (
            f"{int(fila.Anio):04d}"
            f"{int(fila.Mes):02d}"
        )

        periodos_excel.add(periodo)

    periodos_pendientes = [
        periodo
        for periodo in periodos_totales
        if periodo not in periodos_excel
    ]

    periodos_consolidados = [
        periodo
        for periodo in periodos_totales
        if periodo in periodos_excel
    ]

    return (
        periodos_totales,
        periodos_consolidados,
        periodos_pendientes
    )


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

    for intento in range(3):
        try:
            respuesta = requests.get(
                URL_PRMTE,
                params=params,
                timeout=60
            )

            if respuesta.status_code == 200:
                contenido = respuesta.json()

                if isinstance(contenido, list):
                    return contenido

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
    advertencias = []

    for periodo in periodos:
        registros_periodo = []
        mpids_correctos = 0

        for mpid in BESS_MPIDS:
            datos = obtener_datos_prmte(
                periodo + "012345",
                mpid
            )

            if datos is None:
                advertencias.append(
                    f"{periodo_a_texto(periodo)}: "
                    f"PRMTE no respondió para {mpid}."
                )

                continue

            registros_mpid = []

            for bloque in datos:
                mediciones = bloque.get(
                    "measurement",
                    []
                )

                for medida in mediciones:
                    registros_mpid.append({
                        "Fecha": medida.get("dateRange"),
                        "Energia_kWh": medida.get("channel3"),
                        "MPID": mpid,
                        "Periodo": periodo
                    })

            if registros_mpid:
                mpids_correctos += 1
                registros_periodo.extend(
                    registros_mpid
                )
            else:
                advertencias.append(
                    f"{periodo_a_texto(periodo)}: "
                    f"sin mediciones para {mpid}."
                )

        # Se utiliza el mes solamente si respondieron
        # correctamente los cuatro medidores.
        if mpids_correctos == len(BESS_MPIDS):
            registros_totales.extend(
                registros_periodo
            )
        else:
            advertencias.append(
                f"{periodo_a_texto(periodo)} no se incorporó "
                "porque no están disponibles los cuatro MPID."
            )

    if not registros_totales:
        return (
            pd.DataFrame(
                columns=["Fecha", "Energia_kWh"]
            ),
            sorted(set(advertencias))
        )

    df_mediciones = pd.DataFrame(
        registros_totales
    )

    df_mediciones["Fecha"] = parse_fecha_safe(
        df_mediciones["Fecha"]
    )

    df_mediciones["Energia_kWh"] = convertir_numerico(
        df_mediciones["Energia_kWh"]
    ).fillna(0)

    df_mediciones = df_mediciones.dropna(
        subset=["Fecha"]
    )

    df_bess = (
        df_mediciones
        .groupby("Fecha", as_index=False)
        .agg(
            Energia_kWh=(
                "Energia_kWh",
                "sum"
            )
        )
        .sort_values("Fecha")
    )

    return df_bess, sorted(set(advertencias))


# =========================================================
# CONSULTA API CMG
# =========================================================

def descargar_cmg_mes(periodo):
    anio = int(periodo[:4])
    mes = int(periodo[4:6])

    ultimo_dia_numero = calendar.monthrange(
        anio,
        mes
    )[1]

    fecha_inicio_mes = pd.Timestamp(
        year=anio,
        month=mes,
        day=1
    )

    fecha_fin_mes = pd.Timestamp(
        year=anio,
        month=mes,
        day=ultimo_dia_numero
    )

    hoy = pd.Timestamp.now().normalize()

    if (
        anio == hoy.year
        and mes == hoy.month
    ):
        fecha_fin_mes = min(
            fecha_fin_mes,
            hoy
        )

    registros = []
    pagina = 0

    while True:
        params = {
            "startDate": fecha_inicio_mes.strftime(
                "%Y-%m-%d"
            ),
            "endDate": fecha_fin_mes.strftime(
                "%Y-%m-%d"
            ),
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

            contenido = respuesta.json()

        except (
            requests.RequestException,
            ValueError
        ):
            return None

        datos_pagina = contenido.get(
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
            columns=["Fecha", "CMG_USD_MWh"]
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

    df_cmg["CMG_USD_MWh"] = convertir_numerico(
        df_cmg["cmg_usd_mwh_"]
    ).fillna(0)

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
        .sort_values("Fecha")
    )

    return df_cmg


@st.cache_data(
    ttl=DOCE_HORAS_SEGUNDOS,
    show_spinner=False
)
def descargar_cmg_periodos(periodos):
    dataframes = []
    advertencias = []

    for periodo in periodos:
        resultado = descargar_cmg_mes(
            periodo
        )

        if resultado is None:
            advertencias.append(
                f"{periodo_a_texto(periodo)}: "
                "CMG no disponible."
            )

            continue

        if resultado.empty:
            advertencias.append(
                f"{periodo_a_texto(periodo)}: "
                "CMG sin registros."
            )

            continue

        dataframes.append(resultado)

    if not dataframes:
        return (
            pd.DataFrame(
                columns=["Fecha", "CMG_USD_MWh"]
            ),
            sorted(set(advertencias))
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

    return df_cmg, sorted(set(advertencias))


# =========================================================
# CARGA Y VALIDACION DEL EXCEL
# =========================================================

try:
    df_budgets = cargar_budgets()
    df_consolidado_excel = (
        cargar_consolidado_excel()
    )

except Exception as error:
    st.error(
        "No fue posible validar budgets.xlsx."
    )

    st.info(str(error))

    st.stop()


(
    periodos_totales,
    periodos_consolidados,
    periodos_pendientes
) = determinar_periodos(
    FECHA_INICIO,
    FECHA_FIN,
    df_consolidado_excel
)


# =========================================================
# ESTADO DE LAS FUENTES
# =========================================================

with st.expander(
    "Estado de las fuentes de información",
    expanded=False
):
    st.markdown(
        "**Meses tomados desde el Excel:**"
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

    st.markdown(
        "**Meses que se consultarán automáticamente "
        "desde las API:**"
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
# DESCARGA DE PERIODOS PENDIENTES
# =========================================================

with st.spinner(
    "Consultando períodos pendientes en las API..."
):
    df_bess, advertencias_prmte = (
        descargar_bess_periodos(
            periodos_pendientes
        )
    )

    df_cmg, advertencias_cmg = (
        descargar_cmg_periodos(
            periodos_pendientes
        )
    )


advertencias_api = sorted(
    set(
        advertencias_prmte
        + advertencias_cmg
    )
)

if advertencias_api:
    st.warning(
        "Algunas consultas no estuvieron disponibles. "
        "El dashboard continuará con la información válida."
    )

    with st.expander(
        "Ver advertencias de las API"
    ):
        for advertencia in advertencias_api:
            st.write(
                f"- {advertencia}"
            )


# =========================================================
# PROCESAMIENTO DE LA API
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

    # -----------------------------------------------------
    # RESUMEN DIARIO
    # -----------------------------------------------------

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

    resumen_diario_api = (
        resumen_diario_api.rename(
            columns={
                "PPA": "Energia_PPA_MWh",
                "SPOT": "Energia_Spot_MWh"
            }
        )
    )

    resumen_diario_api[
        "Energia_Total_MWh"
    ] = (
        resumen_diario_api[
            "Energia_PPA_MWh"
        ]
        + resumen_diario_api[
            "Energia_Spot_MWh"
        ]
    )

    # -----------------------------------------------------
    # RESUMEN MENSUAL
    # -----------------------------------------------------

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

    resumen_mensual_api[
        "Energia_Total_MWh"
    ] = (
        resumen_mensual_api[
            "Energia_PPA_MWh"
        ]
        + resumen_mensual_api[
            "Energia_Spot_MWh"
        ]
    )

    resumen_mensual_api[
        "Ingreso_Total_USD"
    ] = (
        resumen_mensual_api[
            "Ingreso_PPA_USD"
        ]
        + resumen_mensual_api[
            "Ingreso_Spot_USD"
        ]
    )

    resumen_mensual_api[
        "Precio_Promedio_USD_MWh"
    ] = (
        resumen_mensual_api[
            "Ingreso_Total_USD"
        ]
        / resumen_mensual_api[
            "Energia_Total_MWh"
        ].replace(0, pd.NA)
    ).fillna(0)

    resumen_mensual_api["Fuente"] = "API"


# =========================================================
# UNION DEL EXCEL Y LA API
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

fuentes = [
    df_consolidado_excel[
        columnas_finales
    ]
]

if not resumen_mensual_api.empty:
    fuentes.append(
        resumen_mensual_api[
            columnas_finales
        ]
    )

resumen_mensual_final = pd.concat(
    fuentes,
    ignore_index=True
)

# El Excel siempre tiene prioridad si por alguna razón
# el mismo mes aparece también en la API.

resumen_mensual_final["Prioridad"] = (
    resumen_mensual_final["Fuente"]
    .apply(
        lambda valor: (
            2
            if valor == "Excel consolidado"
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
# FILTRO ANUAL
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

resumen_anio = resumen_mensual_final[
    resumen_mensual_final["Anio"]
    == anio_seleccionado
].copy()


# =========================================================
# INDICADORES ANUALES
# =========================================================

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

st.subheader(
    "Distribución diaria de energía"
)

if resumen_diario_api.empty:
    st.info(
        "No hay información diaria disponible desde las API. "
        "Los meses del Excel contienen totales mensuales."
    )

else:
    meses_diarios = (
        resumen_diario_api
        .assign(
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

    opciones_diarias = [
        (
            int(fila.Anio),
            int(fila.Mes)
        )
        for fila in meses_diarios.itertuples()
    ]

    seleccion_diaria = st.sidebar.selectbox(
        "Mes para gráfico diario",
        options=opciones_diarias,
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

    primer_dia = pd.Timestamp(
        year=anio_diario,
        month=mes_diario,
        day=1
    )

    ultimo_dia_numero = calendar.monthrange(
        anio_diario,
        mes_diario
    )[1]

    ultimo_dia = pd.Timestamp(
        year=anio_diario,
        month=mes_diario,
        day=ultimo_dia_numero
    )

    hoy = pd.Timestamp.now().normalize()

    if (
        anio_diario == hoy.year
        and mes_diario == hoy.month
    ):
        ultimo_dia = min(
            ultimo_dia,
            hoy
        )

    calendario_mes = pd.DataFrame({
        "Fecha_Dia": pd.date_range(
            start=primer_dia,
            end=ultimo_dia,
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
            y=diario_mes[
                "Energia_PPA_MWh"
            ],
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
            y=diario_mes[
                "Energia_Spot_MWh"
            ],
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
            primer_dia
            - pd.Timedelta(hours=12),
            ultimo_dia
            + pd.Timedelta(hours=12)
        ],
        showgrid=False
    )

    fig_diario.update_yaxes(
        rangemode="tozero",
        gridcolor=(
            "rgba(140, 140, 140, 0.25)"
        )
    )

    st.plotly_chart(
        fig_diario,
        use_container_width=True
    )


# =========================================================
# GRAFICO CONSOLIDADO ANUAL
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

columnas_resultados = [
    "Energia_PPA_MWh",
    "Energia_Spot_MWh",
    "Energia_Total_MWh",
    "Ingreso_PPA_USD",
    "Ingreso_Spot_USD",
    "Ingreso_Total_USD"
]

for columna in columnas_resultados:
    if columna not in consolidado_anual.columns:
        consolidado_anual[columna] = 0

consolidado_anual[columnas_resultados] = (
    consolidado_anual[columnas_resultados]
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
        y=consolidado_anual[
            "Energia_PPA_MWh"
        ],
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
        y=consolidado_anual[
            "Energia_Spot_MWh"
        ],
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
            "Budget generación: "
            "%{y:,.2f} MWh"
            "<extra></extra>"
        )
    )
)

fig_anual.add_trace(
    go.Scatter(
        x=consolidado_anual["Nombre_Mes"],
        y=consolidado_anual[
            "Budget_PPA_MWh"
        ],
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
    gridcolor=(
        "rgba(140, 140, 140, 0.25)"
    )
)

st.plotly_chart(
    fig_anual,
    use_container_width=True
)


# =========================================================
# CUMPLIMIENTO ACUMULADO
# =========================================================

meses_con_resultados = (
    consolidado_anual[
        "Energia_Total_MWh"
    ] > 0
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
# FUENTE UTILIZADA
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
    "Las consultas API y la caché se actualizan "
    "cada 12 horas."
)