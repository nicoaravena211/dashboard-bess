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
st.set_page_config(page_title="Dashboard de Operación", page_icon="⚡", layout="wide")

CACHE_SEGUNDOS = 12 * 60 * 60
REFRESCO_MS = 12 * 60 * 60 * 1000
st_autorefresh(interval=REFRESCO_MS, key="actualizacion_dashboard_12h")

st.title("Dashboard de operación")
st.caption("Seguimiento del BESS María Elena y generación mensual de proyectos en operación")

try:
    API_KEY_PRMTE = st.secrets["API_KEY_PRMTE"]
    API_KEY_CMG = st.secrets["API_KEY_CMG"]
except KeyError:
    st.error("No se encontraron las credenciales API en Streamlit Secrets.")
    st.info("Configura API_KEY_PRMTE y API_KEY_CMG desde Manage app > Settings > Secrets.")
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
URL_PRMTE = "https://medidas.api.coordinador.cl/medidas-v2/measurement"
URL_CMG = "https://sipub.api.coordinador.cl/costo-marginal-real/v4/findByDate"

PROYECTOS = {
    "MARIA ELENA PFV": [{"mpid": "MARELENA_220_JT1_GSS", "canal": 3}],
    "SAN PEDRO III": [
        {"mpid": "SOLRJAMA_220_JT1_RUC", "canal": 3},
        {"mpid": "SOLRJAMA_220_JT2_RUC", "canal": 3},
    ],
    "DOÑA CARMEN": [{"mpid": "CDNCARMN_220_J1_ECM", "canal": 3}],
    "LA QUINTA": [{"mpid": "CABRERO_023_PMGD7_QTC", "canal": 3}],
    "LA PERLA": [{"mpid": "LSANGLES_013_PMGD7_LAP", "canal": 3}],
    "LA HUERTA": [{"mpid": "PARRONAL_013_PMGD2_HUE", "canal": 3}],
    "CHACAICO": [{"mpid": "LSANGLES_013_PMGD8_CCC", "canal": 3}],
    "SANCLEMENTE": [{"mpid": "SCLMENTE_013_PMGD4_CFL", "canal": 3}],
    "TRILALEO": [{"mpid": "CHOLGUAN_015_PMGD6_YTL", "canal": 3}],
    "COLLANCO": [{"mpid": "CNSTUCON_023_PMGD4_OAO", "canal": 3}],
}

BESS_MPIDS = [
    "MARELENA_023_E1_GSS",
    "MARELENA_023_E7_GSS",
    "MARELENA_023_E8_GSS",
    "MARELENA_023_E11_GSS",
]

PROYECTOS_PMGD = [
    "LA QUINTA",
    "LA PERLA",
    "LA HUERTA",
    "CHACAICO",
    "SANCLEMENTE",
    "TRILALEO",
    "COLLANCO",
]

ORDEN_GRAFICOS = [
    "MARIA ELENA PFV",
    "SAN PEDRO III",
    "DOÑA CARMEN",
    "CONSOLIDADO PMGDS",
    "LA QUINTA",
    "LA PERLA",
    "LA HUERTA",
    "CHACAICO",
    "SANCLEMENTE",
    "TRILALEO",
    "COLLANCO",
]

MESES_CORTOS = {1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"}
MESES_COMPLETOS = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

# =========================================================
# FUNCIONES GENERALES
# =========================================================
def quitar_tildes(texto):
    return "".join(c for c in unicodedata.normalize("NFD", str(texto)) if unicodedata.category(c) != "Mn")


def normalizar_nombre_columna(nombre):
    nombre = quitar_tildes(nombre).strip().replace(" ", "_").replace("-", "_")
    while "__" in nombre:
        nombre = nombre.replace("__", "_")
    return nombre.lower()


def normalizar_columnas_excel(df):
    equivalencias = {
        "ano":"Anio", "anio":"Anio", "mes":"Mes", "proyecto":"Proyecto",
        "budget_generacion_mwh":"Budget_Generacion_MWh", "budget_ppa_mwh":"Budget_PPA_MWh",
        "energia_ppa_mwh":"Energia_PPA_MWh", "energia_spot_mwh":"Energia_Spot_MWh",
        "ingreso_ppa_usd":"Ingreso_PPA_USD", "ingreso_spot_usd":"Ingreso_Spot_USD",
        "generacion_mwh":"Generacion_MWh",
    }
    return df.rename(columns={c: equivalencias.get(normalizar_nombre_columna(c), str(c).strip()) for c in df.columns})


def convertir_numerico(serie):
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")
    texto = serie.astype(str).str.strip().str.replace(" ", "", regex=False).str.replace(",", ".", regex=False)
    return pd.to_numeric(texto, errors="coerce")


def parse_fecha_safe(serie):
    return pd.to_datetime(serie.astype(str).str[:19], errors="coerce")


def generar_periodos(fecha_inicio, fecha_fin):
    inicio = datetime.strptime(fecha_inicio, "%Y-%m")
    fin = datetime.strptime(fecha_fin, "%Y-%m")
    periodos, actual = [], inicio
    while actual <= fin:
        periodos.append(actual.strftime("%Y%m"))
        actual = actual.replace(year=actual.year + 1, month=1) if actual.month == 12 else actual.replace(month=actual.month + 1)
    return periodos


def periodo_a_texto(periodo):
    return f"{MESES_COMPLETOS[int(periodo[4:6])]} {int(periodo[:4])}"


PERIODOS_DISPONIBLES = generar_periodos(FECHA_INICIO, FECHA_FIN)

# =========================================================
# EXCEL
# =========================================================
def leer_hoja_excel(hoja, columnas_requeridas):
    if not os.path.exists(ARCHIVO_EXCEL):
        raise FileNotFoundError(f"No se encontró {ARCHIVO_EXCEL}.")
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=hoja, engine="openpyxl")
    df = normalizar_columnas_excel(df)
    faltantes = [c for c in columnas_requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en {hoja}: " + ", ".join(faltantes))
    return df[columnas_requeridas].copy()


@st.cache_data(ttl=CACHE_SEGUNDOS, show_spinner=False)
def cargar_budgets_bess():
    req = ["Anio", "Mes", "Budget_Generacion_MWh", "Budget_PPA_MWh"]
    df = leer_hoja_excel(HOJA_BUDGETS_BESS, req)
    for c in req:
        df[c] = convertir_numerico(df[c])
    df = df.dropna(subset=req)
    df[["Anio", "Mes"]] = df[["Anio", "Mes"]].astype(int)
    if df.empty:
        raise ValueError("La hoja Budgets no contiene filas válidas.")
    if not df["Mes"].between(1, 12).all():
        raise ValueError("La hoja Budgets contiene meses fuera del rango 1 a 12.")
    if df.duplicated(["Anio", "Mes"]).any():
        raise ValueError("Hay meses duplicados en la hoja Budgets.")
    return df


@st.cache_data(ttl=CACHE_SEGUNDOS, show_spinner=False)
def cargar_consolidado_bess():
    req = ["Anio", "Mes", "Energia_PPA_MWh", "Energia_Spot_MWh", "Ingreso_PPA_USD", "Ingreso_Spot_USD"]
    df = leer_hoja_excel(HOJA_CONSOLIDADO_BESS, req)
    for c in req:
        df[c] = convertir_numerico(df[c])
    df = df.dropna(subset=req)
    df[["Anio", "Mes"]] = df[["Anio", "Mes"]].astype(int)
    if df.duplicated(["Anio", "Mes"]).any():
        raise ValueError("Hay meses duplicados en Consolidado.")
    df["Energia_Total_MWh"] = df["Energia_PPA_MWh"] + df["Energia_Spot_MWh"]
    df["Ingreso_Total_USD"] = df["Ingreso_PPA_USD"] + df["Ingreso_Spot_USD"]
    df["Fuente"] = "Excel consolidado"
    return df


@st.cache_data(ttl=CACHE_SEGUNDOS, show_spinner=False)
def cargar_budgets_proyectos():
    req = ["Anio", "Mes", "Proyecto", "Budget_Generacion_MWh"]
    df = leer_hoja_excel(HOJA_BUDGETS_PROYECTOS, req)
    df["Anio"] = convertir_numerico(df["Anio"])
    df["Mes"] = convertir_numerico(df["Mes"])
    df["Budget_Generacion_MWh"] = convertir_numerico(df["Budget_Generacion_MWh"])
    df["Proyecto"] = df["Proyecto"].astype(str).str.strip().str.upper()
    df = df.dropna(subset=req)
    df[["Anio", "Mes"]] = df[["Anio", "Mes"]].astype(int)
    desconocidos = sorted(set(df["Proyecto"]) - set(PROYECTOS))
    if desconocidos:
        raise ValueError("Proyectos no reconocidos en Budgets_Proyectos: " + ", ".join(desconocidos))
    if df.duplicated(["Anio", "Mes", "Proyecto"]).any():
        raise ValueError("Hay budgets duplicados en Budgets_Proyectos.")
    return df


@st.cache_data(ttl=CACHE_SEGUNDOS, show_spinner=False)
def cargar_consolidado_proyectos():
    req = ["Anio", "Mes", "Proyecto", "Generacion_MWh"]
    df = leer_hoja_excel(HOJA_CONSOLIDADO_PROYECTOS, req)
    df["Anio"] = convertir_numerico(df["Anio"])
    df["Mes"] = convertir_numerico(df["Mes"])
    df["Generacion_MWh"] = convertir_numerico(df["Generacion_MWh"])
    df["Proyecto"] = df["Proyecto"].astype(str).str.strip().str.upper()
    df = df.dropna(subset=req)
    df[["Anio", "Mes"]] = df[["Anio", "Mes"]].astype(int)
    desconocidos = sorted(set(df["Proyecto"]) - set(PROYECTOS))
    if desconocidos:
        raise ValueError("Proyectos no reconocidos en Consolidado_Proyectos: " + ", ".join(desconocidos))
    if df.duplicated(["Anio", "Mes", "Proyecto"]).any():
        raise ValueError("Hay registros duplicados en Consolidado_Proyectos.")
    if (df["Generacion_MWh"] < 0).any():
        raise ValueError("Consolidado_Proyectos contiene generación negativa.")
    df["Fuente"] = "Excel consolidado"
    return df

# =========================================================
# API
# =========================================================
def obtener_datos_prmte(periodo, mpid, canal):
    params = {"channelId": canal, "measurePointId": mpid, "period": periodo, "user_key": API_KEY_PRMTE}
    for intento in range(3):
        try:
            r = requests.get(URL_PRMTE, params=params, timeout=60)
            if r.status_code == 200:
                data = r.json()
                return data if isinstance(data, list) else []
        except (requests.RequestException, ValueError):
            pass
        if intento < 2:
            time.sleep(2)
    return None


def extraer_mediciones(datos, canal):
    registros, campo = [], f"channel{canal}"
    if datos is None:
        return registros
    for bloque in datos:
        for m in bloque.get("measurement", []):
            registros.append({"Fecha": m.get("dateRange"), "Energia_kWh": m.get(campo)})
    return registros


def descargar_cmg_mes(periodo):
    anio, mes = int(periodo[:4]), int(periodo[4:6])
    inicio = pd.Timestamp(anio, mes, 1)
    fin = pd.Timestamp(anio, mes, calendar.monthrange(anio, mes)[1])
    hoy_n = pd.Timestamp.now().normalize()
    if anio == hoy_n.year and mes == hoy_n.month:
        fin = min(fin, hoy_n)
    registros, pagina = [], 0
    while True:
        params = {"startDate": inicio.strftime("%Y-%m-%d"), "endDate": fin.strftime("%Y-%m-%d"), "page": pagina, "limit": 5000, "type": TIPO_CMG, "bar_transf": BARRA_CMG_BESS, "user_key": API_KEY_CMG}
        try:
            r = requests.get(URL_CMG, params=params, timeout=60)
            if r.status_code != 200:
                return None
            datos = r.json().get("data", [])
        except (requests.RequestException, ValueError):
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
    req = ["fecha", "hra", "min", "cmg_usd_mwh_"]
    if not all(c in df.columns for c in req):
        return None
    df["hra"] = convertir_numerico(df["hra"]).fillna(0)
    df["min"] = convertir_numerico(df["min"]).fillna(0)
    df["CMG_USD_MWh"] = convertir_numerico(df["cmg_usd_mwh_"]).fillna(0)
    df["Fecha"] = pd.to_datetime(df["fecha"], errors="coerce") + pd.to_timedelta(df["hra"], unit="h") + pd.to_timedelta(df["min"], unit="m")
    return df[["Fecha", "CMG_USD_MWh"]].dropna(subset=["Fecha"]).groupby("Fecha", as_index=False).agg(CMG_USD_MWh=("CMG_USD_MWh", "mean"))

# =========================================================
# DESCARGA DE PROYECTOS Y BESS
# =========================================================
def obtener_pendientes_proyectos(df_consolidado):
    existentes = {(str(f.Proyecto), f"{int(f.Anio):04d}{int(f.Mes):02d}") for f in df_consolidado.itertuples()}
    return {p: [per for per in PERIODOS_DISPONIBLES if (p, per) not in existentes] for p in PROYECTOS}


@st.cache_data(ttl=CACHE_SEGUNDOS, show_spinner=False)
def descargar_proyectos_pendientes(pendientes):
    resultados, advertencias = [], []
    for proyecto, periodos in pendientes.items():
        for periodo in periodos:
            registros, correctos = [], 0
            for conf in PROYECTOS[proyecto]:
                datos = obtener_datos_prmte(periodo + "012345", conf["mpid"], conf["canal"])
                regs = extraer_mediciones(datos, conf["canal"])
                if datos is None or not regs:
                    advertencias.append(f"{proyecto}, {periodo_a_texto(periodo)}: sin datos para {conf['mpid']}.")
                    continue
                correctos += 1
                registros.extend(regs)
            if correctos != len(PROYECTOS[proyecto]) or not registros:
                continue
            df = pd.DataFrame(registros)
            df["Fecha"] = parse_fecha_safe(df["Fecha"])
            df["Energia_kWh"] = convertir_numerico(df["Energia_kWh"]).fillna(0)
            df = df.dropna(subset=["Fecha"]).groupby("Fecha", as_index=False).agg(Energia_kWh=("Energia_kWh", "sum"))
            resultados.append({"Anio":int(periodo[:4]), "Mes":int(periodo[4:6]), "Proyecto":proyecto, "Generacion_MWh":df["Energia_kWh"].abs().sum()/1000, "Fuente":"API"})
    return pd.DataFrame(resultados, columns=["Anio","Mes","Proyecto","Generacion_MWh","Fuente"]), sorted(set(advertencias))


@st.cache_data(ttl=CACHE_SEGUNDOS, show_spinner=False)
def procesar_bess_api(periodos):
    mensuales, diarios, advertencias = [], [], []
    for periodo in periodos:
        registros, correctos = [], 0
        for mpid in BESS_MPIDS:
            datos = obtener_datos_prmte(periodo + "012345", mpid, 3)
            regs = extraer_mediciones(datos, 3)
            if datos is None or not regs:
                advertencias.append(f"BESS, {periodo_a_texto(periodo)}: sin datos para {mpid}.")
                continue
            correctos += 1
            registros.extend(regs)
        if correctos != len(BESS_MPIDS) or not registros:
            continue
        df = pd.DataFrame(registros)
        df["Fecha"] = parse_fecha_safe(df["Fecha"])
        df["Energia_kWh"] = convertir_numerico(df["Energia_kWh"]).fillna(0)
        df = df.dropna(subset=["Fecha"]).groupby("Fecha", as_index=False).agg(Energia_kWh=("Energia_kWh", "sum"))
        cmg = descargar_cmg_mes(periodo)
        if cmg is None or cmg.empty:
            advertencias.append(f"BESS, {periodo_a_texto(periodo)}: CMG no disponible.")
            continue
        df = df.merge(cmg, on="Fecha", how="left")
        df["CMG_USD_MWh"] = df["CMG_USD_MWh"].fillna(0)
        df["Energia_MWh"] = df["Energia_kWh"].abs()/1000
        df["Tipo"] = "PPA"
        df.loc[(df["Fecha"].dt.hour >= 6) & (df["Fecha"].dt.hour < 21), "Tipo"] = "SPOT"
        df["Precio"] = PRECIO_PPA
        df.loc[df["Tipo"] == "SPOT", "Precio"] = df["CMG_USD_MWh"]
        df["Ingreso_USD"] = df["Energia_MWh"] * df["Precio"]
        df["Fecha_Dia"] = df["Fecha"].dt.normalize()

        diario = df.groupby(["Fecha_Dia","Tipo"], as_index=False).agg(Energia_MWh=("Energia_MWh","sum")).pivot_table(index="Fecha_Dia", columns="Tipo", values="Energia_MWh", fill_value=0).reset_index()
        diario.columns.name = None
        for c in ["PPA","SPOT"]:
            if c not in diario.columns:
                diario[c] = 0
        diario = diario.rename(columns={"PPA":"Energia_PPA_MWh", "SPOT":"Energia_Spot_MWh"})
        diario["Energia_Total_MWh"] = diario["Energia_PPA_MWh"] + diario["Energia_Spot_MWh"]
        diario["Anio"] = diario["Fecha_Dia"].dt.year
        diario["Mes"] = diario["Fecha_Dia"].dt.month
        diarios.append(diario)

        eppa = df.loc[df["Tipo"] == "PPA", "Energia_MWh"].sum()
        espot = df.loc[df["Tipo"] == "SPOT", "Energia_MWh"].sum()
        ippa = df.loc[df["Tipo"] == "PPA", "Ingreso_USD"].sum()
        ispot = df.loc[df["Tipo"] == "SPOT", "Ingreso_USD"].sum()
        mensuales.append({"Anio":int(periodo[:4]), "Mes":int(periodo[4:6]), "Energia_PPA_MWh":eppa, "Energia_Spot_MWh":espot, "Energia_Total_MWh":eppa+espot, "Ingreso_PPA_USD":ippa, "Ingreso_Spot_USD":ispot, "Ingreso_Total_USD":ippa+ispot, "Fuente":"API"})

    df_m = pd.DataFrame(mensuales, columns=["Anio","Mes","Energia_PPA_MWh","Energia_Spot_MWh","Energia_Total_MWh","Ingreso_PPA_USD","Ingreso_Spot_USD","Ingreso_Total_USD","Fuente"])
    df_d = pd.concat(diarios, ignore_index=True) if diarios else pd.DataFrame(columns=["Fecha_Dia","Energia_PPA_MWh","Energia_Spot_MWh","Energia_Total_MWh","Anio","Mes"])
    return df_m, df_d, sorted(set(advertencias))

# =========================================================
# CARGA Y UNIONES
# =========================================================
try:
    df_budgets_bess = cargar_budgets_bess()
    df_consolidado_bess = cargar_consolidado_bess()
    df_budgets_proyectos = cargar_budgets_proyectos()
    df_consolidado_proyectos = cargar_consolidado_proyectos()
except Exception as error:
    st.error("No fue posible validar budgets.xlsx.")
    st.info(str(error))
    st.stop()

with st.spinner("Consultando meses pendientes de proyectos..."):
    df_proyectos_api, advertencias_proyectos = descargar_proyectos_pendientes(obtener_pendientes_proyectos(df_consolidado_proyectos))

fuentes_p = [df_consolidado_proyectos] + ([df_proyectos_api] if not df_proyectos_api.empty else [])
df_proyectos_final = pd.concat(fuentes_p, ignore_index=True)
df_proyectos_final["Prioridad"] = df_proyectos_final["Fuente"].map({"API":1,"Excel consolidado":2}).fillna(2)
df_proyectos_final = df_proyectos_final.sort_values("Prioridad").drop_duplicates(["Anio","Mes","Proyecto"], keep="last").drop(columns="Prioridad").sort_values(["Proyecto","Anio","Mes"])

periodos_bess_excel = {f"{int(f.Anio):04d}{int(f.Mes):02d}" for f in df_consolidado_bess.itertuples()}
periodos_bess_pendientes = [p for p in PERIODOS_DISPONIBLES if p not in periodos_bess_excel]
with st.spinner("Consultando meses pendientes del BESS..."):
    df_bess_api, df_bess_diario_api, advertencias_bess = procesar_bess_api(periodos_bess_pendientes)

fuentes_b = [df_consolidado_bess] + ([df_bess_api] if not df_bess_api.empty else [])
df_bess_final = pd.concat(fuentes_b, ignore_index=True)
df_bess_final["Prioridad"] = df_bess_final["Fuente"].map({"API":1,"Excel consolidado":2}).fillna(2)
df_bess_final = df_bess_final.sort_values("Prioridad").drop_duplicates(["Anio","Mes"], keep="last").drop(columns="Prioridad").sort_values(["Anio","Mes"])

advertencias = sorted(set(advertencias_proyectos + advertencias_bess))
if advertencias:
    st.warning("Algunas mediciones no estuvieron disponibles. El dashboard continuará con la información válida.")
    with st.expander("Ver advertencias de las API"):
        for a in advertencias:
            st.write(f"- {a}")

st.sidebar.header("Filtros")
anio_seleccionado = st.sidebar.selectbox("Año", [ANIO_OPERACIONAL], index=0)

# =========================================================
# BESS
# =========================================================
st.header("BESS María Elena")
st.subheader("Distribución diaria de energía del BESS")
if df_bess_diario_api.empty:
    st.info("No hay información diaria disponible desde la API. Los meses del Excel contienen información mensual.")
else:
    opciones = [(int(f.Anio),int(f.Mes)) for f in df_bess_diario_api[["Anio","Mes"]].drop_duplicates().sort_values(["Anio","Mes"], ascending=False).itertuples()]
    ad, md = st.selectbox("Mes para gráfico diario del BESS", opciones, format_func=lambda x:f"{MESES_COMPLETOS[x[1]]} {x[0]}", key="periodo_diario_bess")
    diario_mes = df_bess_diario_api[(df_bess_diario_api["Anio"]==ad)&(df_bess_diario_api["Mes"]==md)].copy()
    primer = pd.Timestamp(ad,md,1)
    ultimo = pd.Timestamp(ad,md,calendar.monthrange(ad,md)[1])
    hoy_n = pd.Timestamp.now().normalize()
    if ad == hoy_n.year and md == hoy_n.month:
        ultimo = min(ultimo,hoy_n)
    diario_mes = pd.DataFrame({"Fecha_Dia":pd.date_range(primer,ultimo,freq="D")}).merge(diario_mes[["Fecha_Dia","Energia_PPA_MWh","Energia_Spot_MWh","Energia_Total_MWh"]], on="Fecha_Dia", how="left").fillna(0)
    ancho = 0.40*24*60*60*1000
    fig_d = go.Figure()
    fig_d.add_bar(x=diario_mes["Fecha_Dia"], y=diario_mes["Energia_PPA_MWh"], name="PPA real", marker_color="#1565C0", width=ancho)
    fig_d.add_bar(x=diario_mes["Fecha_Dia"], y=diario_mes["Energia_Spot_MWh"], name="Spot real", marker_color="#65BDEB", width=ancho)
    fig_d.update_layout(
    title=(
    f"Inyección diaria BESS, "
    f"{MESES_COMPLETOS[md]} {ad}"
    ),
    barmode="stack",
    xaxis_title="Día",
    yaxis_title="Energía [MWh]",
    hovermode="x unified",
    height=550,
    legend=dict(
    orientation="h",
    yanchor="top",
    y=-0.22,
    xanchor="center",
    x=0.5
    ),
    margin=dict(
    l=60,
    r=30,
    t=70,
    b=110
    )
    )
    fig_d.update_xaxes(dtick=24*60*60*1000, tickformat="%d", range=[primer-pd.Timedelta(hours=12),ultimo+pd.Timedelta(hours=12)], showgrid=False)
    fig_d.update_yaxes(
    range=[0, 120],
    tickmode="linear",
    tick0=0,
    dtick=20,
    title_text="Energía [MWh]",
    gridcolor="rgba(128, 128, 128, 0.20)",
    zeroline=True,
    zerolinecolor="rgba(80, 80, 80, 0.50)",
    zerolinewidth=1
    )
    st.plotly_chart(fig_d, use_container_width=True, key="grafico_diario_bess")

bess_anio = df_bess_final[df_bess_final["Anio"]==anio_seleccionado].copy()
if not bess_anio.empty:
    meses_bess = sorted(bess_anio["Mes"].astype(int).unique(), reverse=True)
    mes_bess = st.selectbox("Mes para indicadores BESS", meses_bess, format_func=lambda m:MESES_COMPLETOS[m])
    fila = bess_anio[bess_anio["Mes"]==mes_bess].iloc[0]
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Inyección total del mes", f"{fila['Energia_Total_MWh']:,.1f} MWh")
    c2.metric("Energía PPA del mes", f"{fila['Energia_PPA_MWh']:,.1f} MWh")
    c3.metric("Energía Spot del mes", f"{fila['Energia_Spot_MWh']:,.1f} MWh")
    c4.metric("Ingreso del mes", f"USD {fila['Ingreso_Total_USD']:,.0f}")
    st.caption(f"Fuente: {fila['Fuente']}. Precio PPA: USD {PRECIO_PPA:,.2f}/MWh.")

con_b = pd.DataFrame({"Mes":range(1,13)}).merge(bess_anio,on="Mes",how="left")
bud_b = df_budgets_bess[df_budgets_bess["Anio"]==anio_seleccionado]
con_b = con_b.merge(bud_b[["Mes","Budget_Generacion_MWh","Budget_PPA_MWh"]],on="Mes",how="left")
for c in ["Energia_PPA_MWh","Energia_Spot_MWh","Energia_Total_MWh","Ingreso_Total_USD"]:
    if c not in con_b.columns:
        con_b[c]=0
    con_b[c]=con_b[c].fillna(0)
con_b["Nombre_Mes"]=con_b["Mes"].map(MESES_CORTOS)
fig_b=go.Figure()
fig_b.add_bar(x=con_b["Nombre_Mes"],y=con_b["Energia_PPA_MWh"],name="PPA real",marker_color="#1565C0",width=0.58)
fig_b.add_bar(x=con_b["Nombre_Mes"],y=con_b["Energia_Spot_MWh"],name="Spot real",marker_color="#65BDEB",width=0.58)
fig_b.add_scatter(x=con_b["Nombre_Mes"],y=con_b["Budget_Generacion_MWh"],name="Budget generación",mode="lines+markers",line=dict(color="#0B3D91",width=3))
fig_b.add_scatter(x=con_b["Nombre_Mes"],y=con_b["Budget_PPA_MWh"],name="Budget PPA",mode="lines+markers",line=dict(color="#2EAD5B",width=3))
fig_b.update_layout(title="Consolidado anual BESS",barmode="stack",xaxis_title="Mes",yaxis_title="Energía [MWh]",hovermode="x unified",height=560,legend=dict(orientation="h",y=1.02))
st.plotly_chart(fig_b,use_container_width=True)

with st.expander("Ver acumulados del BESS"):
    a1,a2,a3,a4=st.columns(4)
    a1.metric("Inyección acumulada",f"{bess_anio['Energia_Total_MWh'].sum():,.1f} MWh")
    a2.metric("PPA acumulado",f"{bess_anio['Energia_PPA_MWh'].sum():,.1f} MWh")
    a3.metric("Spot acumulado",f"{bess_anio['Energia_Spot_MWh'].sum():,.1f} MWh")
    a4.metric("Ingreso acumulado",f"USD {bess_anio['Ingreso_Total_USD'].sum():,.0f}")

# =========================================================
# PROYECTOS Y CONSOLIDADO PMGDS
# =========================================================
st.header("Generación mensual de proyectos")
st.caption("Cada gráfico utiliza el Excel para los meses consolidados y la API para los meses pendientes.")
proyectos_anio=df_proyectos_final[df_proyectos_final["Anio"]==anio_seleccionado].copy()
budgets_proyectos_anio=df_budgets_proyectos[df_budgets_proyectos["Anio"]==anio_seleccionado].copy()
columnas=st.columns(2)

for i, elemento in enumerate(ORDEN_GRAFICOS):
    if elemento == "CONSOLIDADO PMGDS":
        datos = proyectos_anio[proyectos_anio["Proyecto"].isin(PROYECTOS_PMGD)].groupby("Mes",as_index=False).agg(Generacion_MWh=("Generacion_MWh","sum"))
        bud = budgets_proyectos_anio[budgets_proyectos_anio["Proyecto"].isin(PROYECTOS_PMGD)].groupby("Mes",as_index=False).agg(Budget_Generacion_MWh=("Budget_Generacion_MWh","sum"))
        titulo = "Consolidado PMGDs"
        color_barra = "#1976D2"
    else:
        datos = proyectos_anio[proyectos_anio["Proyecto"]==elemento][["Mes","Generacion_MWh"]]
        bud = budgets_proyectos_anio[budgets_proyectos_anio["Proyecto"]==elemento][["Mes","Budget_Generacion_MWh"]]
        titulo = elemento
        color_barra = "#2176C7"

    con=pd.DataFrame({"Mes":range(1,13)}).merge(datos,on="Mes",how="left").merge(bud,on="Mes",how="left")
    con["Generacion_MWh"]=con["Generacion_MWh"].fillna(0)
    con["Nombre_Mes"]=con["Mes"].map(MESES_CORTOS)
    fig=go.Figure()
    fig.add_bar(x=con["Nombre_Mes"],y=con["Generacion_MWh"],name="Generación",marker_color=color_barra,width=0.52)
    fig.add_scatter(x=con["Nombre_Mes"],y=con["Budget_Generacion_MWh"],name="Budget",mode="lines+markers",line=dict(color="#F05A28",width=2),marker=dict(size=5))
    fig.update_layout(title=titulo,height=330,margin=dict(l=35,r=15,t=50,b=35),showlegend=(i==0),legend=dict(orientation="h",y=1.02),hovermode="x unified",yaxis_title="MWh",bargap=0.40)
    fig.update_yaxes(rangemode="tozero",gridcolor="rgba(140,140,140,0.18)")
    clave=quitar_tildes(elemento).lower().replace(" ","_")
    with columnas[i%2]:
        st.plotly_chart(fig,use_container_width=True,key=f"grafico_{clave}")

# =========================================================
# TOTAL PROYECTOS
# =========================================================
gen_total=proyectos_anio.groupby("Mes",as_index=False).agg(Generacion_Total_MWh=("Generacion_MWh","sum"))
bud_total=budgets_proyectos_anio.groupby("Mes",as_index=False).agg(Budget_Total_MWh=("Budget_Generacion_MWh","sum"))
total=pd.DataFrame({"Mes":range(1,13)}).merge(gen_total,on="Mes",how="left").merge(bud_total,on="Mes",how="left")
total["Generacion_Total_MWh"]=total["Generacion_Total_MWh"].fillna(0)
total["Nombre_Mes"]=total["Mes"].map(MESES_CORTOS)
st.subheader("Generación mensual total de proyectos")
fig_t=go.Figure()
fig_t.add_bar(x=total["Nombre_Mes"],y=total["Generacion_Total_MWh"],name="Generación real total",marker_color="#1565C0",width=0.38,offsetgroup="real")
fig_t.add_bar(x=total["Nombre_Mes"],y=total["Budget_Total_MWh"],name="Budget total",marker_color="#E58A21",width=0.38,offsetgroup="budget")
fig_t.update_layout(barmode="group",xaxis_title="Mes",yaxis_title="Generación [MWh]",height=520,hovermode="x unified",legend=dict(orientation="h",y=1.02))
st.plotly_chart(fig_t,use_container_width=True)

meses_con_gen=total[total["Generacion_Total_MWh"]>0]
if not meses_con_gen.empty:
    ultimo_mes=int(meses_con_gen["Mes"].max())
    fila=total[total["Mes"]==ultimo_mes].iloc[0]
    gen_acum=total.loc[total["Mes"]<=ultimo_mes,"Generacion_Total_MWh"].sum()
    bud_acum=total.loc[total["Mes"]<=ultimo_mes,"Budget_Total_MWh"].fillna(0).sum()
    t1,t2,t3,t4=st.columns(4)
    t1.metric(f"Generación total {MESES_COMPLETOS[ultimo_mes]}",f"{fila['Generacion_Total_MWh']:,.1f} MWh")
    t2.metric(f"Budget total {MESES_COMPLETOS[ultimo_mes]}",f"{fila['Budget_Total_MWh']:,.1f} MWh" if pd.notna(fila["Budget_Total_MWh"]) else "Sin budget")
    t3.metric("Generación acumulada",f"{gen_acum:,.1f} MWh")
    t4.metric("Diferencia acumulada vs budget",f"{gen_acum-bud_acum:,.1f} MWh")

with st.expander("Ver fuentes utilizadas por proyecto"):
    tabla=proyectos_anio[["Proyecto","Mes","Fuente"]].sort_values(["Proyecto","Mes"]).copy()
    tabla["Mes"]=tabla["Mes"].map(MESES_COMPLETOS)
    st.dataframe(tabla,use_container_width=True,hide_index=True)

st.caption(f"Dashboard procesado: {datetime.now():%d-%m-%Y %H:%M:%S}")
st.caption("Las consultas a las API y la caché se renuevan cada 12 horas.")
