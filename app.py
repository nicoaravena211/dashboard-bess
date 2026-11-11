"""
Dashboard de operación · Atlantica
BESS María Elena y portafolio de generación.

Qué hace:
  - Meses cerrados: se leen de budgets.xlsx (no se consultan a la API).
  - Meses pendientes y últimos 30 días: se consultan a las API del Coordinador (en paralelo y con caché).
  - PRMTE horario con CMG (preliminar o definitivo), consolidados mensual y anual en Excel.
  - Reporte semanal (lunes a viernes) en PDF para San Pedro III, Doña Carmen y María Elena.

Secrets requeridos (Manage app > Settings > Secrets): API_KEY_PRMTE y API_KEY_CMG.
"""
import calendar
import io
import threading
import time
import unicodedata
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import date, timedelta
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402
import plotly.graph_objects as go  # noqa: E402
import requests  # noqa: E402
import streamlit as st  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402
from reportlab.lib import colors  # noqa: E402
from reportlab.lib.units import mm  # noqa: E402
from reportlab.lib.utils import ImageReader  # noqa: E402
from reportlab.pdfgen import canvas as rl_canvas  # noqa: E402
from reportlab.platypus import Table, TableStyle  # noqa: E402

# =========================================================
# CONFIGURACION GENERAL
# =========================================================
st.set_page_config(page_title="Dashboard de operación · Atlantica", page_icon="⚡", layout="wide",
                   initial_sidebar_state="expanded")

RAIZ = Path(__file__).resolve().parent
LOGO = RAIZ / "logo_atlantica.png"
ARCHIVO_EXCEL = RAIZ / "budgets.xlsx"

COLOR_PRIMARIO, COLOR_SECUNDARIO = "#1565C0", "#65BDEB"
COLOR_BUDGET, COLOR_BUDGET_PPA, COLOR_BUDGET_PROYECTO = "#0B3D91", "#2EAD5B", "#F05A28"
COLOR_TEXTO, COLOR_TEXTO_2, COLOR_BORDE = "#172B4D", "#5E6C84", "#DFE1E6"
COLOR_REJILLA = "rgba(94,108,132,0.18)"
COLOR_LOGO = (0, 96, 114)  # verde azulado del logo de Atlantica
FUENTE_WEB = "Gadugi, 'Segoe UI', Arial, sans-serif"
FUENTE_EXCEL = "Gadugi"  # cambia a "Calibri" si prefieres la fuente anterior de los Excel

TZ = "America/Santiago"
ANIO_OPERACIONAL = 2026
PRECIO_PPA = 83.49  # USD/MWh
TIPO_CMG_DASHBOARD = "PRELIMINAR"  # tipo de CMG para valorizar el Spot del BESS en el dashboard
BARRA_CMG_BESS = "M.ELENA_______220"
UMBRAL_OFFSET_CARGA_BESS_KWH = 450  # carga horaria menor a este valor = offset -> 0

TTL_DATOS = 6 * 3600          # los datos de API se refrescan como máximo cada 6 h (o con el botón Actualizar)
TTL_MES_CERRADO = 24 * 3600   # meses anteriores al actual cambian poco
TTL_ERROR = 300               # una consulta fallida se reintenta a los 5 min
MAX_HILOS = 6                 # consultas simultáneas a la API
SEGUNDOS_ROTACION_GRAFICOS = 10
DIAS_CARRUSEL = 30

ARCHIVO_HOJAS = {"bess_bud": "Budgets", "bess_con": "Consolidado", "proy_bud": "Budgets_Proyectos", "proy_con": "Consolidado_Proyectos"}
URL_PRMTE = "https://medidas.api.coordinador.cl/medidas-v2/measurement"
URL_CMG = "https://sipub.api.coordinador.cl/costo-marginal-real/v4/findByDate"

PROYECTOS = {
    "MARIA ELENA PFV": [{"mpid": "MARELENA_220_JT1_GSS", "canal": 3}],
    "SAN PEDRO III": [{"mpid": "SOLRJAMA_220_JT1_RUC", "canal": 3}, {"mpid": "SOLRJAMA_220_JT2_RUC", "canal": 3}],
    "DOÑA CARMEN": [{"mpid": "CDNCARMN_220_J1_ECM", "canal": 3}],
    "LA QUINTA": [{"mpid": "CABRERO_023_PMGD7_QTC", "canal": 3}],
    "LA PERLA": [{"mpid": "LSANGLES_013_PMGD7_LAP", "canal": 3}],
    "LA HUERTA": [{"mpid": "PARRONAL_013_PMGD2_HUE", "canal": 3}],
    "CHACAICO": [{"mpid": "LSANGLES_013_PMGD8_CCC", "canal": 3}],
    "SANCLEMENTE": [{"mpid": "SCLMENTE_013_PMGD4_CFL", "canal": 3}],
    "TRILALEO": [{"mpid": "CHOLGUAN_015_PMGD6_YTL", "canal": 3}],
    "COLLANCO": [{"mpid": "CNSTUCON_023_PMGD4_OAO", "canal": 3}],
}
BESS_MPIDS = ["MARELENA_023_E1_GSS", "MARELENA_023_E7_GSS", "MARELENA_023_E8_GSS", "MARELENA_023_E11_GSS"]
PROYECTOS_PMGD = ["LA QUINTA", "LA PERLA", "LA HUERTA", "CHACAICO", "SANCLEMENTE", "TRILALEO", "COLLANCO"]
ORDEN_GRAFICOS = ["MARIA ELENA PFV", "SAN PEDRO III", "DOÑA CARMEN", "CONSOLIDADO PMGDS", "LA QUINTA", "LA PERLA", "LA HUERTA",
                  "CHACAICO", "SANCLEMENTE", "TRILALEO", "COLLANCO"]
ORDEN_CARRUSEL = ["MARIA ELENA BESS"] + list(PROYECTOS)
ORDEN_MENSUAL = ["LA QUINTA", "LA PERLA", "LA HUERTA", "CHACAICO", "SANCLEMENTE", "TRILALEO", "COLLANCO", "DOÑA CARMEN",
                 "MARIA ELENA PFV", "SAN PEDRO III"]
ORDEN_ANUAL = ["MARIA ELENA BESS", "MARIA ELENA PFV", "SAN PEDRO III", "DOÑA CARMEN", "LA QUINTA", "LA PERLA", "LA HUERTA",
               "CHACAICO", "SANCLEMENTE", "TRILALEO", "COLLANCO"]
MESES_CORTOS = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]

# Barra de costo marginal por instalación (CMG horario del PRMTE y reporte semanal)
BARRAS_CMG = {"MARIA ELENA PFV": "M.ELENA_______220", "BESS MARIA ELENA": "M.ELENA_______220",
              "SAN PEDRO III": "SOLARJAMA_____220", "DOÑA CARMEN": "DONA_CARMEN___220"}
TIPOS_CMG = {"Definitivo": "DEFINITIVO", "Preliminar": "PRELIMINAR"}

# Reporte semanal: una ficha por proyecto
REPORTES = [
    {"id": "SAN PEDRO III", "titulo": "San Pedro III", "col": "Inyección kWh",
     "nota": "Inyección en el punto de medida de San Pedro III (JT1 + JT2)."},
    {"id": "DOÑA CARMEN", "titulo": "Doña Carmen", "col": "Inyección kWh",
     "nota": "Inyección en el punto de medida de Doña Carmen."},
    {"id": "MARIA ELENA PFV", "titulo": "María Elena PFV", "col": "Neto PV Red kWh",
     "nota": "Inyección neta a la red: se descuentan los tramos de 15 min en que el BESS inyecta."},
]

# Credenciales
try:
    API_KEY_PRMTE = st.secrets["API_KEY_PRMTE"]
    API_KEY_CMG = st.secrets["API_KEY_CMG"]
except (KeyError, FileNotFoundError):
    st.error("No se encontraron las credenciales API en Streamlit Secrets.")
    st.info("Configura API_KEY_PRMTE y API_KEY_CMG desde Manage app > Settings > Secrets.")
    st.stop()

st.markdown(f"""
<style>
html, body, .stApp, .stApp p, .stApp li, .stApp label, .stApp button, .stApp h1, .stApp h2, .stApp h3, .stApp h4,
.stApp [data-testid="stMetricLabel"], .stApp [data-testid="stMetricValue"], .stApp [data-baseweb="tab"] p,
.stApp [data-testid="stCaptionContainer"] {{ font-family:{FUENTE_WEB}; }}
.block-container {{ padding-top:1.6rem; padding-bottom:3rem; max-width:1600px; }}
h1 {{ color:{COLOR_TEXTO}; font-size:2.1rem !important; font-weight:700 !important; letter-spacing:-.02em; margin-bottom:.1rem; padding:0 !important; }}
h2, h3 {{ color:{COLOR_TEXTO}; font-weight:650 !important; }}
section[data-testid="stSidebar"] {{ background:#FFFFFF; border-right:1px solid {COLOR_BORDE}; }}
div[data-testid="stMetric"] {{ background:#FFFFFF; border:1px solid {COLOR_BORDE}; border-radius:12px; padding:14px 16px;
  box-shadow:0 2px 7px rgba(23,43,77,.06); min-height:108px; }}
div[data-testid="stMetricLabel"] p {{ color:{COLOR_TEXTO_2}; font-size:.86rem; font-weight:600; }}
div[data-testid="stMetricValue"] {{ color:{COLOR_TEXTO}; font-size:1.45rem; font-weight:700; }}
div[data-testid="stPlotlyChart"] {{ background:#FFFFFF; border:1px solid {COLOR_BORDE}; border-radius:12px; padding:8px;
  box-shadow:0 2px 7px rgba(23,43,77,.05); }}
div[data-testid="stExpander"] {{ background:#FFFFFF; border:1px solid {COLOR_BORDE}; border-radius:10px; }}
div[data-testid="stVerticalBlockBorderWrapper"] {{ border-radius:12px; }}
button[data-baseweb="tab"] {{ font-weight:600; padding:10px 18px; }}
div[data-baseweb="tab-list"] {{ gap:4px; border-bottom:1px solid {COLOR_BORDE}; }}
.chip {{ background:#EAF2FD; color:{COLOR_BUDGET}; border:1px solid #B7D4F4; border-radius:10px; padding:9px 14px;
  text-align:center; font-size:.84rem; font-weight:600; }}
.chip.aviso {{ background:#FFF4E5; color:#8A4B00; border-color:#FFD59A; }}
#MainMenu, footer {{ visibility:hidden; }}
</style>
""", unsafe_allow_html=True)

if LOGO.exists():
    st.logo(str(LOGO), size="large")

# =========================================================
# UTILIDADES
# =========================================================
def ahora():
    """Hora actual de Chile, sin zona horaria (el servidor de Streamlit usa UTC)."""
    return pd.Timestamp.now(tz=TZ).tz_localize(None)


def fmt(valor, decimales=1):
    """Número con formato es-CL: 1.234,5"""
    if valor is None or pd.isna(valor):
        return "–"
    return f"{valor:,.{decimales}f}".replace(",", "§").replace(".", ",").replace("§", ".")


def fmt_signo(valor, decimales=1):
    return ("+" if valor >= 0 else "-") + fmt(abs(valor), decimales)


def quitar_tildes(texto):
    return "".join(c for c in unicodedata.normalize("NFD", str(texto)) if unicodedata.category(c) != "Mn")


def slug(texto):
    return quitar_tildes(texto).lower().replace(" ", "_")


def normalizar_columnas_excel(df):
    eq = {"ano": "Anio", "anio": "Anio", "mes": "Mes", "proyecto": "Proyecto", "budget_generacion_mwh": "Budget_Generacion_MWh",
          "budget_ppa_mwh": "Budget_PPA_MWh", "energia_ppa_mwh": "Energia_PPA_MWh", "energia_spot_mwh": "Energia_Spot_MWh",
          "ingreso_ppa_usd": "Ingreso_PPA_USD", "ingreso_spot_usd": "Ingreso_Spot_USD", "generacion_mwh": "Generacion_MWh"}

    def clave(nombre):
        n = quitar_tildes(nombre).strip().replace(" ", "_").replace("-", "_")
        while "__" in n:
            n = n.replace("__", "_")
        return n.lower()
    return df.rename(columns={c: eq.get(clave(c), str(c).strip()) for c in df.columns})


def convertir_numerico(serie):
    if pd.api.types.is_numeric_dtype(serie):
        return pd.to_numeric(serie, errors="coerce")
    return pd.to_numeric(serie.astype(str).str.strip().str.replace(" ", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")


def parse_fecha(serie):
    return pd.to_datetime(serie.astype(str).str[:19], errors="coerce")


def periodos_de_rango(inicio, fin):
    """Meses YYYYMM que cubren un rango de fechas."""
    return [p.strftime("%Y%m") for p in pd.period_range(pd.Timestamp(inicio), pd.Timestamp(fin), freq="M")]


def periodo_a_texto(periodo):
    return f"{MESES[int(periodo[4:6]) - 1]} {int(periodo[:4])}"


# Rangos de fechas de la sesión
HOY = ahora().normalize()
FECHA_FIN = HOY.strftime("%Y-%m") if HOY.year == ANIO_OPERACIONAL else (f"{ANIO_OPERACIONAL}-12" if HOY.year > ANIO_OPERACIONAL else f"{ANIO_OPERACIONAL}-01")
PERIODOS_ANIO = periodos_de_rango(f"{ANIO_OPERACIONAL}-01-01", pd.Timestamp(f"{FECHA_FIN}-01"))
CARRUSEL_FIN = HOY
CARRUSEL_INI = HOY - pd.Timedelta(days=DIAS_CARRUSEL - 1)
PERIODOS_CARRUSEL = periodos_de_rango(CARRUSEL_INI, CARRUSEL_FIN)

# =========================================================
# CAPA API (en paralelo y con caché compartida entre usuarios)
# =========================================================
@st.cache_resource
def _almacen():
    return {"lock": threading.Lock(), "datos": {}, "ultima": None}


ALM = _almacen()


def _con_cache(clave, ttl, funcion):
    """Caché en memoria, segura para hilos. Los errores se guardan poco tiempo."""
    with ALM["lock"]:
        e = ALM["datos"].get(clave)
        if e and time.time() - e[0] < e[2]:
            return e[1]
    resultado = funcion()
    ok = resultado[0] is not None
    with ALM["lock"]:
        ALM["datos"][clave] = (time.time(), resultado, ttl if ok else TTL_ERROR)
        if ok:
            ALM["ultima"] = time.time()
    return resultado


def _ttl_periodo(periodo):
    return TTL_DATOS if periodo >= ahora().strftime("%Y%m") else TTL_MES_CERRADO


def _pedir_prmte(periodo, mpid, canal):
    params = {"channelId": canal, "measurePointId": mpid, "period": periodo + "012345", "user_key": API_KEY_PRMTE}
    error = None
    for intento in range(3):
        try:
            r = requests.get(URL_PRMTE, params=params, timeout=60)
            if r.status_code == 200:
                data = r.json()
                if isinstance(data, list) and data:
                    return data, None
                error = "respuesta sin registros"
            elif r.status_code == 429:
                error = "límite temporal de solicitudes (HTTP 429)"
            elif r.status_code >= 500:
                error = f"error temporal del servidor (HTTP {r.status_code})"
            else:
                error = f"respuesta HTTP {r.status_code}"
        except requests.Timeout:
            error = "tiempo de espera agotado"
        except requests.ConnectionError:
            error = "error de conexión"
        except requests.RequestException as e:
            error = f"error de comunicación ({type(e).__name__})"
        except ValueError:
            error = "respuesta JSON inválida"
        if intento < 2:
            time.sleep(3 * (intento + 1))
    return None, error


def prmte(periodo, mpid, canal):
    return _con_cache(("prmte", periodo, mpid, canal), _ttl_periodo(periodo), lambda: _pedir_prmte(periodo, mpid, canal))


def _pedir_cmg(periodo, barra, tipo):
    """CMG (cada 15 min) de una barra para un mes. Se pide página a página hasta que venga vacía."""
    anio, mes = int(periodo[:4]), int(periodo[4:6])
    inicio = pd.Timestamp(anio, mes, 1)
    fin = min(pd.Timestamp(anio, mes, calendar.monthrange(anio, mes)[1]), ahora().normalize())
    if fin < inicio:
        return None, "mes futuro"
    registros, pagina = [], 0
    while pagina < 100:
        params = {"startDate": inicio.strftime("%Y-%m-%d"), "endDate": fin.strftime("%Y-%m-%d"), "page": pagina, "limit": 5000,
                  "type": tipo, "bar_transf": barra, "user_key": API_KEY_CMG}
        try:
            r = requests.get(URL_CMG, params=params, timeout=60)
            if r.status_code != 200:
                return None, f"respuesta HTTP {r.status_code}"
            datos = r.json().get("data", [])
        except (requests.RequestException, ValueError) as e:
            return None, f"error de comunicación ({type(e).__name__})"
        if not datos:
            break
        registros.extend(datos)
        pagina += 1
    if not registros:
        return None, "sin registros"
    df = pd.DataFrame(registros)
    if not {"fecha", "hra", "min", "cmg_usd_mwh_"}.issubset(df.columns):
        return None, "formato inesperado"
    df["Fecha"] = (pd.to_datetime(df["fecha"], errors="coerce") + pd.to_timedelta(convertir_numerico(df["hra"]).fillna(0), unit="h")
                   + pd.to_timedelta(convertir_numerico(df["min"]).fillna(0), unit="m"))
    df["CMG_USD_MWh"] = convertir_numerico(df["cmg_usd_mwh_"])
    df = df.dropna(subset=["Fecha", "CMG_USD_MWh"]).groupby("Fecha", as_index=False).agg(CMG_USD_MWh=("CMG_USD_MWh", "mean"))
    return (df, None) if not df.empty else (None, "sin registros")


def cmg(periodo, barra, tipo):
    return _con_cache(("cmg", periodo, barra, tipo), _ttl_periodo(periodo), lambda: _pedir_cmg(periodo, barra, tipo))


def lote(funcion, items):
    """Ejecuta funcion(*item) para cada item en paralelo. Devuelve {item: resultado}."""
    items = list(dict.fromkeys(items))
    if not items:
        return {}
    with ThreadPoolExecutor(MAX_HILOS) as ex:
        return dict(zip(items, ex.map(lambda it: funcion(*it), items)))


def extraer_mediciones(datos, canal):
    campo = f"channel{canal}"
    return [{"Fecha": m.get("dateRange"), "kWh": m.get(campo)} for bloque in (datos or []) for m in bloque.get("measurement", [])]


def serie_15m(brutos, configuraciones, periodos, canal, columna, etiqueta, advertencias, inicio=None, fin=None):
    """Suma por fecha los puntos de medida de una instalación. Devuelve (DataFrame, completo)."""
    registros, completo = [], True
    for conf in configuraciones:
        c = conf.get("canal", canal)
        for per in periodos:
            datos, error = brutos[(per, conf["mpid"], c)]
            med = extraer_mediciones(datos, c)
            if datos is None or not med:
                completo = False
                advertencias.append(f"{etiqueta}, {periodo_a_texto(per)}, {conf['mpid']}: {error or 'sin mediciones'}")
                continue
            registros.extend(med)
    if not registros:
        return pd.DataFrame(columns=["Fecha", columna]), False
    df = pd.DataFrame(registros)
    df["Fecha"] = parse_fecha(df["Fecha"])
    df[columna] = convertir_numerico(df["kWh"]).fillna(0)
    df = df.dropna(subset=["Fecha"])
    if inicio is not None:
        df = df[(df["Fecha"] >= pd.Timestamp(inicio)) & (df["Fecha"] < pd.Timestamp(fin) + pd.Timedelta(days=1))]
    if df.empty:
        return pd.DataFrame(columns=["Fecha", columna]), False
    return df.groupby("Fecha", as_index=False)[columna].sum().sort_values("Fecha"), completo


# =========================================================
# LECTURA DE budgets.xlsx (meses cerrados y budgets)
# =========================================================
def _leer_hoja(hoja, requeridas):
    if not ARCHIVO_EXCEL.exists():
        raise FileNotFoundError(f"No se encontró {ARCHIVO_EXCEL.name} junto a app.py.")
    df = normalizar_columnas_excel(pd.read_excel(ARCHIVO_EXCEL, sheet_name=hoja, engine="openpyxl"))
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(f"Faltan columnas en {hoja}: " + ", ".join(faltantes))
    df = df[requeridas].copy()
    for c in requeridas:
        if c != "Proyecto":
            df[c] = convertir_numerico(df[c])
    if "Proyecto" in df:
        df["Proyecto"] = df["Proyecto"].astype(str).str.strip().str.upper()
    df = df.dropna(subset=requeridas)
    df[["Anio", "Mes"]] = df[["Anio", "Mes"]].astype(int)
    return df


@st.cache_data(ttl=TTL_DATOS, show_spinner=False)
def cargar_excel(mtime):
    """mtime invalida la caché cuando se sube un budgets.xlsx nuevo."""
    bud_b = _leer_hoja(ARCHIVO_HOJAS["bess_bud"], ["Anio", "Mes", "Budget_Generacion_MWh", "Budget_PPA_MWh"])
    if bud_b.empty or not bud_b["Mes"].between(1, 12).all() or bud_b.duplicated(["Anio", "Mes"]).any():
        raise ValueError("La hoja Budgets no contiene información válida o tiene duplicados.")
    con_b = _leer_hoja(ARCHIVO_HOJAS["bess_con"], ["Anio", "Mes", "Energia_PPA_MWh", "Energia_Spot_MWh", "Ingreso_PPA_USD", "Ingreso_Spot_USD"])
    if con_b.duplicated(["Anio", "Mes"]).any():
        raise ValueError("Hay meses duplicados en Consolidado.")
    con_b["Energia_Total_MWh"] = con_b["Energia_PPA_MWh"] + con_b["Energia_Spot_MWh"]
    con_b["Ingreso_Total_USD"] = con_b["Ingreso_PPA_USD"] + con_b["Ingreso_Spot_USD"]
    con_b["Fuente"] = "Excel consolidado"
    bud_p = _leer_hoja(ARCHIVO_HOJAS["proy_bud"], ["Anio", "Mes", "Proyecto", "Budget_Generacion_MWh"])
    con_p = _leer_hoja(ARCHIVO_HOJAS["proy_con"], ["Anio", "Mes", "Proyecto", "Generacion_MWh"])
    for nombre, d in (("Budgets_Proyectos", bud_p), ("Consolidado_Proyectos", con_p)):
        desconocidos = sorted(set(d["Proyecto"]) - set(PROYECTOS))
        if desconocidos or d.duplicated(["Anio", "Mes", "Proyecto"]).any():
            raise ValueError(f"{nombre} contiene proyectos no reconocidos ({desconocidos}) o duplicados.")
    if (con_p["Generacion_MWh"] < 0).any():
        raise ValueError("Consolidado_Proyectos contiene generación negativa.")
    con_p["Fuente"] = "Excel consolidado"
    return bud_b, con_b, bud_p, con_p


# =========================================================
# MESES PENDIENTES Y CARRUSEL (API)
# =========================================================
@st.cache_data(ttl=TTL_DATOS, show_spinner=False)
def cargar_mensuales_api(pend_proyectos, pend_bess):
    """pend_proyectos: ((proyecto, periodo), ...)   pend_bess: (periodo, ...)"""
    adv = []
    items = {(per, c["mpid"], c["canal"]) for proy, per in pend_proyectos for c in PROYECTOS[proy]}
    items |= {(per, m, 3) for per in pend_bess for m in BESS_MPIDS}
    brutos = lote(prmte, items)
    cmgs = lote(cmg, [(per, BARRA_CMG_BESS, TIPO_CMG_DASHBOARD) for per in pend_bess])

    filas_p = []
    for proy, per in pend_proyectos:
        df, completo = serie_15m(brutos, PROYECTOS[proy], [per], 3, "kWh", proy, adv)
        if completo:
            filas_p.append({"Anio": int(per[:4]), "Mes": int(per[4:6]), "Proyecto": proy,
                            "Generacion_MWh": df["kWh"].abs().sum() / 1000, "Fuente": "API"})

    mensual_b, diarios = [], []
    for per in pend_bess:
        df, completo = serie_15m(brutos, [{"mpid": m, "canal": 3} for m in BESS_MPIDS], [per], 3, "kWh", "BESS", adv)
        if not completo:
            continue
        df["MWh"] = df["kWh"].abs() / 1000
        df["Tipo"] = "PPA"
        df.loc[(df["Fecha"].dt.hour >= 6) & (df["Fecha"].dt.hour < 21), "Tipo"] = "SPOT"
        df["Dia"] = df["Fecha"].dt.normalize()
        d = df.pivot_table(index="Dia", columns="Tipo", values="MWh", aggfunc="sum", fill_value=0).reset_index()
        d.columns.name = None
        for c in ("PPA", "SPOT"):
            if c not in d:
                d[c] = 0.0
        d = d.rename(columns={"Dia": "Fecha_Dia", "PPA": "Energia_PPA_MWh", "SPOT": "Energia_Spot_MWh"})
        d["Energia_Total_MWh"] = d["Energia_PPA_MWh"] + d["Energia_Spot_MWh"]
        d["Anio"], d["Mes"] = d["Fecha_Dia"].dt.year, d["Fecha_Dia"].dt.month
        diarios.append(d)

        df_cmg, err = cmgs[(per, BARRA_CMG_BESS, TIPO_CMG_DASHBOARD)]
        if df_cmg is None:
            adv.append(f"BESS, {periodo_a_texto(per)}: CMG no disponible ({err}); ingresos Spot no valorizados")
            df["CMG"] = 0.0
        else:
            df = df.merge(df_cmg.rename(columns={"CMG_USD_MWh": "CMG"}), on="Fecha", how="left")
            df["CMG"] = df["CMG"].fillna(0)
        df["Precio"] = PRECIO_PPA
        df.loc[df["Tipo"] == "SPOT", "Precio"] = df["CMG"]
        df["Ingreso"] = df["MWh"] * df["Precio"]
        e = df.groupby("Tipo")["MWh"].sum()
        i = df.groupby("Tipo")["Ingreso"].sum()
        ep, es, ip, isp = e.get("PPA", 0.0), e.get("SPOT", 0.0), i.get("PPA", 0.0), i.get("SPOT", 0.0)
        mensual_b.append({"Anio": int(per[:4]), "Mes": int(per[4:6]), "Energia_PPA_MWh": ep, "Energia_Spot_MWh": es,
                          "Energia_Total_MWh": ep + es, "Ingreso_PPA_USD": ip, "Ingreso_Spot_USD": isp,
                          "Ingreso_Total_USD": ip + isp, "Fuente": "API"})

    cols_p = ["Anio", "Mes", "Proyecto", "Generacion_MWh", "Fuente"]
    cols_b = ["Anio", "Mes", "Energia_PPA_MWh", "Energia_Spot_MWh", "Energia_Total_MWh", "Ingreso_PPA_USD", "Ingreso_Spot_USD",
              "Ingreso_Total_USD", "Fuente"]
    cols_d = ["Fecha_Dia", "Energia_PPA_MWh", "Energia_Spot_MWh", "Energia_Total_MWh", "Anio", "Mes"]
    return (pd.DataFrame(filas_p, columns=cols_p), pd.DataFrame(mensual_b, columns=cols_b),
            pd.concat(diarios, ignore_index=True)[cols_d] if diarios else pd.DataFrame(columns=cols_d), sorted(set(adv)))


@st.cache_data(ttl=TTL_DATOS, show_spinner=False)
def cargar_diarios_carrusel(periodos, inicio, fin):
    adv, resultados = [], []
    inicio, fin = pd.Timestamp(inicio), pd.Timestamp(fin)
    confs = {"MARIA ELENA BESS": [{"mpid": m, "canal": 3} for m in BESS_MPIDS], **PROYECTOS}
    brutos = lote(prmte, {(per, c["mpid"], c["canal"]) for conf in confs.values() for c in conf for per in periodos})
    for proyecto in ORDEN_CARRUSEL:
        partes = []
        for per in periodos:
            df, completo = serie_15m(brutos, confs[proyecto], [per], 3, "kWh", f"Carrusel {proyecto}", adv)
            if completo:
                partes.append(df)
        if not partes:
            adv.append(f"Carrusel {proyecto}: sin datos diarios entre {inicio:%d-%m-%Y} y {fin:%d-%m-%Y}")
            continue
        df = pd.concat(partes, ignore_index=True)
        df["Fecha_Dia"] = df["Fecha"].dt.normalize()
        d = (df.groupby("Fecha_Dia")["kWh"].apply(lambda s: s.abs().sum() / 1000).rename("Energia_MWh").reset_index())
        d = d[(d["Fecha_Dia"] >= inicio) & (d["Fecha_Dia"] <= fin)]
        if d.empty:
            adv.append(f"Carrusel {proyecto}: sin mediciones en los últimos {DIAS_CARRUSEL} días")
            continue
        d["Proyecto"] = proyecto
        resultados.append(d)
    cols = ["Fecha_Dia", "Proyecto", "Energia_MWh"]
    return (pd.concat(resultados, ignore_index=True)[cols] if resultados else pd.DataFrame(columns=cols)), sorted(set(adv))


# =========================================================
# SERIES HORARIAS CON CMG (PRMTE horario y reporte semanal)
# =========================================================
def a_horario(df, columna, indice, metodo="sum"):
    if df is None or df.empty:
        return pd.Series(float("nan"), index=indice)
    g = df.set_index("Fecha")[columna].resample("h")
    serie = g.sum(min_count=1) if metodo == "sum" else g.mean()
    return serie.reindex(indice)


def ingreso_horario(df_energia, columna, df_cmg, indice):
    """Ingreso spot horario = suma de (energía de 15 min x CMG de 15 min), como en el Colab."""
    if df_energia is None or df_energia.empty or df_cmg is None or df_cmg.empty:
        return pd.Series(float("nan"), index=indice)
    m = df_energia[["Fecha", columna]].merge(df_cmg, on="Fecha", how="inner")
    if m.empty:
        return pd.Series(float("nan"), index=indice)
    m["Ingreso"] = m[columna] / 1000 * m["CMG_USD_MWh"]
    return m.set_index("Fecha")["Ingreso"].resample("h").sum(min_count=1).reindex(indice)


@st.cache_data(ttl=TTL_DATOS, show_spinner=False)
def construir_horario(inicio, fin, tipo, parques):
    """
    Series horarias de los parques pedidos entre dos fechas (YYYY-MM-DD).
    Devuelve ({parque: DataFrame horario}, advertencias).
    Columnas: Fecha, Inyección kWh, [Neto PV Red kWh], [Carga BESS kWh], [CMG USD/MWh, Ingreso USD].
    """
    inicio, fin = pd.Timestamp(inicio).normalize(), pd.Timestamp(fin).normalize()
    periodos = periodos_de_rango(inicio, fin)
    indice = pd.date_range(inicio, fin + pd.Timedelta(hours=23), freq="h")
    codigo = TIPOS_CMG[tipo]
    adv = []

    bess_3 = [{"mpid": m, "canal": 3} for m in BESS_MPIDS]
    bess_1 = [{"mpid": m, "canal": 1} for m in BESS_MPIDS]
    con_bess = "BESS MARIA ELENA" in parques
    necesita_bess = con_bess or "MARIA ELENA PFV" in parques
    items = {(p, c["mpid"], c["canal"]) for n in parques if n in PROYECTOS for c in PROYECTOS[n] for p in periodos}
    if necesita_bess:
        items |= {(p, m, 3) for m in BESS_MPIDS for p in periodos}
    if con_bess:
        items |= {(p, m, 1) for m in BESS_MPIDS for p in periodos}
    brutos = lote(prmte, items)

    barras = sorted({BARRAS_CMG[n] for n in parques if n in BARRAS_CMG})
    cmgs = lote(cmg, [(p, b, codigo) for b in barras for p in periodos])
    cmg15 = {}
    for b in barras:
        partes = []
        for p in periodos:
            d, err = cmgs[(p, b, codigo)]
            if d is None:
                adv.append(f"CMG {tipo} {periodo_a_texto(p)} (barra {b.strip('_')}): {err}")
            else:
                partes.append(d)
        cmg15[b] = (pd.concat(partes).drop_duplicates("Fecha") if partes else pd.DataFrame(columns=["Fecha", "CMG_USD_MWh"]))

    bess_iny = bess_car = pd.DataFrame(columns=["Fecha", "x"])
    if necesita_bess:
        bess_iny, _ = serie_15m(brutos, bess_3, periodos, 3, "BESS Raw kWh", "BESS", adv, inicio, fin)
    if con_bess:
        bess_car, _ = serie_15m(brutos, bess_1, periodos, 1, "Carga BESS kWh", "BESS (carga)", adv, inicio, fin)

    resultado = {}
    for nombre in parques:
        if nombre == "BESS MARIA ELENA":
            df15 = bess_iny.rename(columns={"BESS Raw kWh": "Inyección kWh"})
        else:
            df15, _ = serie_15m(brutos, PROYECTOS[nombre], periodos, 3, "Inyección kWh", nombre, adv, inicio, fin)
        h = pd.DataFrame({"Fecha": indice})
        h["Inyección kWh"] = a_horario(df15, "Inyección kWh", indice).to_numpy()
        col_energia = "Inyección kWh"
        if nombre == "MARIA ELENA PFV":
            base = df15.merge(bess_iny, on="Fecha", how="left")
            base["BESS Raw kWh"] = base["BESS Raw kWh"].fillna(0)
            base["Neto PV Red kWh"] = base["Inyección kWh"]
            base.loc[base["BESS Raw kWh"] > 0, "Neto PV Red kWh"] = 0
            h["Neto PV Red kWh"] = a_horario(base, "Neto PV Red kWh", indice).to_numpy()
            df15, col_energia = base, "Neto PV Red kWh"
        if nombre == "BESS MARIA ELENA":
            carga = pd.to_numeric(a_horario(bess_car, "Carga BESS kWh", indice), errors="coerce")
            carga[carga < UMBRAL_OFFSET_CARGA_BESS_KWH] = 0
            h["Carga BESS kWh"] = carga.to_numpy()
        if nombre in BARRAS_CMG:
            barra = BARRAS_CMG[nombre]
            h["CMG USD/MWh"] = a_horario(cmg15[barra], "CMG_USD_MWh", indice, "mean").to_numpy()
            if nombre != "BESS MARIA ELENA":
                h["Ingreso USD"] = ingreso_horario(df15, col_energia, cmg15[barra], indice).to_numpy()
        resultado[nombre] = h
    return resultado, sorted(set(adv))


# =========================================================
# EXCEL
# =========================================================
def _fuente(**kw):
    return Font(name=FUENTE_EXCEL, **kw)


@st.cache_data(ttl=TTL_DATOS, show_spinner=False)
def generar_excel_prmte(inicio, fin, tipo):
    """PRMTE horario de todos los parques, con CMG horario (solo Doña Carmen, San Pedro III, María Elena y BESS)."""
    parques = tuple(PROYECTOS) + ("BESS MARIA ELENA",)
    horario, adv = construir_horario(inicio, fin, tipo, parques)
    col_cmg = f"CMG {tipo} USD/MWh"
    hojas, resumen = {}, []
    for nombre in parques:
        h = horario[nombre].drop(columns=["Ingreso USD"], errors="ignore").rename(columns={"CMG USD/MWh": col_cmg})
        hojas[nombre] = h
        resumen.append({"Parque": nombre, "Inyección MWh": h["Inyección kWh"].sum(min_count=1) / 1000,
                        "Neto PV Red MWh": h["Neto PV Red kWh"].sum(min_count=1) / 1000 if "Neto PV Red kWh" in h else None,
                        "Carga MWh": h["Carga BESS kWh"].sum(min_count=1) / 1000 if "Carga BESS kWh" in h else None})
    df_res = pd.DataFrame(resumen)
    df_res.insert(0, "Fecha inicio", inicio)
    df_res.insert(1, "Fecha fin", fin)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df_res.to_excel(w, sheet_name="Resumen", index=False)
        for nombre in parques:
            hojas[nombre].to_excel(w, sheet_name=nombre[:31], index=False)
        lado = Side(style="thin", color="DFE1E6")
        borde = Border(left=lado, right=lado, top=lado, bottom=lado)
        for nombre_hoja, ws in w.sheets.items():
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            ws.sheet_view.showGridLines = False
            for c in ws[1]:
                c.fill = PatternFill("solid", fgColor="1565C0")
                c.font = _fuente(size=11, bold=True, color="FFFFFF")
                c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[1].height = 22
            for fila in ws.iter_rows(min_row=2):
                for c in fila:
                    c.border = borde
                    c.font = _fuente(size=10)
                    if c.column == 1 and nombre_hoja != "Resumen":
                        c.number_format = "yyyy-mm-dd hh:mm"
                    elif isinstance(c.value, (int, float)):
                        c.number_format = "#,##0.000"
            for col in ws.columns:
                largo = max(len(str(c.value)) if c.value is not None else 0 for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(largo + 2, 14), 30)
            ws.page_setup.orientation = "landscape"
            ws.page_setup.fitToWidth = 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True
    return buf.getvalue(), adv


@st.cache_data(show_spinner=False)
def generar_excel_consolidado_mes(df_bess, df_proyectos, anio, mes):
    periodo = f"{anio:04d}-{mes:02d}"
    dp = df_proyectos[(df_proyectos["Anio"] == anio) & (df_proyectos["Mes"] == mes)].set_index("Proyecto")
    filas = [{"Mes": periodo, "Parque": p, "Total Inyección MWh": (dp.loc[p, "Generacion_MWh"] if p in dp.index else None),
              "Distribución contrato": None} for p in ORDEN_MENSUAL]
    b = df_bess[(df_bess["Anio"] == anio) & (df_bess["Mes"] == mes)]
    total = ppa = spot = None
    if not b.empty:
        total, ppa, spot = b.iloc[-1][["Energia_Total_MWh", "Energia_PPA_MWh", "Energia_Spot_MWh"]]
    filas.append({"Mes": periodo, "Parque": "BESS MARIA ELENA (PPA)", "Total Inyección MWh": total, "Distribución contrato": ppa})
    filas.append({"Mes": periodo, "Parque": "BESS MARIA ELENA (SPOT)", "Total Inyección MWh": None, "Distribución contrato": spot})
    nombre_hoja = f"Consolidado {MESES[mes - 1]}"[:31]
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(filas).to_excel(w, sheet_name=nombre_hoja, index=False)
        ws = w.sheets[nombre_hoja]
        punteado = Side(style="dotted", color="000000")
        borde = Border(left=punteado, right=punteado, top=punteado, bottom=punteado)
        for c in ws[1]:
            c.font = _fuente(size=11, bold=True)
            c.fill = PatternFill("solid", fgColor="FFFFFF")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = Border(bottom=Side(style="thin", color="000000"))
        ws.row_dimensions[1].height = 20
        fila_ppa = len(ORDEN_MENSUAL) + 2
        ws.merge_cells(start_row=fila_ppa, start_column=3, end_row=fila_ppa + 1, end_column=3)
        ws.cell(fila_ppa, 3).value = total
        for fila in range(2, ws.max_row + 1):
            for col in range(1, 5):
                c = ws.cell(fila, col)
                c.font = _fuente(size=11)
                c.border = borde
                c.alignment = Alignment(horizontal="left" if col <= 2 else "right", vertical="center")
                if col >= 3:
                    c.fill = PatternFill("solid", fgColor="EBF1DE")
                    c.number_format = "#,##0.0"
            ws.row_dimensions[fila].height = 19
        for letra, ancho in zip("ABCD", (13, 25, 22, 23)):
            ws.column_dimensions[letra].width = ancho
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:D{ws.max_row}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_area = f"A1:D{ws.max_row}"
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def generar_excel_consolidado_anual(df_bess, df_proyectos, anio):
    dp = df_proyectos[df_proyectos["Anio"] == anio][["Proyecto", "Mes", "Generacion_MWh"]]
    db = df_bess[df_bess["Anio"] == anio][["Mes", "Energia_Total_MWh"]].rename(columns={"Energia_Total_MWh": "Generacion_MWh"})
    db["Proyecto"] = "MARIA ELENA BESS"
    datos = pd.concat([db[["Proyecto", "Mes", "Generacion_MWh"]], dp], ignore_index=True)
    tabla = datos.pivot_table(index="Proyecto", columns="Mes", values="Generacion_MWh", aggfunc="sum").reindex(ORDEN_ANUAL).reindex(columns=range(1, 13))
    tabla.columns = MESES_CORTOS
    tabla["Inyección Anual MWh"] = tabla.sum(axis=1, min_count=1)
    tabla = tabla.reset_index().rename(columns={"Proyecto": "Parque"})
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        tabla.to_excel(w, sheet_name="Consolidado Anual", index=False, startrow=1)
        ws = w.sheets["Consolidado Anual"]
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=14)
        t = ws.cell(1, 1)
        t.value = f"CONSOLIDADO ANUAL {anio} [MWh]"
        t.fill = PatternFill("solid", fgColor="0B3D91")
        t.font = _fuente(size=13, bold=True, color="FFFFFF")
        t.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[1].height = 25
        lado = Side(style="thin", color="C9D1D9")
        borde = Border(left=lado, right=lado, top=lado, bottom=lado)
        for col in range(1, 15):
            c = ws.cell(2, col)
            c.font = _fuente(size=11, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1565C0")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = borde
        ws.row_dimensions[2].height = 22
        for fila in range(3, ws.max_row + 1):
            for col in range(1, 15):
                c = ws.cell(fila, col)
                c.border = borde
                c.font = _fuente(size=11, bold=(col == 14))
                c.alignment = Alignment(horizontal="left" if col == 1 else "right", vertical="center")
                if col >= 2:
                    c.number_format = "#,##0.0"
                if col == 14:
                    c.fill = PatternFill("solid", fgColor="EBF1DE")
                elif fila % 2 == 0:
                    c.fill = PatternFill("solid", fgColor="F5F7FA")
            ws.row_dimensions[fila].height = 20
        ft = ws.max_row + 1
        ws.cell(ft, 1).value = "TOTAL PORTAFOLIO"
        for col in range(2, 15):
            letra = get_column_letter(col)
            ws.cell(ft, col).value = f"=SUM({letra}3:{letra}{ft - 1})"
            ws.cell(ft, col).number_format = "#,##0.0"
        for col in range(1, 15):
            c = ws.cell(ft, col)
            c.fill = PatternFill("solid", fgColor="DCE6F1")
            c.font = _fuente(size=11, bold=True, color="172B4D")
            c.border = borde
            c.alignment = Alignment(horizontal="left" if col == 1 else "right", vertical="center")
        ws.column_dimensions["A"].width = 25
        for col in range(2, 14):
            ws.column_dimensions[get_column_letter(col)].width = 12
        ws.column_dimensions["N"].width = 22
        ws.freeze_panes = "B3"
        ws.sheet_view.showGridLines = False
        ws.auto_filter.ref = f"A2:N{ft}"
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.print_area = f"A1:N{ft}"
    return buf.getvalue()


# =========================================================
# REPORTE SEMANAL (PDF)
# =========================================================
def semanas_disponibles(anio):
    """Semanas ISO (lunes a viernes) del año hasta hoy: [(n, lunes, viernes)], la más reciente primero."""
    hoy, salida = HOY.date(), []
    for n in range(1, 54):
        try:
            lunes = date.fromisocalendar(anio, n, 1)
        except ValueError:
            break
        if lunes > hoy:
            break
        salida.append((n, lunes, lunes + timedelta(days=4)))
    return salida[::-1]


def calcular_semana(df, cfg, tipo, lunes):
    """Totales por día y de la semana. CMG ponderado = suma(energía x CMG) / suma(energía)."""
    col = cfg["col"]
    d = df.copy()
    d["Dia"] = d["Fecha"].dt.normalize()
    d["MWh"] = d[col] / 1000
    filas = []
    for i in range(5):
        dia = pd.Timestamp(lunes) + pd.Timedelta(days=i)
        x = d[(d["Dia"] == dia) & d["MWh"].notna()]
        xc = x[x["CMG USD/MWh"].notna()]
        filas.append({"fecha": dia, "mwh": x["MWh"].sum(), "mwh_cmg": xc["MWh"].sum(), "ingreso": xc["Ingreso USD"].sum(),
                      "h_energia": len(x), "h_cmg": len(xc)})
    tot = {k: sum(f[k] for f in filas) for k in ("mwh", "mwh_cmg", "ingreso", "h_energia", "h_cmg")}
    if tot["h_energia"] == 0:
        raise ValueError(f"Sin datos de inyección para {cfg['titulo']} en esa semana.")
    if tot["h_cmg"] == 0:
        sugerencia = ("El definitivo suele publicarse con retraso; prueba con Preliminar." if tipo == "Definitivo"
                      else "Prueba con Definitivo.")
        raise ValueError(f"No hay CMG {tipo} para {cfg['titulo']} en esa semana. {sugerencia}")
    return filas, tot


def ponderado(x):
    return x["ingreso"] / x["mwh_cmg"] if x["mwh_cmg"] > 0 else None


def _grafico_png(figura):
    buf = io.BytesIO()
    figura.savefig(buf, format="png", dpi=200, bbox_inches="tight", facecolor="white")
    plt.close(figura)
    buf.seek(0)
    return buf


def _estilo_ejes(ax):
    for lado in ("top", "right"):
        ax.spines[lado].set_visible(False)
    ax.spines["left"].set_color("#DFE1E6")
    ax.spines["bottom"].set_color("#DFE1E6")
    ax.tick_params(colors="#5E6C84", labelsize=8)
    ax.grid(axis="y", color="#DFE1E6", linewidth=.6)
    ax.set_axisbelow(True)


def crear_pdf_reporte(cfg, semana, tipo, df, filas, tot, lunes, viernes, generado):
    barra = BARRAS_CMG[cfg["id"]]
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=(210 * mm, 297 * mm))
    c.setTitle(f"Reporte semana {semana} · {cfg['titulo']}")
    W, mx = 210 * mm, 15 * mm
    ancho = W - 2 * mx
    teal = colors.Color(*[v / 255 for v in COLOR_LOGO])
    gris = colors.HexColor("#5E6C84")
    texto = colors.HexColor("#172B4D")

    # Encabezado
    if LOGO.exists():
        img = ImageReader(str(LOGO))
        iw, ih = img.getSize()
        lw = 46 * mm
        c.drawImage(img, mx, 297 * mm - 11 * mm - lw * ih / iw, width=lw, height=lw * ih / iw, mask="auto")
    c.setFillColor(teal)
    c.setFont("Helvetica-Bold", 21)
    c.drawRightString(W - mx, 297 * mm - 18 * mm, f"Reporte semana {semana}")
    c.setFont("Helvetica-Bold", 13)
    c.drawRightString(W - mx, 297 * mm - 25 * mm, cfg["titulo"])
    c.setFillColor(gris)
    c.setFont("Helvetica", 9)
    c.drawRightString(W - mx, 297 * mm - 31 * mm, f"Lunes {lunes:%d-%m-%Y} a viernes {viernes:%d-%m-%Y}  ·  CMG {tipo}")
    c.setStrokeColor(teal)
    c.setLineWidth(1.4)
    c.line(mx, 297 * mm - 35 * mm, W - mx, 297 * mm - 35 * mm)

    # Indicadores
    pond_tot = ponderado(tot)
    kpis = [("Generación inyectada", f"{fmt(tot['mwh'], 1)} MWh"), ("CMG ponderado", f"USD {fmt(pond_tot, 2)}/MWh"),
            ("Ingresos spot", f"USD {fmt(tot['ingreso'], 0)}")]
    kw = (ancho - 8 * mm) / 3
    y_k = 297 * mm - 59 * mm
    for i, (etiqueta, valor) in enumerate(kpis):
        x = mx + i * (kw + 4 * mm)
        c.setFillColor(colors.HexColor("#F5F7FA"))
        c.setStrokeColor(colors.HexColor("#DFE1E6"))
        c.setLineWidth(.6)
        c.roundRect(x, y_k, kw, 19 * mm, 2 * mm, fill=1, stroke=1)
        c.setFillColor(gris)
        c.setFont("Helvetica", 8.5)
        c.drawString(x + 4 * mm, y_k + 12.5 * mm, etiqueta)
        c.setFillColor(texto)
        c.setFont("Helvetica-Bold", 14)
        c.drawString(x + 4 * mm, y_k + 4.5 * mm, valor)

    # Tabla diaria
    datos = [["Día", "Inyección (MWh)", "CMG ponderado (USD/MWh)", "Ingreso spot (USD)"]]
    for i, f in enumerate(filas):
        datos.append([f"{DIAS_SEMANA[i]} {f['fecha']:%d-%m-%Y}", fmt(f["mwh"], 1) if f["h_energia"] else "–",
                      fmt(ponderado(f), 2) if f["h_cmg"] else "–", fmt(f["ingreso"], 0) if f["h_cmg"] else "–"])
    datos.append(["Total semana", fmt(tot["mwh"], 1), fmt(pond_tot, 2), fmt(tot["ingreso"], 0)])
    tabla = Table(datos, colWidths=[ancho * .34, ancho * .22, ancho * .26, ancho * .18], rowHeights=7 * mm)
    tabla.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 9), ("TEXTCOLOR", (0, 0), (-1, -1), texto),
        ("BACKGROUND", (0, 0), (-1, 0), teal), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (1, 0), (-1, -1), "RIGHT"), ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#DCE6F1")), ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), .5, colors.HexColor("#DFE1E6")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6)]))
    tw, th = tabla.wrapOn(c, ancho, 100 * mm)
    y = y_k - 5 * mm - th
    tabla.drawOn(c, mx, y)

    # Gráfico 1: inyección diaria
    etiquetas = [f"{DIAS_SEMANA[i][:3]} {f['fecha']:%d-%m}" for i, f in enumerate(filas)]
    fig, ax = plt.subplots(figsize=(10, 2.9))
    ax.bar(etiquetas, [f["mwh"] if f["h_energia"] else 0 for f in filas], color="#1565C0", width=.55)
    ax.set_title("Generación inyectada por día (MWh)", loc="left", fontsize=11, color="#172B4D", fontweight="bold")
    _estilo_ejes(ax)
    g1 = _grafico_png(fig)
    h1 = ancho * 0.29
    y -= 4 * mm + h1
    c.drawImage(ImageReader(g1), mx, y, width=ancho, height=h1)

    # Gráfico 2: inyección y CMG por hora
    sem = df[(df["Fecha"] >= pd.Timestamp(lunes)) & (df["Fecha"] < pd.Timestamp(viernes) + pd.Timedelta(days=1))].reset_index(drop=True)
    fig, ax = plt.subplots(figsize=(10, 3.2))
    ax.bar(range(len(sem)), sem[cfg["col"]] / 1000, color="#65BDEB", width=.9, label="Inyección horaria (MWh)")
    ax2 = ax.twinx()
    ax2.plot(range(len(sem)), sem["CMG USD/MWh"], color="#F05A28", linewidth=1.6, label=f"CMG {tipo} (USD/MWh)")
    marcas = list(range(0, len(sem), 12))
    ax.set_xticks(marcas)
    ax.set_xticklabels([sem["Fecha"].iloc[i].strftime("%d-%m %Hh") for i in marcas], fontsize=8)
    ax.set_ylabel("MWh", fontsize=8, color="#5E6C84")
    ax2.set_ylabel("USD/MWh", fontsize=8, color="#5E6C84")
    ax2.tick_params(colors="#5E6C84", labelsize=8)
    ax2.spines[["top"]].set_visible(False)
    _estilo_ejes(ax)
    ax.set_title("Inyección y CMG por hora", loc="left", fontsize=11, color="#172B4D", fontweight="bold")
    manijas = ax.get_legend_handles_labels()[0] + ax2.get_legend_handles_labels()[0]
    etiq = ax.get_legend_handles_labels()[1] + ax2.get_legend_handles_labels()[1]
    ax.legend(manijas, etiq, loc="upper center", bbox_to_anchor=(.5, -.12), ncol=2, frameon=False, fontsize=8)
    g2 = _grafico_png(fig)
    h2 = ancho * 0.33
    y -= 3 * mm + h2
    c.drawImage(ImageReader(g2), mx, y, width=ancho, height=h2)

    # Notas
    notas = [f"CMG ponderado = suma(energía x CMG) / suma(energía). Ingreso spot = energía de cada tramo de 15 min x CMG {tipo.lower()} de la barra {barra}.",
             cfg["nota"],
             f"Horas con medida de energía: {tot['h_energia']} de 120. Horas con CMG disponible: {tot['h_cmg']}." +
             (" Las horas sin CMG no se valorizan." if tot["h_cmg"] < tot["h_energia"] else ""),
             f"Reporte generado el {generado:%d-%m-%Y %H:%M} (hora de Chile)."]
    c.setFillColor(gris)
    c.setFont("Helvetica", 7.5)
    y -= 5 * mm
    for n in notas:
        c.drawString(mx, y, n)
        y -= 3.6 * mm
    c.showPage()
    c.save()
    return buf.getvalue()


def generar_reportes_semana(semana, lunes, viernes, tipo):
    """Devuelve ({archivo.pdf: bytes}, errores, advertencias_api)."""
    horario, adv = construir_horario(lunes.isoformat(), viernes.isoformat(), tipo, tuple(r["id"] for r in REPORTES))
    pdfs, errores = {}, []
    for cfg in REPORTES:
        try:
            filas, tot = calcular_semana(horario[cfg["id"]], cfg, tipo, lunes)
            pdfs[f"Reporte_Semana{semana}_{cfg['titulo'].replace(' ', '_')}.pdf"] = crear_pdf_reporte(
                cfg, semana, tipo, horario[cfg["id"]], filas, tot, lunes, viernes, ahora())
        except ValueError as e:
            errores.append(str(e))
    return pdfs, errores, adv


# =========================================================
# CARGA DE DATOS Y UNIONES
# =========================================================
try:
    df_bud_bess, df_con_bess, df_bud_proy, df_con_proy = cargar_excel(ARCHIVO_EXCEL.stat().st_mtime if ARCHIVO_EXCEL.exists() else 0)
except Exception as error:
    st.error("No fue posible validar budgets.xlsx.")
    st.info(str(error))
    st.stop()

existentes_p = {(r.Proyecto, f"{r.Anio:04d}{r.Mes:02d}") for r in df_con_proy.itertuples()}
pend_proy = tuple((p, per) for p in PROYECTOS for per in PERIODOS_ANIO if (p, per) not in existentes_p)
existentes_b = {f"{r.Anio:04d}{r.Mes:02d}" for r in df_con_bess.itertuples()}
pend_bess = tuple(per for per in PERIODOS_ANIO if per not in existentes_b)

with st.spinner("Consultando datos a las API del Coordinador…"):
    api_proy, api_bess, api_bess_diario, adv_mensual = cargar_mensuales_api(pend_proy, pend_bess)
    df_carrusel, adv_carrusel = cargar_diarios_carrusel(tuple(PERIODOS_CARRUSEL), CARRUSEL_INI.isoformat(), CARRUSEL_FIN.isoformat())


def unir(excel, api, claves):
    """El Excel consolidado tiene prioridad sobre la API si un mes existe en ambos."""
    df = pd.concat([excel] + ([api] if not api.empty else []), ignore_index=True)
    df["_p"] = df["Fuente"].map({"API": 1, "Excel consolidado": 2}).fillna(2)
    return df.sort_values("_p").drop_duplicates(claves, keep="last").drop(columns="_p").sort_values(claves)


df_proy = unir(df_con_proy, api_proy, ["Anio", "Mes", "Proyecto"])
df_bess = unir(df_con_bess, api_bess, ["Anio", "Mes"])
advertencias = sorted(set(adv_mensual + adv_carrusel))
st.session_state.setdefault("carga_ts", time.time())

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.header("Filtros")
anio = st.sidebar.selectbox("Año operacional", [ANIO_OPERACIONAL], index=0)
st.sidebar.divider()
st.sidebar.subheader("Datos")
ultima = ALM["ultima"]
st.sidebar.caption("Última consulta a las API: " + (pd.Timestamp(ultima, unit="s", tz="UTC").tz_convert(TZ).strftime("%d-%m-%Y %H:%M") if ultima else "–"))
if st.sidebar.button("Actualizar datos ahora", width="stretch", icon=":material/refresh:",
                     help="Vuelve a consultar las API del Coordinador. Puede tardar unos segundos."):
    with ALM["lock"]:
        ALM["datos"].clear()
    st.cache_data.clear()
    st.session_state.carga_ts = time.time()
    st.session_state.pop("prmte_archivo", None)
    st.session_state.pop("reportes", None)
    st.rerun()
st.sidebar.caption("Los meses cerrados vienen de budgets.xlsx. Las API se consultan en paralelo y se guardan en caché "
                   f"(se refrescan solas cada {TTL_DATOS // 3600} horas).")

# =========================================================
# ENCABEZADO Y RESUMEN
# =========================================================
col_t, col_e = st.columns([3.4, 1.2], vertical_alignment="center")
with col_t:
    st.title("Dashboard de operación")
    st.caption(f"BESS María Elena y portafolio de generación · Período {ANIO_OPERACIONAL}-01 a {FECHA_FIN}")
with col_e:
    if advertencias:
        st.markdown(f'<div class="chip aviso">⚠ {len(advertencias)} advertencias de datos</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="chip">● Datos completos</div>', unsafe_allow_html=True)
if advertencias:
    with st.expander("Detalle de advertencias"):
        st.warning("Algunas consultas no estuvieron disponibles. Los meses consolidados del Excel no se ven afectados.")
        for a in advertencias:
            st.markdown(f"- {a}")

bess_anio = df_bess[df_bess["Anio"] == anio].copy()
proy_anio = df_proy[df_proy["Anio"] == anio].copy()
bud_proy_anio = df_bud_proy[df_bud_proy["Anio"] == anio].copy()
bud_bess_anio = df_bud_bess[df_bud_bess["Anio"] == anio].copy()

gen_mes = proy_anio.groupby("Mes")["Generacion_MWh"].sum()
bud_mes = bud_proy_anio.groupby("Mes")["Budget_Generacion_MWh"].sum()
ultimo_mes_gen = int(gen_mes[gen_mes > 0].index.max()) if (gen_mes > 0).any() else 0
# Los indicadores comparan contra budget solo hasta el último mes cerrado (el del Excel), no contra un mes a medias
cierre_p = int(df_con_proy[df_con_proy["Anio"] == anio]["Mes"].max()) if (df_con_proy["Anio"] == anio).any() else 0
cierre_b = int(df_con_bess[df_con_bess["Anio"] == anio]["Mes"].max()) if (df_con_bess["Anio"] == anio).any() else 0
k1, k2, k3, k4 = st.columns(4)
if cierre_p:
    g_ac = gen_mes[gen_mes.index <= cierre_p].sum()
    b_ac = bud_mes[bud_mes.index <= cierre_p].sum()
    k1.metric(f"Generación (acum. a {MESES_CORTOS[cierre_p - 1]})", f"{fmt(g_ac)} MWh", f"{fmt_signo((g_ac / b_ac - 1) * 100 if b_ac else 0)} % vs budget")
if cierre_b:
    cerr_b = bess_anio[bess_anio["Mes"] <= cierre_b]
    bb_ac = bud_bess_anio[bud_bess_anio["Mes"] <= cierre_b]["Budget_Generacion_MWh"].sum()
    be_ac = cerr_b["Energia_Total_MWh"].sum()
    k2.metric(f"Inyección BESS (acum. a {MESES_CORTOS[cierre_b - 1]})", f"{fmt(be_ac)} MWh", f"{fmt_signo((be_ac / bb_ac - 1) * 100 if bb_ac else 0)} % vs budget")
    k3.metric(f"Ingreso BESS (acum. a {MESES_CORTOS[cierre_b - 1]})", f"USD {fmt(cerr_b['Ingreso_Total_USD'].sum(), 0)}")
k4.metric("Mes en curso", MESES[HOY.month - 1] if HOY.year == anio else "–", f"PPA {fmt(PRECIO_PPA, 2)} USD/MWh", delta_color="off")


# =========================================================
# GRAFICOS (plotly)
# =========================================================
def estilo(fig, alto, leyenda_abajo=False, leyenda=True):
    fig.update_layout(
        height=alto, paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF", font=dict(family=FUENTE_WEB, color=COLOR_TEXTO),
        title=dict(font=dict(size=16, color=COLOR_TEXTO), x=.01, xanchor="left"), hovermode="x unified", showlegend=leyenda,
        separators=",.", margin=dict(l=65, r=30, t=70, b=110 if leyenda_abajo else 55),
        legend=(dict(orientation="h", yanchor="top", y=-.2, xanchor="center", x=.5) if leyenda_abajo
                else dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0)))
    fig.update_xaxes(showgrid=False, showline=True, linecolor=COLOR_BORDE, tickfont=dict(color=COLOR_TEXTO_2))
    fig.update_yaxes(showgrid=True, gridcolor=COLOR_REJILLA, zeroline=True, zerolinecolor=COLOR_BORDE, tickfont=dict(color=COLOR_TEXTO_2))
    return fig


def mostrar(fig, clave):
    st.plotly_chart(fig, width="stretch", key=clave, config={"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d"]})


# =========================================================
# PESTAÑAS
# =========================================================
tab_car, tab_bess, tab_port, tab_rep = st.tabs(["Monitoreo diario", "BESS María Elena", "Portafolio", "Reportes y descargas"])

# ---------------- Monitoreo diario ----------------
with tab_car:
    st.caption(f"Rotación automática cada {SEGUNDOS_ROTACION_GRAFICOS} segundos · últimos {DIAS_CARRUSEL} días "
               f"({CARRUSEL_INI:%d-%m-%Y} al {CARRUSEL_FIN:%d-%m-%Y}) · Fuente: API")
    for llave, valor in {"car_i": 0, "car_pausa": False, "car_init": False, "car_manual": False}.items():
        st.session_state.setdefault(llave, valor)
    disponibles = [p for p in ORDEN_CARRUSEL if p in set(df_carrusel["Proyecto"])]

    def car_mover(paso):
        if disponibles:
            st.session_state.car_i = (st.session_state.car_i + paso) % len(disponibles)
            st.session_state.car_manual = True

    def car_pausa():
        st.session_state.car_pausa = not st.session_state.car_pausa
        st.session_state.car_manual = True

    def car_elegir():
        sel = st.session_state.car_sel
        if sel in disponibles:
            st.session_state.car_i = disponibles.index(sel)
        st.session_state.car_manual = True

    @st.fragment(run_every=SEGUNDOS_ROTACION_GRAFICOS)
    def carrusel():
        if time.time() - st.session_state.carga_ts > TTL_DATOS:   # pantallas fijas: refresco completo periódico
            st.session_state.carga_ts = time.time()
            st.rerun()
        if not disponibles:
            st.info("No hay datos diarios disponibles para mostrar.")
            return
        n = len(disponibles)
        st.session_state.car_i %= n
        if not st.session_state.car_init:
            st.session_state.car_init = True
        elif st.session_state.car_manual:
            st.session_state.car_manual = False
        elif not st.session_state.car_pausa:
            st.session_state.car_i = (st.session_state.car_i + 1) % n
        proyecto = disponibles[st.session_state.car_i]
        st.session_state.car_sel = proyecto
        b1, b2, b3, b4 = st.columns([1, 1, 1, 3])
        b1.button("◀ Anterior", width="stretch", key="car_prev", on_click=car_mover, args=(-1,))
        b2.button("▶ Reanudar" if st.session_state.car_pausa else "⏸ Pausar", width="stretch", key="car_pause", on_click=car_pausa)
        b3.button("Siguiente ▶", width="stretch", key="car_next", on_click=car_mover, args=(1,))
        b4.selectbox("Ir al proyecto", disponibles, key="car_sel", on_change=car_elegir, label_visibility="collapsed")
        d = df_carrusel[df_carrusel["Proyecto"] == proyecto][["Fecha_Dia", "Energia_MWh"]]
        d = pd.DataFrame({"Fecha_Dia": pd.date_range(CARRUSEL_INI, CARRUSEL_FIN, freq="D")}).merge(d, on="Fecha_Dia", how="left")
        fig = go.Figure(go.Bar(x=d["Fecha_Dia"], y=d["Energia_MWh"], marker_color=COLOR_PRIMARIO, width=.58 * 864e5,
                               hovertemplate="<b>%{x|%d-%m-%Y}</b><br>Energía: %{y:,.2f} MWh<extra></extra>"))
        fig.update_layout(title=f"Energía diaria · {proyecto}", xaxis_title="Fecha", yaxis_title="Energía [MWh]")
        fig.update_xaxes(dtick=864e5, tickformat="%d-%m", range=[CARRUSEL_INI - pd.Timedelta(hours=12), CARRUSEL_FIN + pd.Timedelta(hours=12)])
        fig.update_yaxes(rangemode="tozero")
        mostrar(estilo(fig, 500, leyenda=False), f"car_{slug(proyecto)}")
        estado = "Pausado" if st.session_state.car_pausa else f"Rotación automática cada {SEGUNDOS_ROTACION_GRAFICOS} s"
        st.caption(f"Vista {st.session_state.car_i + 1} de {n} · {estado}")

    carrusel()

# ---------------- BESS ----------------
with tab_bess:
    if bess_anio.empty:
        st.info("No hay información disponible del BESS.")
    else:
        meses_b = sorted(bess_anio["Mes"].unique().tolist(), reverse=True)
        mes_b = st.selectbox("Mes para indicadores", meses_b, format_func=lambda m: MESES[m - 1], key="mes_kpi_bess")
        f = bess_anio[bess_anio["Mes"] == mes_b].iloc[0]
        bm = bud_bess_anio[bud_bess_anio["Mes"] == mes_b]
        dif = f["Energia_Total_MWh"] - bm.iloc[0]["Budget_Generacion_MWh"] if not bm.empty else 0
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Inyección total del mes", f"{fmt(f['Energia_Total_MWh'])} MWh", f"{fmt_signo(dif)} MWh vs budget")
        c2.metric("Energía PPA del mes", f"{fmt(f['Energia_PPA_MWh'])} MWh")
        c3.metric("Energía Spot del mes", f"{fmt(f['Energia_Spot_MWh'])} MWh")
        c4.metric("Ingreso del mes", f"USD {fmt(f['Ingreso_Total_USD'], 0)}")
        st.caption(f"Fuente: {f['Fuente']} · Precio PPA: USD {fmt(PRECIO_PPA, 2)}/MWh")

    st.subheader("Distribución diaria de energía")
    if api_bess_diario.empty:
        st.info("No hay información diaria desde la API. Los meses del Excel contienen información mensual.")
    else:
        opciones = sorted({(int(a), int(m)) for a, m in zip(api_bess_diario["Anio"], api_bess_diario["Mes"])}, reverse=True)
        ad, md = st.selectbox("Mes para gráfico diario", opciones, format_func=lambda x: f"{MESES[x[1] - 1]} {x[0]}", key="mes_diario_bess")
        primero, ultimo = pd.Timestamp(ad, md, 1), min(pd.Timestamp(ad, md, calendar.monthrange(ad, md)[1]), HOY)
        dm = api_bess_diario[(api_bess_diario["Anio"] == ad) & (api_bess_diario["Mes"] == md)]
        dm = pd.DataFrame({"Fecha_Dia": pd.date_range(primero, ultimo, freq="D")}).merge(
            dm[["Fecha_Dia", "Energia_PPA_MWh", "Energia_Spot_MWh"]], on="Fecha_Dia", how="left").fillna(0)
        fig = go.Figure()
        fig.add_bar(x=dm["Fecha_Dia"], y=dm["Energia_PPA_MWh"], name="PPA real", marker_color=COLOR_PRIMARIO, width=.4 * 864e5)
        fig.add_bar(x=dm["Fecha_Dia"], y=dm["Energia_Spot_MWh"], name="Spot real", marker_color=COLOR_SECUNDARIO, width=.4 * 864e5)
        fig.update_layout(title=f"Inyección diaria BESS · {MESES[md - 1]} {ad}", barmode="stack", xaxis_title="Día del mes", yaxis_title="Energía [MWh]")
        fig.update_xaxes(dtick=864e5, tickformat="%d", range=[primero - pd.Timedelta(hours=12), ultimo + pd.Timedelta(hours=12)])
        fig.update_yaxes(rangemode="tozero")
        mostrar(estilo(fig, 520, leyenda_abajo=True), "g_bess_diario")

    cb = pd.DataFrame({"Mes": range(1, 13)}).merge(bess_anio, on="Mes", how="left").merge(
        bud_bess_anio[["Mes", "Budget_Generacion_MWh", "Budget_PPA_MWh"]], on="Mes", how="left")
    for col in ("Energia_PPA_MWh", "Energia_Spot_MWh"):
        cb[col] = cb[col].fillna(0)
    nombres = cb["Mes"].map(lambda m: MESES_CORTOS[m - 1])
    fig = go.Figure()
    fig.add_bar(x=nombres, y=cb["Energia_PPA_MWh"], name="PPA real", marker_color=COLOR_PRIMARIO, width=.58)
    fig.add_bar(x=nombres, y=cb["Energia_Spot_MWh"], name="Spot real", marker_color=COLOR_SECUNDARIO, width=.58)
    fig.add_scatter(x=nombres, y=cb["Budget_Generacion_MWh"], name="Budget generación", mode="lines+markers", line=dict(color=COLOR_BUDGET, width=3))
    fig.add_scatter(x=nombres, y=cb["Budget_PPA_MWh"], name="Budget PPA", mode="lines+markers", line=dict(color=COLOR_BUDGET_PPA, width=3))
    fig.update_layout(title="Consolidado anual BESS", barmode="stack", xaxis_title="Mes", yaxis_title="Energía [MWh]")
    fig.update_yaxes(rangemode="tozero")
    mostrar(estilo(fig, 600, leyenda_abajo=True), "g_bess_anual")
    if not bess_anio.empty:
        with st.expander("Ver acumulados del BESS"):
            a1, a2, a3, a4 = st.columns(4)
            a1.metric("Inyección acumulada", f"{fmt(bess_anio['Energia_Total_MWh'].sum())} MWh")
            a2.metric("PPA acumulado", f"{fmt(bess_anio['Energia_PPA_MWh'].sum())} MWh")
            a3.metric("Spot acumulado", f"{fmt(bess_anio['Energia_Spot_MWh'].sum())} MWh")
            a4.metric("Ingreso acumulado", f"USD {fmt(bess_anio['Ingreso_Total_USD'].sum(), 0)}")

# ---------------- Portafolio ----------------
with tab_port:
    total = pd.DataFrame({"Mes": range(1, 13)})
    total["Gen"] = total["Mes"].map(gen_mes).fillna(0)
    total["Bud"] = total["Mes"].map(bud_mes)
    if ultimo_mes_gen:
        fila = total[total["Mes"] == ultimo_mes_gen].iloc[0]
        ga = total.loc[total["Mes"] <= ultimo_mes_gen, "Gen"].sum()
        ba = total.loc[total["Mes"] <= ultimo_mes_gen, "Bud"].fillna(0).sum()
        t1, t2, t3, t4 = st.columns(4)
        t1.metric(f"Generación total {MESES[ultimo_mes_gen - 1]}", f"{fmt(fila['Gen'])} MWh")
        t2.metric(f"Budget total {MESES[ultimo_mes_gen - 1]}", f"{fmt(fila['Bud'])} MWh" if pd.notna(fila["Bud"]) else "Sin budget")
        t3.metric("Generación acumulada", f"{fmt(ga)} MWh")
        t4.metric("Diferencia acumulada", f"{fmt_signo(ga - ba)} MWh", f"{fmt_signo(ga - ba)} MWh vs budget")
    nombres = total["Mes"].map(lambda m: MESES_CORTOS[m - 1])
    fig = go.Figure()
    fig.add_bar(x=nombres, y=total["Gen"], name="Generación real total", marker_color=COLOR_PRIMARIO, offsetgroup="real", width=.38)
    fig.add_bar(x=nombres, y=total["Bud"], name="Budget total", marker_color=COLOR_BUDGET_PROYECTO, offsetgroup="budget", width=.38)
    fig.update_layout(title="Generación mensual total vs budget", barmode="group", xaxis_title="Mes", yaxis_title="Generación [MWh]")
    mostrar(estilo(fig, 520, leyenda_abajo=True), "g_total")

    st.subheader("Detalle mensual por proyecto")
    mes_actual = int(FECHA_FIN[5:7])
    cols = st.columns(2)
    for i, elemento in enumerate(ORDEN_GRAFICOS):
        pmgd = elemento == "CONSOLIDADO PMGDS"
        sel = (lambda d: d["Proyecto"].isin(PROYECTOS_PMGD)) if pmgd else (lambda d, e=elemento: d["Proyecto"] == e)
        g = proy_anio[sel(proy_anio)].groupby("Mes")["Generacion_MWh"].sum()
        b = bud_proy_anio[sel(bud_proy_anio)].groupby("Mes")["Budget_Generacion_MWh"].sum()
        meses = list(range(1, mes_actual + 1))
        fig = go.Figure()
        fig.add_bar(x=[MESES_CORTOS[m - 1] for m in meses], y=[g.get(m, 0) for m in meses], name="Generación", marker_color=COLOR_PRIMARIO, width=.52)
        fig.add_scatter(x=[MESES_CORTOS[m - 1] for m in meses], y=[b.get(m) for m in meses], name="Budget", mode="lines+markers",
                        line=dict(color=COLOR_BUDGET_PROYECTO, width=2), marker=dict(size=5))
        fig.update_layout(title="Consolidado PMGDs" if pmgd else elemento, yaxis_title="MWh", bargap=.38)
        with cols[i % 2]:
            mostrar(estilo(fig, 340, leyenda=(i == 0)), f"g_{slug(elemento)}")
    with st.expander("Fuentes utilizadas por proyecto"):
        fuentes = proy_anio[["Proyecto", "Mes", "Fuente"]].sort_values(["Proyecto", "Mes"]).copy()
        fuentes["Mes"] = fuentes["Mes"].map(lambda m: MESES[m - 1])
        st.dataframe(fuentes, width="stretch", hide_index=True)

# ---------------- Reportes y descargas ----------------
with tab_rep:
    st.subheader("Reporte semanal (lunes a viernes)")
    st.caption("Un PDF por proyecto (San Pedro III, Doña Carmen y María Elena) con generación inyectada, CMG ponderado e ingresos spot.")
    semanas = semanas_disponibles(anio)
    if not semanas:
        st.info("Aún no hay semanas disponibles.")
    else:
        r1, r2, r3 = st.columns([2, 1, 1], vertical_alignment="bottom")
        completas = [i for i, (_, _, vie) in enumerate(semanas) if vie <= HOY.date()]
        sel_sem = r1.selectbox("Semana del año", range(len(semanas)), index=completas[0] if completas else 0,
                               format_func=lambda i: f"Semana {semanas[i][0]} · {semanas[i][1]:%d-%m} al {semanas[i][2]:%d-%m}"
                               + (" (en curso)" if semanas[i][2] > HOY.date() else ""), key="sel_semana")
        tipo_sem = r2.selectbox("Tipo de CMG", list(TIPOS_CMG), index=1, key="tipo_cmg_semana")
        n_sem, lunes, viernes = semanas[sel_sem]
        clave_sel = (n_sem, anio, tipo_sem)
        if r3.button("Generar reportes", type="primary", width="stretch", key="b_generar_reportes"):
            with st.spinner(f"Generando reportes de la semana {n_sem}…"):
                pdfs, errores, adv_rep = generar_reportes_semana(n_sem, lunes, viernes, tipo_sem)
                zip_buf = io.BytesIO()
                if pdfs:
                    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
                        for nombre, datos in pdfs.items():
                            z.writestr(nombre, datos)
                st.session_state.reportes = {"clave": clave_sel, "pdfs": pdfs, "zip": zip_buf.getvalue(), "errores": errores, "adv": adv_rep}
        rep = st.session_state.get("reportes")
        if rep and rep["clave"] == clave_sel:
            for e in rep["errores"]:
                st.error(e)
            if rep["pdfs"]:
                st.success(f"Semana {n_sem}: {len(rep['pdfs'])} reporte(s) listo(s) (CMG {tipo_sem}).")
                bcols = st.columns(len(rep["pdfs"]) + 1)
                for col, (nombre, datos) in zip(bcols, rep["pdfs"].items()):
                    col.download_button(f"PDF {nombre.split('_', 2)[2].replace('_', ' ')[:-4]}", datos, file_name=nombre, mime="application/pdf",
                                        width="stretch", key=f"dl_{nombre}", on_click="ignore", icon=":material/picture_as_pdf:")
                bcols[-1].download_button("Todos (ZIP)", rep["zip"], file_name=f"Reportes_Semana{n_sem}.zip", mime="application/zip",
                                          width="stretch", key="dl_zip_reportes", on_click="ignore", icon=":material/folder_zip:")
            if rep["adv"]:
                with st.expander("Advertencias de la consulta"):
                    for a in rep["adv"]:
                        st.markdown(f"- {a}")
        elif rep:
            st.info("Cambió la selección: pulsa «Generar reportes» para actualizar los archivos.")

    st.divider()
    st.subheader("Consolidados Excel")
    st.caption("Usan los valores ya cargados en el dashboard; no vuelven a consultar la API.")
    periodos_rep = sorted({(int(r.Anio), int(r.Mes)) for r in df_proy.itertuples()} | {(int(r.Anio), int(r.Mes)) for r in df_bess.itertuples()}, reverse=True)
    d1, d2, d3 = st.columns([2, 1, 1], vertical_alignment="bottom")
    per_desc = d1.selectbox("Mes del consolidado mensual", periodos_rep, format_func=lambda x: f"{MESES[x[1] - 1]} {x[0]}", key="periodo_mensual")
    d2.download_button("Consolidado mensual", generar_excel_consolidado_mes(df_bess, df_proy, per_desc[0], per_desc[1]),
                       file_name=f"{per_desc[1]:02d}{per_desc[0]:04d}.xlsx", width="stretch", key="dl_mensual", on_click="ignore",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", icon=":material/download:")
    d3.download_button("Consolidado anual", generar_excel_consolidado_anual(df_bess, df_proy, anio),
                       file_name=f"Consolidado Anual {anio}.xlsx", width="stretch", key="dl_anual", on_click="ignore",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", icon=":material/download:")

    st.divider()
    st.subheader("PRMTE horario con CMG")
    st.caption("Inyección horaria de todos los parques. Doña Carmen, San Pedro III, María Elena y BESS incluyen el CMG horario del tipo elegido. "
               "Se consulta la API solo para el rango pedido.")
    p1, p2, p3, p4 = st.columns([1, 1, 1, 1], vertical_alignment="bottom")
    min_f, max_f = date(ANIO_OPERACIONAL, 1, 1), HOY.date()
    f_ini = p1.date_input("Fecha inicial", value=max(HOY.replace(day=1).date(), min_f), min_value=min_f, max_value=max_f, key="f_ini_prmte")
    f_fin = p2.date_input("Fecha final", value=max_f, min_value=min_f, max_value=max_f, key="f_fin_prmte")
    tipo_prmte = p3.selectbox("Tipo de CMG", list(TIPOS_CMG), key="tipo_cmg_prmte")
    if p4.button("Preparar archivo PRMTE", width="stretch", key="b_prmte"):
        if f_fin < f_ini:
            st.error("La fecha final debe ser igual o posterior a la inicial.")
        else:
            with st.spinner("Descargando y consolidando PRMTE horario…"):
                datos_x, adv_x = generar_excel_prmte(f_ini.isoformat(), f_fin.isoformat(), tipo_prmte)
            st.session_state.prmte_archivo = {"datos": datos_x, "adv": adv_x, "clave": (f_ini, f_fin, tipo_prmte)}
    arch = st.session_state.get("prmte_archivo")
    if arch:
        i, f_, t_ = arch["clave"]
        st.download_button(f"Descargar PRMTE {i:%d-%m-%Y} a {f_:%d-%m-%Y} (CMG {t_})", arch["datos"],
                           file_name=f"PRMTE_{i:%Y%m%d}_{f_:%Y%m%d}_CMG{t_}.xlsx", key="dl_prmte", on_click="ignore",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", icon=":material/download:")
        if arch["adv"]:
            with st.expander(f"Advertencias del archivo ({len(arch['adv'])})"):
                for a in arch["adv"]:
                    st.markdown(f"- {a}")

st.divider()
st.caption(f"Dashboard de operación · Atlantica Sustainable Infrastructure · Última carga: {ahora():%d-%m-%Y %H:%M} (hora de Chile)")
