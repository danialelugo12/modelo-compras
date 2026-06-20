"""
================================================================================
MODELO DE COMPRAS - PROTOTIPO WEB APP
================================================================================
Dashboard interactivo del modelo de compras. Diseñado como prototipo para
mostrar cómo se vería el módulo dentro de la app empresarial.

Para ejecutar:
    streamlit run app.py

Se abre automáticamente en el navegador en http://localhost:8501
================================================================================
"""

import io
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

# Reutilizar lógica ya construida
from config import CONFIG
from modelo_compras import (
    procesar_ventas,
    procesar_fallidas,
    procesar_compras,
    construir_modelo,
    validar_columnas,
)
from exportar_excel import generar_excel
import logging


# ============================================================================
# CONFIGURACIÓN PÁGINA
# ============================================================================
st.set_page_config(
    page_title="Modelo de Compras",
    page_icon="🚗",
    layout="wide",
    initial_sidebar_state="expanded",
)

# === EVITAR QUE EL NAVEGADOR AUTOTRADUZCA (la "C" se volvia "do") ===
import streamlit.components.v1 as components
components.html(
    """
    <script>
      const doc = window.parent.document;
      doc.documentElement.setAttribute('translate', 'no');
      doc.documentElement.classList.add('notranslate');
      const meta = doc.createElement('meta');
      meta.name = 'google';
      meta.content = 'notranslate';
      doc.head.appendChild(meta);
    </script>
    """,
    height=0,
)

# CSS personalizado para look profesional
st.markdown("""
<style>
    /* === OCULTAR BOTONES DE GITHUB Y FORK === */
    .stActionButton, 
    [data-testid="stToolbar"],
    [data-testid="baseButton-header"],
    a[href*="github.com"],
    button[kind="header"] {
        display: none !important;
        visibility: hidden !important;
    }
    header [data-testid="stToolbar"] {
        display: none !important;
    }
    /* === FORZAR FONDO BLANCO === */
    .stApp {
        background-color: #FFFFFF !important;
    }
    [data-testid="stMain"] {
        background-color: #FFFFFF !important;
    }
    [data-testid="stHeader"] {
        background-color: #FFFFFF !important;
    }
    .main .block-container {
        background-color: #FFFFFF !important;
        padding-top: 2rem;
    }
    /* Texto general en oscuro para fondo blanco */
    .stApp, .stApp p, .stApp label, .stApp span, .stApp div {
        color: #262730;
    }
    /* Sidebar mantiene su color */
    [data-testid="stSidebar"] {
        background-color: #F5F5F5 !important;
    }
    [data-testid="stSidebar"] * {
        color: #262730 !important;
    }
    /* Excepción: SOLUPARTS dentro del bloque rojo va en blanco */
    .soluparts-logo,
    .soluparts-logo *,
    .soluparts-text,
    [data-testid="stSidebar"] .soluparts-logo,
    [data-testid="stSidebar"] .soluparts-logo *,
    [data-testid="stSidebar"] .soluparts-text,
    [data-testid="stSidebar"] div[style*="background-color: #8B0000"] span,
    [data-testid="stSidebar"] div[style*="background-color: #8B0000"] * {
        color: #FFFFFF !important;
    }
    /* === TÍTULO PRINCIPAL === */
    .main-header {
        font-size: 3rem !important;
        font-weight: bold !important;
        color: #8B0000 !important;
        padding: 1rem 0;
        border-bottom: 3px solid #8B0000;
        margin-bottom: 2rem;
    }
    .stApp .main-header,
    .stApp div.main-header {
        color: #8B0000 !important;
    }
    /* === BOTONES Y CONTROLES DE SIDEBAR === */
    /* File uploaders - dropzone con rojo claro */
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
        background-color: #FFE6E6 !important;
        border: 2px dashed #C77C7C !important;
    }
    [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * {
        color: #8B0000 !important;
    }
    /* Botones generales del sidebar */
    [data-testid="stSidebar"] .stButton button {
        background-color: #FFE6E6 !important;
        color: #8B0000 !important;
        border: 1px solid #C77C7C !important;
    }
    [data-testid="stSidebar"] .stButton button:hover {
        background-color: #FFCCCC !important;
        border-color: #8B0000 !important;
    }
    /* === TARJETAS DE ARCHIVOS YA CARGADOS (uploaded file chips) === */
    [data-testid="stSidebar"] [data-testid="stFileUploaderFile"],
    [data-testid="stSidebar"] [data-testid="stFileUploaderFileName"],
    [data-testid="stSidebar"] li[class*="uploadedFile"],
    [data-testid="stSidebar"] div[class*="uploadedFile"],
    [data-testid="stSidebar"] section[class*="uploadedFile"] {
        background-color: #FFE6E6 !important;
        border: 1px solid #C77C7C !important;
        border-radius: 6px !important;
    }
    [data-testid="stSidebar"] [data-testid="stFileUploaderFile"] *,
    [data-testid="stSidebar"] [data-testid="stFileUploaderFileName"] *,
    [data-testid="stSidebar"] li[class*="uploadedFile"] *,
    [data-testid="stSidebar"] div[class*="uploadedFile"] *,
    [data-testid="stSidebar"] section[class*="uploadedFile"] * {
        color: #8B0000 !important;
        background-color: transparent !important;
    }
    /* === BOTÓN DESCARGAR EXCEL (download button) === */
    [data-testid="stDownloadButton"] button {
        background-color: #FFE6E6 !important;
        color: #8B0000 !important;
        border: 1px solid #C77C7C !important;
        font-weight: 600 !important;
    }
    [data-testid="stDownloadButton"] button:hover {
        background-color: #FFCCCC !important;
        border-color: #8B0000 !important;
        color: #8B0000 !important;
    }
    [data-testid="stDownloadButton"] button * {
        color: #8B0000 !important;
    }
    /* === TABLAS Y DATAFRAMES - fondo ROJO CLARO === */
    /* Iframe contenedor (Streamlit usa esto a veces) */
    [data-testid="stDataFrame"] iframe {
        background-color: #FFF5F5 !important;
    }
    /* Contenedor del dataframe */
    [data-testid="stDataFrame"] {
        background-color: #FFF5F5 !important;
        border: 1px solid #F8CBCB !important;
        border-radius: 8px;
        padding: 4px;
    }
    [data-testid="stDataFrame"] > div {
        background-color: #FFF5F5 !important;
    }
    [data-testid="stDataFrame"] * {
        background-color: transparent !important;
        color: #262730 !important;
    }
    /* Streamlit usa glide_data_grid internamente */
    [data-testid="stDataFrame"] canvas {
        background-color: #FFF5F5 !important;
    }
    /* st.table (HTML tradicional) */
    .stApp table {
        background-color: #FFF5F5 !important;
        color: #262730 !important;
        border: 1px solid #F8CBCB !important;
    }
    .stApp table thead th {
        background-color: #FFE6E6 !important;
        color: #8B0000 !important;
        font-weight: 600;
    }
    .stApp table tbody td {
        background-color: #FFF5F5 !important;
        color: #262730 !important;
    }
    .stApp table tbody tr:nth-child(even) td {
        background-color: #FFFAFA !important;
    }
    .stApp table td {
        background-color: #FFFFFF !important;
        color: #262730 !important;
    }
    /* === TARJETAS === */
    .metric-card {
        background: #f0f2f6;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #8B0000;
    }
    .alert-critical {
        background: #FFE6E6;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #C00000;
    }
    .alert-warning {
        background: #FFF4E6;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #ED7D31;
    }
    .alert-success {
        background: #E6F4EA;
        padding: 1rem;
        border-radius: 8px;
        border-left: 4px solid #2E7D32;
    }
    div[data-testid="stMetricValue"] {
        font-size: 1.8rem;
        color: #8B0000 !important;
    }
    div[data-testid="stMetricLabel"] {
        color: #262730 !important;
    }
    /* === DASHBOARD KPI CARDS === */
    .kpi-card {
        background: #FAFAFA;
        padding: 1.2rem;
        border-radius: 10px;
        border: 1px solid #E0E0E0;
        border-left: 4px solid #8B0000;
        text-align: center;
    }
    .kpi-card-title {
        font-size: 0.85rem;
        color: #666666;
        margin: 0;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .kpi-card-value {
        font-size: 1.8rem;
        font-weight: bold;
        color: #8B0000;
        margin: 0.5rem 0;
    }
    .kpi-card-subtitle {
        font-size: 0.8rem;
        color: #888888;
        margin: 0;
    }
    /* === TABS === */
    .stTabs [data-baseweb="tab-list"] button[aria-selected="true"] {
        background-color: #8B0000;
        color: white;
    }
</style>
""", unsafe_allow_html=True)


# ============================================================================
# LOGGER FAKE (para reusar funciones existentes)
# ============================================================================
class StreamlitLogger:
    """Logger que no escribe a archivo, solo silencia los .info()."""
    def info(self, msg): pass
    def warning(self, msg): pass
    def error(self, msg): pass
    def exception(self, msg): pass


# ============================================================================
# HELPER DE ESTILO PARA TABLAS
# ============================================================================
def estilizar_df(df, formato_columnas=None):
    """Aplica estilo rojo claro al dataframe usando pandas Styler.
    Esto se ve mejor que el estilo por defecto oscuro de Streamlit."""
    styler = df.style

    # Estilo general: fondo rojo muy claro, texto oscuro, líneas separadoras
    styler = styler.set_properties(**{
        'background-color': '#FFF5F5',
        'color': '#262730',
        'border': '1px solid #F8CBCB',
        'padding': '6px 8px',
    })

    # Estilo de encabezados
    styler = styler.set_table_styles([
        {'selector': 'thead th', 'props': [
            ('background-color', '#FFE6E6'),
            ('color', '#8B0000'),
            ('font-weight', 'bold'),
            ('border', '1px solid #F8CBCB'),
            ('padding', '8px'),
            ('text-align', 'left'),
        ]},
        {'selector': 'tbody tr:nth-child(even) td', 'props': [
            ('background-color', '#FFFAFA'),
        ]},
        {'selector': 'tbody tr:hover td', 'props': [
            ('background-color', '#FFE6E6'),
        ]},
    ])

    # Aplicar formatos numéricos si se especifican
    if formato_columnas:
        styler = styler.format(formato_columnas)

    return styler


# ============================================================================
# CACHE DE PROCESAMIENTO
# ============================================================================
@st.cache_data(show_spinner=False)
def procesar_todo(maestro_bytes, ventas_bytes, fallidas_bytes, compras_bytes):
    """Procesa los 4 archivos y devuelve el modelo + ranking SIN ID."""
    logger = StreamlitLogger()

    # Cargar
    xls_compras = pd.ExcelFile(io.BytesIO(compras_bytes))
    hoja_maestro = next((h for h in xls_compras.sheet_names if "MAESTRO" in h.upper()), None)

    if hoja_maestro:
        maestro = pd.read_excel(io.BytesIO(compras_bytes), sheet_name=hoja_maestro)
    else:
        maestro = pd.read_excel(io.BytesIO(maestro_bytes))

    ventas_df = pd.read_excel(io.BytesIO(ventas_bytes))
    fallidas_df = pd.read_excel(io.BytesIO(fallidas_bytes))
    compras_df = pd.read_excel(io.BytesIO(compras_bytes), sheet_name="COMPRA NACIONAL")

    # Procesar
    ventas = procesar_ventas(ventas_df, logger)
    fallidas, ranking_sin_id = procesar_fallidas(fallidas_df, logger)
    costos, transito = procesar_compras(compras_df, logger)
    modelo = construir_modelo(maestro, ventas, fallidas, costos, transito, logger)

    return modelo, ranking_sin_id


@st.cache_data(show_spinner=False)
def generar_excel_bytes(_modelo, _ranking):
    """Genera el Excel en memoria y devuelve los bytes."""
    logger = StreamlitLogger()
    tmp_path = Path("OUTPUT") / f"temp_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    tmp_path.parent.mkdir(exist_ok=True)
    generar_excel(_modelo, _ranking, tmp_path, logger)
    with open(tmp_path, "rb") as f:
        data = f.read()
    tmp_path.unlink()
    return data


# ============================================================================
# SIDEBAR - CARGA DE ARCHIVOS
# ============================================================================
with st.sidebar:
    # Logo SOLUPARTS como SVG inline (a prueba de cualquier CSS interferente)
    import base64
    svg_logo = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 300 110" width="100%" height="95">
        <rect x="0" y="0" width="300" height="110" rx="8" ry="8" fill="#8B0000"/>
        <text x="150" y="52" font-family="Segoe UI, Arial, sans-serif" font-size="32" font-weight="bold" font-style="italic" fill="#FFFFFF" text-anchor="middle" letter-spacing="3">SOLUPARTS</text>
        <text x="150" y="82" font-family="Segoe UI, Arial, sans-serif" font-size="15" font-weight="bold" font-style="italic" fill="#FFFFFF" text-anchor="middle" text-decoration="underline" letter-spacing="1">ADQUISICIONES – DANIELA LUGO</text>
    </svg>'''
    svg_b64 = base64.b64encode(svg_logo.encode()).decode()
    st.markdown(
        f'<div style="margin-bottom: 1rem; text-align: center;"><img src="data:image/svg+xml;base64,{svg_b64}" style="max-width: 100%; height: auto;"/></div>',
        unsafe_allow_html=True
    )
    st.markdown("## 📂 Datos de entrada")

    # --- Modo automático: leer desde carpeta INPUT del repositorio ---
    _INPUT = Path(__file__).parent / "INPUT"
    _archivos_auto = {
        "maestro":  next(_INPUT.glob("MAESTRO*.xlsx"), None)   if _INPUT.exists() else None,
        "ventas":   next(_INPUT.glob("VENTAS*.xlsx"), None)    if _INPUT.exists() else None,
        "fallidas": next(_INPUT.glob("Planilla*.xlsx"), None)  if _INPUT.exists() else None,
        "compras":  next(_INPUT.glob("COMPRAS*.xlsx"), None)   if _INPUT.exists() else None,
    }
    _modo_auto = all(_archivos_auto.values())

    if _modo_auto:
        st.success("✅ Archivos del mes cargados automáticamente desde el repositorio.")
        st.caption(f"📁 Maestro: {_archivos_auto['maestro'].name}")
        st.caption(f"📁 Ventas: {_archivos_auto['ventas'].name}")
        st.caption(f"📁 Fallidas: {_archivos_auto['fallidas'].name}")
        st.caption(f"📁 Compras: {_archivos_auto['compras'].name}")
        maestro_file  = type("F", (), {"getvalue": lambda s: open(_archivos_auto["maestro"],  "rb").read()})()
        ventas_file   = type("F", (), {"getvalue": lambda s: open(_archivos_auto["ventas"],   "rb").read()})()
        fallidas_file = type("F", (), {"getvalue": lambda s: open(_archivos_auto["fallidas"], "rb").read()})()
        compras_file  = type("F", (), {"getvalue": lambda s: open(_archivos_auto["compras"],  "rb").read()})()
    else:
        st.caption("Sube los 4 archivos del mes para generar el análisis")
        maestro_file  = st.file_uploader("1. Maestro de productos",  type=["xlsx"], key="maestro")
        ventas_file   = st.file_uploader("2. Ventas 6 meses",        type=["xlsx"], key="ventas")
        fallidas_file = st.file_uploader("3. Ventas fallidas",       type=["xlsx"], key="fallidas")
        compras_file  = st.file_uploader("4. Compras realizadas",    type=["xlsx"], key="compras")

    st.markdown("---")
    st.markdown("### ⚙️ Parámetros del modelo")

    cob_a = st.slider("Cobertura objetivo clase A (meses)", 1.0, 5.0, 2.5, 0.5)
    cob_b = st.slider("Cobertura objetivo clase B (meses)", 1.0, 4.0, 2.0, 0.5)
    cob_c = st.slider("Cobertura objetivo clase C (meses)", 0.5, 3.0, 1.0, 0.5)

    st.markdown("---")
    st.markdown("### 💰 Presupuesto del mes")
    st.caption("Define cuánto tienes disponible para invertir y verás hasta dónde alcanza")
    presupuesto_mes = st.number_input(
        "Monto disponible (CLP)",
        min_value=0,
        max_value=10_000_000_000,
        value=0,
        step=10_000_000,
        format="%d",
        help="Deja en 0 para no usar restricción de presupuesto. Si pones un monto, el sistema te mostrará priorización por uso del presupuesto."
    )

    st.markdown("---")
    st.markdown("### 🌐 Contexto de mercado")
    st.caption("Indicadores económicos en tiempo real para ajustar las sugerencias de compra.")

    import requests as _req
    from datetime import date

    @st.cache_data(ttl=3600, show_spinner=False)
    def _obtener_dolar():
        try:
            r = _req.get("https://mindicador.cl/api/dolar/2026", timeout=6)
            data = r.json()
            valores = [v["valor"] for v in data.get("serie", [])]
            dolar_hoy = valores[0] if valores else None
            promedio_6m = sum(valores[:6]) / len(valores[:6]) if len(valores) >= 6 else None
            return dolar_hoy, promedio_6m
        except Exception:
            return None, None

    @st.cache_data(ttl=86400, show_spinner=False)
    def _calcular_señales_ventas():
        try:
            _HIST = Path(__file__).parent
            archivos_hist = (
                list(_HIST.glob("ENE-MAY*2025*.xlsx")) +
                list(_HIST.glob("JUN-DIC*2025*.xlsx")) +
                list(_HIST.glob("2026*.xlsx"))
            )
            if not archivos_hist:
                return None
            dfs = []
            for f in archivos_hist:
                df = pd.read_excel(f, usecols=["Fecha", "Tipo", "Precio Final"])
                dfs.append(df)
            hist = pd.concat(dfs, ignore_index=True)
            hist = hist[hist["Tipo"] == "BOLETA"].copy()
            hist["Fecha"] = pd.to_datetime(hist["Fecha"], dayfirst=True)
            hist["Año"] = hist["Fecha"].dt.year
            hist["Mes"] = hist["Fecha"].dt.month

            hoy = date.today()
            mes_actual = hoy.month
            año_actual = hoy.year

            def ventas_mes(año, mes):
                return hist[(hist["Año"] == año) & (hist["Mes"] == mes)]["Precio Final"].sum()

            def ventas_mes_hasta_dia(año, mes, dia):
                mask = (hist["Año"] == año) & (hist["Mes"] == mes) & (hist["Fecha"].dt.day <= dia)
                return hist[mask]["Precio Final"].sum()

            # --- Señal 1: Mes actual vs mismo período año anterior (hasta el mismo día) ---
            dia_hoy = hoy.day
            v_mes_actual = ventas_mes_hasta_dia(año_actual, mes_actual, dia_hoy)
            v_mes_ant = ventas_mes_hasta_dia(año_actual - 1, mes_actual, dia_hoy)
            var_anual = ((v_mes_actual - v_mes_ant) / v_mes_ant * 100) if v_mes_ant > 0 else None

            # --- Señal 2: Promedio últimos 3 meses vs mismo período año anterior ---
            meses_3m = []
            for i in range(1, 4):
                m = mes_actual - i
                a = año_actual
                if m <= 0:
                    m += 12
                    a -= 1
                meses_3m.append((a, m))

            v_3m_actual = [ventas_mes(a, m) for a, m in meses_3m]
            v_3m_ant = [ventas_mes(a - 1, m) for a, m in meses_3m]
            v_3m_actual = [v for v in v_3m_actual if v > 0]
            v_3m_ant = [v for v in v_3m_ant if v > 0]
            var_3m_anual = None
            if v_3m_actual and v_3m_ant and len(v_3m_actual) == len(v_3m_ant):
                var_3m_anual = (sum(v_3m_actual) / len(v_3m_actual) - sum(v_3m_ant) / len(v_3m_ant)) / (sum(v_3m_ant) / len(v_3m_ant)) * 100

            # --- Señal 3: Tendencia interna — últimos 3 meses vs 3 meses anteriores (mismo año) ---
            meses_prev = []
            for i in range(4, 7):
                m = mes_actual - i
                a = año_actual
                if m <= 0:
                    m += 12
                    a -= 1
                meses_prev.append((a, m))

            v_prev = [ventas_mes(a, m) for a, m in meses_prev]
            v_prev = [v for v in v_prev if v > 0]
            var_interna = None
            if v_3m_actual and v_prev:
                prom_reciente = sum(v_3m_actual) / len(v_3m_actual)
                prom_previo = sum(v_prev) / len(v_prev)
                var_interna = (prom_reciente - prom_previo) / prom_previo * 100

            return {
                "mes_actual": mes_actual,
                "año_actual": año_actual,
                "v_mes_actual": v_mes_actual,
                "v_mes_ant": v_mes_ant,
                "var_anual": var_anual,
                "prom_3m_actual": sum(v_3m_actual) / len(v_3m_actual) if v_3m_actual else None,
                "prom_3m_ant": sum(v_3m_ant) / len(v_3m_ant) if v_3m_ant else None,
                "var_3m_anual": var_3m_anual,
                "var_interna": var_interna,
            }
        except Exception:
            return None

    dolar_hoy, dolar_prom = _obtener_dolar()
    señales = _calcular_señales_ventas()

    # --- Calcular ajuste combinado con pesos ---
    ajuste_sugerido = 0
    lineas_resumen = []

    # Señal dólar (peso: 20%)
    if dolar_hoy and dolar_prom:
        var_dolar = (dolar_hoy - dolar_prom) / dolar_prom * 100
        lineas_resumen.append(f"<b>💵 Dólar:</b> ${dolar_hoy:,.0f} | Prom.6m: ${dolar_prom:,.0f} | {var_dolar:+.1f}%")
        if var_dolar >= 10:
            ajuste_sugerido -= 6
        elif var_dolar >= 5:
            ajuste_sugerido -= 3
        elif var_dolar <= -5:
            ajuste_sugerido += 3

    if señales:
        # Señal año a año mes actual (peso: 30%)
        if señales["var_anual"] is not None:
            lineas_resumen.append(
                f"<b>📅 Mes actual vs año ant.:</b> ${señales['v_mes_actual']/1e6:.1f}M vs ${señales['v_mes_ant']/1e6:.1f}M | {señales['var_anual']:+.1f}%"
            )
            v = señales["var_anual"]
            if v <= -15:
                ajuste_sugerido -= 9
            elif v <= -5:
                ajuste_sugerido -= 6
            elif v <= 0:
                ajuste_sugerido -= 3
            elif v >= 15:
                ajuste_sugerido += 6
            elif v >= 5:
                ajuste_sugerido += 3

        # Señal 3m vs año anterior (peso: 20%)
        if señales["var_3m_anual"] is not None:
            lineas_resumen.append(
                f"<b>📊 Prom.3m vs año ant.:</b> ${señales['prom_3m_actual']/1e6:.1f}M vs ${señales['prom_3m_ant']/1e6:.1f}M | {señales['var_3m_anual']:+.1f}%"
            )
            v = señales["var_3m_anual"]
            if v <= -15:
                ajuste_sugerido -= 6
            elif v <= -5:
                ajuste_sugerido -= 3
            elif v >= 15:
                ajuste_sugerido += 3

        # Señal tendencia interna (peso dominante — la mas importante)
        if señales["var_interna"] is not None:
            lineas_resumen.append(
                f"<b>📉 Tendencia interna:</b> últimos 3m vs 3m anteriores | {señales['var_interna']:+.1f}%"
            )
            v = señales["var_interna"]
            if v <= -15:
                ajuste_sugerido -= 20
            elif v <= -10:
                ajuste_sugerido -= 15
            elif v <= -5:
                ajuste_sugerido -= 10
            elif v < 0:
                ajuste_sugerido -= 5
            elif v >= 15:
                ajuste_sugerido += 8
            elif v >= 5:
                ajuste_sugerido += 4

            # Umbral minimo: si tendencia interna negativa, condicion NUNCA puede ser Normal o Favorable
            if v < 0:
                ajuste_sugerido = min(ajuste_sugerido, -5)

    # Redondear al múltiplo de 5 más cercano
    ajuste_sugerido = round(ajuste_sugerido / 5) * 5
    ajuste_sugerido = max(-30, min(20, ajuste_sugerido))

    # Condición combinada
    if ajuste_sugerido <= -20:
        condicion = "🔴 Muy adversa"
        color_cond = "#8B0000"
    elif ajuste_sugerido <= -10:
        condicion = "🟠 Adversa"
        color_cond = "#CC4400"
    elif ajuste_sugerido <= -5:
        condicion = "🟡 Moderada"
        color_cond = "#B8860B"
    elif ajuste_sugerido >= 10:
        condicion = "🟢 Favorable"
        color_cond = "#006400"
    elif ajuste_sugerido >= 5:
        condicion = "🟢 Levemente favorable"
        color_cond = "#228B22"
    else:
        condicion = "⚪ Normal"
        color_cond = "#444444"

    if lineas_resumen:
        cuerpo = "<br>".join(lineas_resumen)
        st.markdown(
            f'<div style="background:#F5F5F5; border-radius:8px; padding:10px 14px; margin-bottom:8px; font-size:0.85rem;">'
            f'{cuerpo}<br>'
            f'<b>Condición:</b> <span style="color:{color_cond}; font-weight:bold;">{condicion}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

    ajuste_mercado = st.slider(
        "Ajuste por contexto de mercado (%)",
        min_value=-30, max_value=20,
        value=ajuste_sugerido, step=5,
        help="Ajuste automático basado en 3 señales: dólar, comparación año a año y tendencia interna. Puedes modificarlo."
    )
    if ajuste_mercado != 0:
        st.caption(f"{'⬇️ Reduciendo' if ajuste_mercado < 0 else '⬆️ Aumentando'} sugerencias en {abs(ajuste_mercado)}%")

    st.markdown("---")
    st.caption("📌 Prototipo v1.0")
    st.caption("En la app definitiva, los datos vendrán directo de Ailoo")


# ============================================================================
# MAIN
# ============================================================================
st.markdown('<div class="main-header">🚗 MODELO DE COMPRA INTELIGENTE</div>',
            unsafe_allow_html=True)

# Verificar que todos los archivos estén cargados
todos_cargados = all([maestro_file, ventas_file, fallidas_file, compras_file])

if not todos_cargados:
    st.info("👈 Sube los 4 archivos en el panel lateral para comenzar el análisis")

    # Mostrar preview de cómo se ve el dashboard
    st.markdown("### Vista previa del dashboard")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown('<div class="metric-card"><b>Productos con demanda</b><br><span style="font-size:1.8rem">—</span></div>',
                    unsafe_allow_html=True)
    with col2:
        st.markdown('<div class="alert-critical"><b>Productos CRÍTICOS</b><br><span style="font-size:1.8rem">—</span></div>',
                    unsafe_allow_html=True)
    with col3:
        st.markdown('<div class="alert-warning"><b>Inversión sugerida</b><br><span style="font-size:1.8rem">—</span></div>',
                    unsafe_allow_html=True)
    with col4:
        st.markdown('<div class="alert-success"><b>Capital en sobrestock</b><br><span style="font-size:1.8rem">—</span></div>',
                    unsafe_allow_html=True)

    st.stop()


# Aplicar override de cobertura si el usuario tocó los sliders
CONFIG["cobertura_obj"]["A"] = cob_a
CONFIG["cobertura_obj"]["B"] = cob_b
CONFIG["cobertura_obj"]["C"] = cob_c

# Procesar
with st.spinner("Procesando datos... (toma ~30 segundos la primera vez)"):
    try:
        modelo, ranking_sin_id = procesar_todo(
            maestro_file.getvalue(),
            ventas_file.getvalue(),
            fallidas_file.getvalue(),
            compras_file.getvalue(),
        )
    except Exception as e:
        st.error(f"❌ Error procesando archivos: {e}")
        st.exception(e)
        st.stop()



# ============================================================================
# APLICAR AJUSTE DE MERCADO
# ============================================================================
if ajuste_mercado != 0:
    factor = 1 + ajuste_mercado / 100
    modelo["COMPRA_AJUSTADA"] = (modelo["COMPRA_AJUSTADA"] * factor).clip(lower=0).round(0)
    if "COSTO_COMPRA_AJUSTADA" in modelo.columns:
        modelo["COSTO_COMPRA_AJUSTADA"] = modelo["COSTO_COMPRA_AJUSTADA"] * factor

# ============================================================================
# KPIs PRINCIPALES
# ============================================================================
total_productos = len(modelo)
criticos = (modelo["ESTADO_INVENTARIO"] == "CRÍTICO").sum()
sobrestock = (modelo["ESTADO_INVENTARIO"] == "SOBRESTOCK").sum()
a_comprar = modelo[(modelo["COMPRA_AJUSTADA"] > 0) & modelo["ULTIMO_PRECIO_NETO"].notna()]
inversion = a_comprar["COSTO_COMPRA_AJUSTADA"].sum()

modelo_calc = modelo.copy()
modelo_calc["VALOR_EXCESO"] = np.where(
    modelo_calc["ESTADO_INVENTARIO"] == "SOBRESTOCK",
    (modelo_calc["STOCK_ACTUAL"] - modelo_calc["DEMANDA_MENSUAL"] * modelo_calc["COBERTURA_OBJETIVO"])
    * modelo_calc["ULTIMO_PRECIO_NETO"],
    0,
)
valor_sobrestock = modelo_calc["VALOR_EXCESO"].sum()

col1, col2, col3, col4 = st.columns(4)
col1.metric("Productos con demanda", f"{total_productos:,}".replace(",", "."))
col2.metric("Productos CRÍTICOS", f"{criticos:,}".replace(",", "."),
            delta=f"-{criticos} requieren acción", delta_color="inverse")
col3.metric("Inversión sugerida", f"${inversion:,.0f} CLP".replace(",", "."),
            delta=f"{len(a_comprar):,} productos".replace(",", "."))
col4.metric("Capital en sobrestock", f"${valor_sobrestock:,.0f} CLP".replace(",", "."),
            delta=f"{sobrestock:,} productos".replace(",", "."), delta_color="inverse")


# ============================================================================
# ANÁLISIS DE PRESUPUESTO
# ============================================================================
if presupuesto_mes > 0:
    st.markdown("---")
    st.markdown("### 💰 Análisis de presupuesto disponible")

    # Ordenar productos a comprar por prioridad (ABC + CRÍTICO primero) y monto
    a_comprar_con_costo = a_comprar[a_comprar["ULTIMO_PRECIO_NETO"].notna()].copy()
    # Score de prioridad: A=1, B=2, C=3 | CRÍTICO=1, COMPRAR=2
    abc_score = {"A": 1, "B": 2, "C": 3}
    estado_score = {"CRÍTICO": 1, "COMPRAR": 2}
    a_comprar_con_costo["SCORE"] = (
        a_comprar_con_costo["CLASIFICACION_ABC"].map(abc_score).fillna(4) * 10
        + a_comprar_con_costo["ESTADO_INVENTARIO"].map(estado_score).fillna(3)
    )
    a_comprar_con_costo = a_comprar_con_costo.sort_values(
        ["SCORE", "COSTO_COMPRA_AJUSTADA"], ascending=[True, False]
    )

    # Acumular y cortar en el presupuesto
    a_comprar_con_costo["ACUMULADO"] = a_comprar_con_costo["COSTO_COMPRA_AJUSTADA"].cumsum()
    cubiertos = a_comprar_con_costo[a_comprar_con_costo["ACUMULADO"] <= presupuesto_mes]
    n_cubiertos = len(cubiertos)
    monto_usado = cubiertos["COSTO_COMPRA_AJUSTADA"].sum() if n_cubiertos > 0 else 0
    saldo = presupuesto_mes - monto_usado
    n_no_cubiertos = len(a_comprar_con_costo) - n_cubiertos
    monto_no_cubierto = a_comprar_con_costo["COSTO_COMPRA_AJUSTADA"].sum() - monto_usado

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Presupuesto definido",
              f"${presupuesto_mes:,.0f} CLP".replace(",", "."))
    c2.metric("Productos cubiertos",
              f"{n_cubiertos:,}".replace(",", "."),
              delta=f"${monto_usado:,.0f} CLP usados".replace(",", "."))
    c3.metric("Saldo disponible",
              f"${saldo:,.0f} CLP".replace(",", "."),
              delta="Sobrante después de prioridad")
    c4.metric("Quedan por cubrir",
              f"{n_no_cubiertos:,}".replace(",", "."),
              delta=f"${monto_no_cubierto:,.0f} CLP faltan".replace(",", "."),
              delta_color="inverse")

    # Mensaje contextual
    if presupuesto_mes >= inversion:
        st.success(f"✅ Tu presupuesto cubre el 100% de la inversión sugerida. Te sobran ${saldo:,.0f} CLP".replace(",", "."))
    elif presupuesto_mes >= inversion * 0.7:
        st.info(f"💡 Tu presupuesto cubre {presupuesto_mes/inversion*100:.0f}% de la inversión sugerida. Estás priorizando los productos más críticos.")
    else:
        st.warning(f"⚠️ Tu presupuesto cubre solo {presupuesto_mes/inversion*100:.0f}% de la inversión sugerida. Considera ampliarlo o aceptar quiebres en productos clase B/C.")

    with st.expander("📋 Ver productos cubiertos con el presupuesto actual"):
        if n_cubiertos > 0:
            preview = cubiertos[["Id", "Marca", "Producto", "CLASIFICACION_ABC",
                                  "ESTADO_INVENTARIO", "COMPRA_AJUSTADA",
                                  "ULTIMO_PRECIO_NETO", "COSTO_COMPRA_AJUSTADA"]].copy()
            preview.columns = ["ID", "Marca", "Producto", "ABC", "Estado",
                               "Comprar", "Costo unit (CLP)", "Inversión (CLP)"]
            st.dataframe(preview, use_container_width=True, hide_index=True)
        else:
            st.warning("Tu presupuesto no alcanza ni para el producto más prioritario. Considera ampliarlo.")


# ============================================================================
# DESCARGA EXCEL
# ============================================================================
st.markdown("---")
col1, col2 = st.columns([3, 1])
with col1:
    st.markdown("### 📊 Reporte completo")
    st.caption("Excel con 6 hojas: Resumen, Modelo completo, Prioridad compra, Sobrestock, Acción costos, Oportunidades SIN ID")
with col2:
    excel_bytes = generar_excel_bytes(modelo, ranking_sin_id)
    st.download_button(
        label="⬇️ Descargar Excel",
        data=excel_bytes,
        file_name=f"MODELO_COMPRAS_{datetime.now():%Y-%m-%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


# ============================================================================
# TABS DE ANÁLISIS
# ============================================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📈 Resumen Ejecutivo",
    "🔥 Prioridad Compra",
    "📦 Sobrestock",
    "⚠️ Sin Costo",
    "💡 Oportunidades SIN ID",
    "🔍 Buscar Producto",
    "📊 Dashboard",
    "📝 Registro Compras",
])

# -------- TAB 1: RESUMEN --------
with tab1:
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### Distribución ABC")
        abc_data = modelo.groupby("CLASIFICACION_ABC").agg(
            Productos=("Id", "count"),
            Demanda=("DEMANDA_TOTAL", "sum"),
        ).reset_index()
        # Forzar A, B, C con map explícito (a prueba de balas)
        def _normalizar_abc(x):
            s = str(x).strip().upper()
            # Por si el valor viene mal por algún encoding raro
            if s in ("A",) or "A" == s[:1]: return "A"
            if s in ("B",) or "B" == s[:1]: return "B"
            return "C"  # cualquier otro caso lo forzamos a C
        abc_data["CLASIFICACION_ABC"] = abc_data["CLASIFICACION_ABC"].apply(_normalizar_abc)
        # Agrupar de nuevo en caso de que hubiera duplicados después del mapeo
        abc_data = abc_data.groupby("CLASIFICACION_ABC", as_index=False).agg({
            "Productos": "sum",
            "Demanda": "sum",
        })
        # Asegurar orden A, B, C
        orden_abc = {"A": 1, "B": 2, "C": 3}
        abc_data["_ord"] = abc_data["CLASIFICACION_ABC"].map(orden_abc)
        abc_data = abc_data.sort_values("_ord").drop(columns="_ord").reset_index(drop=True)
        # Demanda como entero, % como entero
        abc_data["Demanda"] = abc_data["Demanda"].astype(int)
        abc_data["% Demanda"] = (abc_data["Demanda"] / abc_data["Demanda"].sum() * 100).round(0).astype(int)
        # Formato con separador de miles
        abc_data_display = abc_data.copy()
        abc_data_display["Productos"] = abc_data_display["Productos"].apply(lambda x: f"{x:,}".replace(",", "."))
        abc_data_display["Demanda"] = abc_data_display["Demanda"].apply(lambda x: f"{x:,}".replace(",", "."))
        abc_data_display["% Demanda"] = abc_data_display["% Demanda"].astype(str) + "%"
        abc_data_display = abc_data_display.rename(columns={"CLASIFICACION_ABC": "Clase"})
        st.table(abc_data_display.set_index("Clase"))

    with col2:
        st.markdown("#### Estado del inventario")
        est_data = modelo["ESTADO_INVENTARIO"].value_counts().reset_index()
        est_data.columns = ["Estado", "Productos"]
        est_data["Productos"] = est_data["Productos"].apply(lambda x: f"{x:,}".replace(",", "."))
        st.table(est_data.set_index("Estado"))

    st.markdown("#### Estado por clasificación ABC")
    # Normalizar columna ABC antes del crosstab
    def _normalizar_abc(x):
        s = str(x).strip().upper()
        if s in ("A",) or "A" == s[:1]: return "A"
        if s in ("B",) or "B" == s[:1]: return "B"
        return "C"
    abc_norm = modelo["CLASIFICACION_ABC"].apply(_normalizar_abc)
    cross = pd.crosstab(abc_norm, modelo["ESTADO_INVENTARIO"])
    cross.index.name = "Clase"
    # Ordenar A, B, C explícitamente
    orden = ["A", "B", "C"]
    cross = cross.reindex([x for x in orden if x in cross.index])
    # Formato con separador
    cross_display = cross.map(lambda x: f"{x:,}".replace(",", "."))
    st.table(cross_display)

    # Alertas críticas
    st.markdown("---")
    st.markdown("### 🚨 Alertas de gestión")

    a_criticos = ((modelo["CLASIFICACION_ABC"] == "A") &
                  (modelo["ESTADO_INVENTARIO"] == "CRÍTICO")).sum()
    sin_costo = ((modelo["ESTADO_INVENTARIO"].isin(["CRÍTICO", "COMPRAR"])) &
                 modelo["ULTIMO_PRECIO_NETO"].isna()).sum()

    if a_criticos > 0:
        st.error(f"🔴 **{a_criticos} productos clase A en estado CRÍTICO** — pérdida de venta inminente. Ver tab 'Prioridad Compra'.")
    if sin_costo > 0:
        st.warning(f"🟡 **{sin_costo} productos críticos sin costo registrado** — gestionar con proveedor. Ver tab 'Sin Costo'.")
    if len(ranking_sin_id) > 0:
        st.info(f"💡 **{len(ranking_sin_id)} productos solicitados sin catálogo** — oportunidad de ampliar surtido. Ver tab 'Oportunidades SIN ID'.")


# -------- TAB 2: PRIORIDAD COMPRA --------
with tab2:
    st.markdown("### Productos a comprar este mes")

    prioridad = modelo[
        (modelo["ESTADO_INVENTARIO"].isin(["CRÍTICO", "COMPRAR"]))
        & (modelo["COMPRA_AJUSTADA"] > 0)
    ].copy()

    # Filtros
    col1, col2, col3 = st.columns(3)
    with col1:
        f_abc = st.multiselect("Filtrar ABC", ["A", "B", "C"], default=["A", "B", "C"])
    with col2:
        f_estado = st.multiselect("Filtrar estado", ["CRÍTICO", "COMPRAR"], default=["CRÍTICO", "COMPRAR"])
    with col3:
        marcas_disp = ["Todas"] + sorted(prioridad["Marca"].dropna().unique().tolist())
        f_marca = st.selectbox("Filtrar marca", marcas_disp)

    p = prioridad[
        prioridad["CLASIFICACION_ABC"].isin(f_abc)
        & prioridad["ESTADO_INVENTARIO"].isin(f_estado)
    ]
    if f_marca != "Todas":
        p = p[p["Marca"] == f_marca]

    # Dos secciones diferenciadas
    p_con_costo = p[p["ULTIMO_PRECIO_NETO"].notna()].sort_values(
        ["CLASIFICACION_ABC", "ESTADO_INVENTARIO", "COSTO_COMPRA_AJUSTADA"],
        ascending=[True, True, False],
    )
    p_sin_costo = p[p["ULTIMO_PRECIO_NETO"].isna()].sort_values(
        ["CLASIFICACION_ABC", "ESTADO_INVENTARIO", "DEMANDA_MENSUAL"],
        ascending=[True, True, False],
    )

    cols_show = ["Id", "Marca", "Producto", "CLASIFICACION_ABC", "ESTADO_INVENTARIO",
                 "DEMANDA_MENSUAL", "STOCK_ACTUAL", "COBERTURA_MESES",
                 "MESES_FINAL", "COMPRA_AJUSTADA", "ULTIMO_PRECIO_NETO",
                 "COSTO_COMPRA_AJUSTADA", "ULTIMO_PROVEEDOR"]
    col_cfg = {
        "DEMANDA_MENSUAL": st.column_config.NumberColumn("Demanda mes", format="%.1f"),
        "COBERTURA_MESES": st.column_config.NumberColumn("Cobertura", format="%.2f meses"),
        "ULTIMO_PRECIO_NETO": st.column_config.NumberColumn("Costo unit", format="$%d"),
        "COSTO_COMPRA_AJUSTADA": st.column_config.NumberColumn("Inversión", format="$%d"),
        "MESES_FINAL": st.column_config.NumberColumn("Meses", format="%.1f"),
        "COMPRA_AJUSTADA": st.column_config.NumberColumn("Comprar", format="%d"),
        "STOCK_ACTUAL": st.column_config.NumberColumn("Stock"),
        "CLASIFICACION_ABC": "ABC",
        "ESTADO_INVENTARIO": "Estado",
    }

    # ---- Sección 1: Listos para comprar ----
    st.success(
        f"✅ **Sección 1: LISTOS PARA COMPRAR** — {len(p_con_costo):,} productos | "
        f"Inversión total: **${p_con_costo['COSTO_COMPRA_AJUSTADA'].sum():,.0f} CLP**".replace(",", ".")
    )
    st.dataframe(p_con_costo[cols_show], use_container_width=True, hide_index=True, column_config=col_cfg)

    # ---- Sección 2: Pendientes de costo ----
    st.markdown("---")
    st.error(
        f"⚠️ **Sección 2: PENDIENTES DE COSTO** — {len(p_sin_costo):,} productos con cantidad calculada pero "
        "sin monto (falta levantar costo con proveedor). La cantidad se calculó usando la cobertura objetivo ABC."
    )
    st.dataframe(p_sin_costo[cols_show], use_container_width=True, hide_index=True, column_config=col_cfg)


# -------- TAB 3: SOBRESTOCK --------
with tab3:
    st.markdown("### Productos en sobrestock (capital inmovilizado)")

    sob = modelo[modelo["ESTADO_INVENTARIO"] == "SOBRESTOCK"].copy()
    sob["EXCESO_UNIDADES"] = (sob["STOCK_ACTUAL"] - sob["DEMANDA_MENSUAL"] * sob["COBERTURA_OBJETIVO"])
    sob["VALOR_EXCESO"] = sob["EXCESO_UNIDADES"] * sob["ULTIMO_PRECIO_NETO"]
    sob = sob.sort_values("VALOR_EXCESO", ascending=False)

    col1, col2 = st.columns(2)
    with col1:
        marcas_sob = ["Todas"] + sorted(sob["Marca"].dropna().unique().tolist())
        f_marca = st.selectbox("Filtrar marca", marcas_sob, key="marca_sob")
    with col2:
        min_valor = st.number_input("Valor exceso mínimo (CLP)", value=0, step=100000)

    s = sob.copy()
    if f_marca != "Todas":
        s = s[s["Marca"] == f_marca]
    s = s[s["VALOR_EXCESO"].fillna(0) >= min_valor]

    st.markdown(f"**{len(s):,} productos** — Capital exceso: **${s['VALOR_EXCESO'].sum():,.0f}**")

    st.dataframe(
        s[["Id", "Marca", "Producto", "CLASIFICACION_ABC", "STOCK_ACTUAL",
           "DEMANDA_MENSUAL", "COBERTURA_MESES", "EXCESO_UNIDADES",
           "ULTIMO_PRECIO_NETO", "VALOR_EXCESO"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "DEMANDA_MENSUAL": st.column_config.NumberColumn("Demanda mes", format="%.1f"),
            "COBERTURA_MESES": st.column_config.NumberColumn("Cobertura", format="%.1f meses"),
            "EXCESO_UNIDADES": st.column_config.NumberColumn("Exceso uds", format="%.1f"),
            "ULTIMO_PRECIO_NETO": st.column_config.NumberColumn("Costo unit", format="$%d"),
            "VALOR_EXCESO": st.column_config.NumberColumn("Valor exceso", format="$%d"),
            "CLASIFICACION_ABC": "ABC",
            "STOCK_ACTUAL": "Stock",
        },
    )


# -------- TAB 4: SIN COSTO --------
with tab4:
    st.markdown("### Productos sin costo registrado")
    st.caption("Productos en CRÍTICO o COMPRAR que no tienen último costo. Gestionar urgente con proveedor.")

    sc = modelo[
        (modelo["ESTADO_INVENTARIO"].isin(["CRÍTICO", "COMPRAR"]))
        & modelo["ULTIMO_PRECIO_NETO"].isna()
    ].sort_values(["CLASIFICACION_ABC", "DEMANDA_MENSUAL"], ascending=[True, False])

    st.markdown(f"**{len(sc):,} productos requieren levantar costo**")

    st.dataframe(
        sc[["Id", "Marca", "Producto", "SKU", "CLASIFICACION_ABC",
            "DEMANDA_MENSUAL", "STOCK_ACTUAL", "COBERTURA_MESES",
            "ESTADO_INVENTARIO", "Precio AILOO"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "DEMANDA_MENSUAL": st.column_config.NumberColumn("Demanda mes", format="%.1f"),
            "COBERTURA_MESES": st.column_config.NumberColumn("Cobertura", format="%.2f"),
            "Precio AILOO": st.column_config.NumberColumn("Precio venta", format="$%d"),
            "CLASIFICACION_ABC": "ABC",
            "ESTADO_INVENTARIO": "Estado",
            "STOCK_ACTUAL": "Stock",
        },
    )


# -------- TAB 5: SIN ID --------
with tab5:
    st.markdown("### Productos solicitados que NO existen en el catálogo")
    st.caption("Demanda real que no se está capturando como venta. Oportunidad de ampliar surtido.")

    top_n = st.slider("Mostrar top N", 50, 500, 100, 50)
    st.markdown(f"**{len(ranking_sin_id):,} productos únicos solicitados sin catálogo**")

    st.dataframe(
        ranking_sin_id.head(top_n),
        use_container_width=True,
        hide_index=True,
        column_config={
            "PRODUCTO_DESCRIPCION": "Producto solicitado",
            "VECES_SOLICITADO": st.column_config.NumberColumn("Veces solicitado", format="%d"),
        },
    )


# -------- TAB 6: BUSCAR --------
with tab6:
    st.markdown("### Buscador de producto")
    busqueda = st.text_input("Busca por ID, SKU, descripción o marca", placeholder="Ej: filtro aceite, 3447716, etc.")

    if busqueda:
        b = busqueda.upper().strip()
        m = modelo[
            modelo["Producto"].astype(str).str.upper().str.contains(b, na=False)
            | modelo["SKU"].astype(str).str.upper().str.contains(b, na=False)
            | modelo["Marca"].astype(str).str.upper().str.contains(b, na=False)
            | modelo["Id"].astype(str).str.contains(b, na=False)
        ]

        st.markdown(f"**{len(m):,} resultados**")

        if len(m) > 0:
            st.dataframe(
                m[["Id", "Marca", "Producto", "SKU", "CLASIFICACION_ABC",
                   "ESTADO_INVENTARIO", "DEMANDA_MENSUAL", "STOCK_ACTUAL",
                   "COBERTURA_MESES", "COMPRA_AJUSTADA", "ULTIMO_PRECIO_NETO",
                   "ULTIMO_PROVEEDOR"]],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "DEMANDA_MENSUAL": st.column_config.NumberColumn("Demanda mes", format="%.1f"),
                    "COBERTURA_MESES": st.column_config.NumberColumn("Cobertura", format="%.2f"),
                    "ULTIMO_PRECIO_NETO": st.column_config.NumberColumn("Último costo", format="$%d"),
                    "ULTIMO_PROVEEDOR": st.column_config.TextColumn("Último proveedor"),
                    "CLASIFICACION_ABC": "ABC",
                    "ESTADO_INVENTARIO": "Estado",
                    "STOCK_ACTUAL": "Stock",
                    "COMPRA_AJUSTADA": st.column_config.NumberColumn("Comprar", format="%d"),
                },
            )


# -------- TAB 7: DASHBOARD EJECUTIVO --------
with tab7:
    st.markdown("## 📊 Dashboard Ejecutivo de Compras")
    st.caption("Vista consolidada para toma de decisiones rápida")
    st.markdown("---")

    # =========================================================================
    # BLOQUE 1: KPIs FINANCIEROS
    # =========================================================================
    st.markdown("### 💰 Indicadores financieros")

    # Cálculos
    modelo_d = modelo.copy()
    modelo_d["VALOR_STOCK"] = modelo_d["STOCK_ACTUAL"] * modelo_d["ULTIMO_PRECIO_NETO"]
    sobre_d = modelo_d[modelo_d["ESTADO_INVENTARIO"] == "SOBRESTOCK"].copy()
    sobre_d["EXCESO"] = (sobre_d["STOCK_ACTUAL"] - sobre_d["DEMANDA_MENSUAL"] * sobre_d["COBERTURA_OBJETIVO"])
    sobre_d["VALOR_EXCESO"] = sobre_d["EXCESO"] * sobre_d["ULTIMO_PRECIO_NETO"]

    inv_sugerida = a_comprar["COSTO_COMPRA_AJUSTADA"].sum()
    cap_sobrestock = sobre_d["VALOR_EXCESO"].sum()
    valor_stock_total = modelo_d["VALOR_STOCK"].sum()
    brecha = cap_sobrestock - inv_sugerida  # cuánto se podría liberar vs lo que hay que invertir

    # Helper para formato chileno
    def fmt_clp(x):
        return f"${x:,.0f}".replace(",", ".")

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.markdown(f"""<div class="kpi-card">
            <p class="kpi-card-title">Inversión a realizar</p>
            <p class="kpi-card-value">{fmt_clp(inv_sugerida)}</p>
            <p class="kpi-card-subtitle">{len(a_comprar):,} productos listos</p>
        </div>""".replace("{:,}".format(len(a_comprar)), "{:,}".format(len(a_comprar)).replace(",", ".")), unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card">
            <p class="kpi-card-title">Capital atrapado</p>
            <p class="kpi-card-value">{fmt_clp(cap_sobrestock)}</p>
            <p class="kpi-card-subtitle">{len(sobre_d):,} en sobrestock</p>
        </div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""<div class="kpi-card">
            <p class="kpi-card-title">Valor stock total</p>
            <p class="kpi-card-value">{fmt_clp(valor_stock_total)}</p>
            <p class="kpi-card-subtitle">Inventario actual</p>
        </div>""", unsafe_allow_html=True)
    with k4:
        color_brecha = "#2E7D32" if brecha > 0 else "#C00000"
        signo = "+" if brecha > 0 else ""
        st.markdown(f"""<div class="kpi-card" style="border-left-color: {color_brecha};">
            <p class="kpi-card-title">Brecha financiera</p>
            <p class="kpi-card-value" style="color: {color_brecha};">{signo}{fmt_clp(brecha)}</p>
            <p class="kpi-card-subtitle">Sobrestock - Inversión</p>
        </div>""", unsafe_allow_html=True)

    if brecha > 0:
        st.info(f"💡 **Oportunidad:** Liberando solo parte del sobrestock se podría financiar la compra urgente.")
    else:
        st.warning(f"⚠️ **Atención:** La inversión necesaria supera el capital atrapado en sobrestock.")

    st.markdown("---")

    # =========================================================================
    # BLOQUE 2: DISTRIBUCIÓN ABC Y ESTADOS (donuts)
    # =========================================================================
    st.markdown("### 🥧 Distribución del portafolio")

    col_abc, col_est = st.columns(2)

    with col_abc:
        st.markdown("**Por clasificación ABC**")
        abc_data = modelo.groupby("CLASIFICACION_ABC").agg(
            Productos=("Id", "count"),
            Demanda=("DEMANDA_TOTAL", "sum"),
        ).reset_index()
        abc_data["% Demanda"] = (abc_data["Demanda"] / abc_data["Demanda"].sum() * 100).round(1)

        # Crear donut con Plotly via plotly_chart o usar tabla con barras
        try:
            import plotly.graph_objects as go
            fig_abc = go.Figure(data=[go.Pie(
                labels=[f"Clase {x}" for x in abc_data["CLASIFICACION_ABC"]],
                values=abc_data["Productos"],
                hole=0.55,
                marker=dict(colors=["#8B0000", "#C77C7C", "#E8C5C5"]),
                textinfo='label+percent',
                textposition='outside',
            )])
            fig_abc.update_layout(
                showlegend=False,
                margin=dict(l=10, r=10, t=10, b=10),
                height=300,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#262730'),
            )
            st.plotly_chart(fig_abc, use_container_width=True)
        except ImportError:
            st.dataframe(abc_data, use_container_width=True, hide_index=True)

    with col_est:
        st.markdown("**Por estado de inventario**")
        est_data = modelo["ESTADO_INVENTARIO"].value_counts().reset_index()
        est_data.columns = ["Estado", "Productos"]

        try:
            import plotly.graph_objects as go
            # Colores por estado
            colores_est = {"CRÍTICO": "#C00000", "COMPRAR": "#FFC000", "OK": "#2E7D32", "SOBRESTOCK": "#A0A0A0"}
            fig_est = go.Figure(data=[go.Pie(
                labels=est_data["Estado"],
                values=est_data["Productos"],
                hole=0.55,
                marker=dict(colors=[colores_est.get(e, "#888") for e in est_data["Estado"]]),
                textinfo='label+percent',
                textposition='outside',
            )])
            fig_est.update_layout(
                showlegend=False,
                margin=dict(l=10, r=10, t=10, b=10),
                height=300,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#262730'),
            )
            st.plotly_chart(fig_est, use_container_width=True)
        except ImportError:
            st.dataframe(est_data, use_container_width=True, hide_index=True)

    st.markdown("---")

    # =========================================================================
    # BLOQUE 3: TOP PRODUCTOS QUE REQUIEREN ATENCIÓN
    # =========================================================================
    st.markdown("### 🚨 Productos que requieren tu atención HOY")

    top1, top2, top3 = st.columns(3)

    with top1:
        st.markdown("**🔴 TOP 5 Clase A en CRÍTICO**")
        st.caption("Mayor pérdida de venta")
        top_critico_a = modelo[
            (modelo["CLASIFICACION_ABC"] == "A") &
            (modelo["ESTADO_INVENTARIO"] == "CRÍTICO")
        ].nlargest(5, "DEMANDA_MENSUAL")[["Producto", "DEMANDA_MENSUAL", "STOCK_ACTUAL"]]
        top_critico_a.columns = ["Producto", "Dem/mes", "Stock"]
        st.dataframe(
            top_critico_a,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Dem/mes": st.column_config.NumberColumn(format="%.1f"),
                "Stock": st.column_config.NumberColumn(format="%d"),
            }
        )

    with top2:
        st.markdown("**📦 TOP 5 Sobrestock (capital)**")
        st.caption("Mayor capital atrapado")
        top_sobre = sobre_d.nlargest(5, "VALOR_EXCESO")[["Producto", "EXCESO", "VALOR_EXCESO"]]
        top_sobre.columns = ["Producto", "Exceso", "Valor (CLP)"]
        st.dataframe(
            top_sobre,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Exceso": st.column_config.NumberColumn(format="%.0f"),
                "Valor (CLP)": st.column_config.NumberColumn(format="$%d"),
            }
        )

    with top3:
        st.markdown("**⚠️ TOP 5 Sin Costo (Clase A)**")
        st.caption("Gestionar con proveedor")
        top_sincosto = modelo[
            (modelo["CLASIFICACION_ABC"] == "A") &
            (modelo["ESTADO_INVENTARIO"].isin(["CRÍTICO", "COMPRAR"])) &
            modelo["ULTIMO_PRECIO_NETO"].isna()
        ].nlargest(5, "DEMANDA_MENSUAL")[["Producto", "DEMANDA_MENSUAL", "STOCK_ACTUAL"]]
        top_sincosto.columns = ["Producto", "Dem/mes", "Stock"]
        st.dataframe(
            top_sincosto,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Dem/mes": st.column_config.NumberColumn(format="%.1f"),
                "Stock": st.column_config.NumberColumn(format="%d"),
            }
        )

    st.markdown("---")

    # =========================================================================
    # BLOQUE 4: TOP MARCAS PROBLEMÁTICAS
    # =========================================================================
    st.markdown("### 📈 Análisis por marca")

    col_m1, col_m2 = st.columns(2)

    with col_m1:
        st.markdown("**Top 10 marcas con más productos CRÍTICOS**")
        marcas_criticas = modelo[modelo["ESTADO_INVENTARIO"] == "CRÍTICO"].groupby("Marca").size().reset_index(name="Productos")
        marcas_criticas = marcas_criticas.nlargest(10, "Productos")

        try:
            import plotly.graph_objects as go
            fig_mc = go.Figure(go.Bar(
                x=marcas_criticas["Productos"],
                y=marcas_criticas["Marca"],
                orientation='h',
                marker=dict(color="#C00000"),
                text=marcas_criticas["Productos"],
                textposition='outside',
            ))
            fig_mc.update_layout(
                yaxis=dict(autorange="reversed"),
                margin=dict(l=10, r=30, t=10, b=10),
                height=350,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#262730'),
                xaxis=dict(showgrid=True, gridcolor='#EEEEEE'),
            )
            st.plotly_chart(fig_mc, use_container_width=True)
        except ImportError:
            st.dataframe(marcas_criticas, use_container_width=True, hide_index=True)

    with col_m2:
        st.markdown("**Top 10 marcas con más capital atrapado**")
        marcas_sobre = sobre_d.groupby("Marca")["VALOR_EXCESO"].sum().reset_index()
        marcas_sobre = marcas_sobre.nlargest(10, "VALOR_EXCESO")
        marcas_sobre["VALOR_M"] = marcas_sobre["VALOR_EXCESO"] / 1_000_000

        try:
            import plotly.graph_objects as go
            fig_ms = go.Figure(go.Bar(
                x=marcas_sobre["VALOR_M"],
                y=marcas_sobre["Marca"],
                orientation='h',
                marker=dict(color="#8B0000"),
                text=[f"${v:.1f}M" for v in marcas_sobre["VALOR_M"]],
                textposition='outside',
            ))
            fig_ms.update_layout(
                yaxis=dict(autorange="reversed"),
                margin=dict(l=10, r=60, t=10, b=10),
                height=350,
                paper_bgcolor='rgba(0,0,0,0)',
                plot_bgcolor='rgba(0,0,0,0)',
                font=dict(color='#262730'),
                xaxis=dict(showgrid=True, gridcolor='#EEEEEE', title="Millones CLP"),
            )
            st.plotly_chart(fig_ms, use_container_width=True)
        except ImportError:
            marcas_sobre["Capital (M CLP)"] = marcas_sobre["VALOR_M"].round(1)
            st.dataframe(marcas_sobre[["Marca", "Capital (M CLP)"]], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.caption("💡 Tip: pasa el mouse sobre los gráficos para ver valores exactos. Clic en las leyendas para filtrar.")



# ============================================================================
# TAB 8: REGISTRO DE COMPRAS
# ============================================================================
with tab8:
    import csv as _csv
    from pathlib import Path as _Path

    st.markdown("## 📝 Registro de compras")
    st.caption("Anota lo que vas comprando y la app te muestra el estado de cada "
               "producto frente a lo sugerido por el modelo.")

    # --- Ubicacion del archivo donde se guardan las anotaciones ---
    try:
        _BASE_DIR = _Path(__file__).parent
    except NameError:
        _BASE_DIR = _Path.cwd()
    REGISTRO_PATH = _BASE_DIR / "registro_compras.csv"

    # --- Helpers ---
    def _cargar_registro():
        cols = ["fecha", "Id", "unidades", "precio_neto"]
        if REGISTRO_PATH.exists():
            try:
                reg = pd.read_csv(REGISTRO_PATH, dtype={"Id": str})
                if reg.empty or "Id" not in reg.columns:
                    return pd.DataFrame(columns=cols)
                reg["Id"] = reg["Id"].astype(str).str.strip()
                reg["unidades"] = pd.to_numeric(reg["unidades"], errors="coerce").fillna(0).astype(int)
                # compatibilidad con registros viejos que no tenian precio
                if "precio_neto" not in reg.columns:
                    reg["precio_neto"] = 0
                reg["precio_neto"] = pd.to_numeric(reg["precio_neto"], errors="coerce").fillna(0).astype(int)
                return reg[cols]
            except Exception:
                return pd.DataFrame(columns=cols)
        return pd.DataFrame(columns=cols)

    def _guardar_compra(id_prod, unidades, precio_neto):
        # Se reescribe el archivo completo (asi se migra solo si era formato viejo)
        try:
            reg = _cargar_registro()
            nueva = pd.DataFrame([{
                "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Id": str(id_prod).strip(),
                "unidades": int(unidades),
                "precio_neto": int(precio_neto),
            }])
            reg = pd.concat([reg, nueva], ignore_index=True)
            reg.to_csv(REGISTRO_PATH, index=False, encoding="utf-8")
            return True, ""
        except PermissionError:
            return False, ("No se pudo guardar: el archivo registro_compras.csv parece estar "
                           "ABIERTO en Excel u otro programa. Ciérralo e intenta de nuevo.")
        except Exception as e:
            return False, f"No se pudo guardar el registro. Detalle: {e}"

    # Ids validos del modelo (para no anotar un Id que no existe)
    _modelo_ids = modelo["Id"].astype(str).str.strip()
    _set_ids = set(_modelo_ids)

    # ------------------------------------------------------------------
    # 1) FORMULARIO PARA ANOTAR UNA COMPRA
    # ------------------------------------------------------------------
    st.markdown("#### ➕ Anotar una compra")
    with st.form("form_registro_compra", clear_on_submit=True):
        c1, c2, c3 = st.columns([2, 1, 1.3])
        with c1:
            id_input = st.text_input("Id del producto", placeholder="Ej: 2396575")
        with c2:
            unidades_input = st.number_input("Unidades compradas", min_value=1, step=1, value=1)
        with c3:
            precio_input = st.number_input("Precio neto unitario (CLP)", min_value=0, step=100, value=0,
                                           help="Precio neto por unidad de esta compra")
        enviado = st.form_submit_button("💾 Guardar compra")

    if enviado:
        id_limpio = str(id_input).strip()
        if id_limpio == "":
            st.error("Escribe el Id del producto.")
        elif id_limpio not in _set_ids:
            st.error(f"El Id {id_limpio} no existe en el modelo. Revisa el numero.")
        else:
            ok_guardar, msg_guardar = _guardar_compra(id_limpio, unidades_input, precio_input)
            if ok_guardar:
                _nombre = modelo.loc[_modelo_ids == id_limpio, "Producto"].iloc[0]
                _precio_txt = (f" a ${int(precio_input):,}".replace(",", ".") + " c/u") if int(precio_input) > 0 else ""
                st.success(f"✅ Anotadas {int(unidades_input)} unidades de {id_limpio} — {_nombre}{_precio_txt}")
            else:
                st.error("⚠️ " + msg_guardar)

    # ------------------------------------------------------------------
    # Funcion para eliminar TODAS las compras anotadas de un producto
    # ------------------------------------------------------------------
    def _eliminar_producto(id_prod):
        """Borra del registro todas las filas de un producto y reescribe el archivo."""
        reg = _cargar_registro()
        if reg.empty:
            return False
        antes = len(reg)
        reg = reg[reg["Id"].astype(str).str.strip() != str(id_prod).strip()].reset_index(drop=True)
        if len(reg) == antes:
            return False
        try:
            reg.to_csv(REGISTRO_PATH, index=False, encoding="utf-8")
            return True
        except PermissionError:
            st.error("⚠️ No se pudo eliminar: cierra el archivo registro_compras.csv si lo "
                     "tienes abierto en Excel e intenta de nuevo.")
            return False

    def _fijar_comprado(id_prod, nueva_cant):
        """Corrige la cantidad comprada de un producto: reemplaza sus anotaciones
        por una sola con la cantidad correcta. Si la cantidad es 0, lo elimina.
        Conserva el ultimo precio neto que se habia registrado para ese producto."""
        idp = str(id_prod).strip()
        reg = _cargar_registro()
        # precio neto que tenia la ultima anotacion de este producto (se conserva)
        previas = reg[reg["Id"].astype(str).str.strip() == idp]
        precio_prev = int(previas["precio_neto"].iloc[-1]) if len(previas) else 0
        # quitar todo lo anotado de este producto
        reg = reg[reg["Id"].astype(str).str.strip() != idp]
        # dejar una sola anotacion con la cantidad corregida (si es > 0)
        if int(nueva_cant) > 0:
            nueva = pd.DataFrame([{
                "fecha": datetime.now().strftime("%Y-%m-%d %H:%M"),
                "Id": idp,
                "unidades": int(nueva_cant),
                "precio_neto": precio_prev,
            }])
            reg = pd.concat([reg, nueva], ignore_index=True)
        reg = reg.reset_index(drop=True)
        try:
            reg.to_csv(REGISTRO_PATH, index=False, encoding="utf-8")
            return True
        except PermissionError:
            st.error("⚠️ No se pudo guardar la corrección: cierra el archivo "
                     "registro_compras.csv si lo tienes abierto en Excel e intenta de nuevo.")
            return False

    # Estado para pedir confirmacion antes de borrar (evita borrados accidentales)
    if "confirmar_borrado_id" not in st.session_state:
        st.session_state.confirmar_borrado_id = None
    if "confirmar_vaciar" not in st.session_state:
        st.session_state.confirmar_vaciar = False

    def _vaciar_registro():
        """Deja el registro de compras en cero (borra todas las anotaciones)."""
        try:
            pd.DataFrame(columns=["fecha", "Id", "unidades", "precio_neto"]).to_csv(
                REGISTRO_PATH, index=False, encoding="utf-8")
            return True
        except PermissionError:
            st.error("⚠️ No se pudo vaciar: cierra el archivo registro_compras.csv si lo "
                     "tienes abierto en Excel e intenta de nuevo.")
            return False

    def _rerun():
        try:
            st.rerun()
        except Exception:
            st.experimental_rerun()

    # ------------------------------------------------------------------
    # 2) ESTADO DE COMPRAS POR PRODUCTO
    # ------------------------------------------------------------------
    st.markdown("#### 📋 Estado de compras por producto")

    registro = _cargar_registro()

    # Base: productos sugeridos por el modelo (COMPRA_AJUSTADA > 0) +
    # cualquier producto que tenga compras anotadas, aunque el modelo no lo haya sugerido
    base = modelo.copy()
    base["Id"] = base["Id"].astype(str).str.strip()
    ids_con_compras = set(registro["Id"].astype(str).str.strip()) if not registro.empty else set()
    mask_base = (base["COMPRA_AJUSTADA"] > 0) | (base["Id"].isin(ids_con_compras))
    base = base[mask_base][["Id", "Marca", "Producto", "COMPRA_AJUSTADA"]].copy()
    base = base.rename(columns={"COMPRA_AJUSTADA": "Sugerido"})
    base["Sugerido"] = pd.to_numeric(base["Sugerido"], errors="coerce").fillna(0).round(0).astype(int)

    # Unidades ya compradas por Id
    if not registro.empty:
        comprado = registro.groupby("Id", as_index=False)["unidades"].sum()
        comprado = comprado.rename(columns={"unidades": "Comprado"})
        # ultimo precio neto registrado por producto (la anotacion mas reciente)
        precio_ult = registro.groupby("Id", as_index=False)["precio_neto"].last()
        precio_ult = precio_ult.rename(columns={"precio_neto": "PrecioNeto"})
    else:
        comprado = pd.DataFrame(columns=["Id", "Comprado"])
        precio_ult = pd.DataFrame(columns=["Id", "PrecioNeto"])

    tabla = base.merge(comprado, on="Id", how="left").merge(precio_ult, on="Id", how="left")
    tabla["Comprado"] = tabla["Comprado"].fillna(0).astype(int)
    tabla["PrecioNeto"] = tabla["PrecioNeto"].fillna(0).astype(int)

    def _estado(row):
        if row["Comprado"] == 0:
            return "⚪ Pendiente"
        if row["Sugerido"] == 0:
            return "🟣 Sin sugerencia"
        if row["Comprado"] < row["Sugerido"]:
            return "🟡 Parcial"
        if row["Comprado"] == row["Sugerido"]:
            return "🟢 Completa"
        return "🔵 Excedida"
    tabla["Estado"] = tabla.apply(_estado, axis=1)

    # --- Resumen rapido ---
    n_completa = (tabla["Estado"] == "🟢 Completa").sum()
    n_parcial  = (tabla["Estado"] == "🟡 Parcial").sum()
    n_excedida = (tabla["Estado"] == "🔵 Excedida").sum()
    n_anotados = int((tabla["Comprado"] > 0).sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Con compras anotadas", f"{n_anotados:,}".replace(",", "."))
    m2.metric("🟢 Completas", f"{n_completa:,}".replace(",", "."))
    m3.metric("🟡 Parciales", f"{n_parcial:,}".replace(",", "."))
    m4.metric("🔵 Excedidas", f"{n_excedida:,}".replace(",", "."))

    # --- Total capital invertido ---
    total_capital = int((tabla["Comprado"] * tabla["PrecioNeto"]).sum())
    if total_capital > 0:
        capital_fmt = "$" + f"{total_capital:,}".replace(",", ".")
        st.markdown(
            f'<div style="background-color:#8B0000; border-radius:10px; padding:14px 24px; '
            f'margin:12px 0 4px 0; display:inline-block;">'
            f'<span style="color:#FFFFFF; font-size:1rem; font-weight:600;">'
            f'💰 Capital total invertido este mes:&nbsp;&nbsp;</span>'
            f'<span style="color:#FFFFFF; font-size:1.4rem; font-weight:800;">{capital_fmt}</span>'
            f'</div>',
            unsafe_allow_html=True,
        )

    # --- Filtros ---
    fc1, fc2 = st.columns([2, 1])
    with fc1:
        buscar = st.text_input("Buscar por Id o nombre", placeholder="Escribe para filtrar...",
                               key="buscar_registro")
    with fc2:
        solo_anotados = st.checkbox("Solo productos con compras anotadas", value=True)

    vista = tabla.copy()
    if solo_anotados:
        vista = vista[vista["Comprado"] > 0]
    if buscar.strip():
        b = buscar.strip().lower()
        vista = vista[
            vista["Id"].str.lower().str.contains(b)
            | vista["Producto"].str.lower().str.contains(b)
            | vista["Marca"].astype(str).str.lower().str.contains(b)
        ]

    vista = vista[["Id", "Marca", "Producto", "Sugerido", "Comprado", "PrecioNeto", "Estado"]]

    if vista.empty:
        st.info("Aun no hay compras anotadas que mostrar. Usa el formulario de arriba "
                "o desmarca 'Solo productos con compras anotadas' para ver todo.")
    elif len(vista) > 80:
        # Lista muy grande: se muestra tabla simple sin boton de borrar (se borra solo lo anotado)
        st.dataframe(
            vista, use_container_width=True, hide_index=True,
            column_config={
                "PrecioNeto": st.column_config.NumberColumn("P. neto unit.", format="$%d"),
            },
        )
        st.caption("ℹ️ Para poder eliminar con la ✕, marca 'Solo productos con compras "
                   "anotadas' o busca un producto puntual.")
    else:
        # Tabla interactiva: 'Comprado' es editable, con boton de guardar y de eliminar por fila
        st.caption("✏️ Puedes corregir la cantidad en la columna **Comprado** y pulsar 💾 para "
                   "guardar. Para borrar el registro de un producto, usa 🗑️.")
        _anchos = [1.2, 1.3, 3.2, 0.9, 1.1, 1.2, 1.2, 0.7, 0.7]
        h = st.columns(_anchos)
        for _c, _t in zip(h, ["Id", "Marca", "Producto", "Sugerido", "Comprado",
                              "P. neto unit.", "Estado", "", ""]):
            _c.markdown(f"**{_t}**")
        st.markdown("<hr style='margin:2px 0 6px 0'>", unsafe_allow_html=True)

        for _, fila in vista.iterrows():
            rid = str(fila["Id"]).strip()
            comprado_actual = int(fila["Comprado"])
            c = st.columns(_anchos, vertical_alignment="center")
            c[0].write(rid)
            c[1].write(str(fila["Marca"]))
            c[2].write(str(fila["Producto"]))
            c[3].write(f'{int(fila["Sugerido"]):,}'.replace(",", "."))
            # campo editable con la cantidad comprada
            nueva_cant = c[4].number_input(
                "Comprado", min_value=0, step=1, value=comprado_actual,
                key=f"comp_{rid}", label_visibility="collapsed",
            )
            # precio neto unitario (ultimo registrado)
            _pn = int(fila["PrecioNeto"])
            c[5].write(("$" + f"{_pn:,}".replace(",", ".")) if _pn > 0 else "—")
            c[6].write(fila["Estado"])

            # boton GUARDAR (solo se activa si cambio la cantidad)
            if c[7].button("💾", key=f"save_{rid}", help="Guardar la cantidad corregida"):
                if int(nueva_cant) != comprado_actual:
                    _fijar_comprado(rid, int(nueva_cant))
                    _rerun()
                else:
                    st.toast("No cambiaste la cantidad de ese producto.")

            # boton ELIMINAR con confirmacion en dos pasos
            if st.session_state.confirmar_borrado_id == rid:
                if c[8].button("✅", key=f"ok_{rid}", help="Confirmar eliminación"):
                    _eliminar_producto(rid)
                    st.session_state.confirmar_borrado_id = None
                    _rerun()
                if c[8].button("✖", key=f"no_{rid}", help="Cancelar"):
                    st.session_state.confirmar_borrado_id = None
                    _rerun()
            else:
                if c[8].button("🗑️", key=f"del_{rid}", help="Eliminar este registro"):
                    st.session_state.confirmar_borrado_id = rid
                    _rerun()

        if st.session_state.confirmar_borrado_id is not None:
            st.warning("⚠️ Confirma con ✅ o cancela con ✖ en la fila marcada para eliminar el registro.")

    st.caption("✏️ Corregir cantidad: edita la columna Comprado y pulsa 💾.  "
               "🗑️ Eliminar: borra el registro de ese producto (con confirmación).  "
               "📁 Todo se guarda en registro_compras.csv (en la carpeta de la app).")

    # ------------------------------------------------------------------
    # 3) VACIAR TODO EL REGISTRO (para empezar de cero cuando quieras)
    # ------------------------------------------------------------------
    st.markdown("---")
    with st.expander("🧹 Vaciar todo el registro de compras"):
        st.caption("Borra TODAS las compras anotadas y deja el registro en cero. "
                   "Útil cuando empiezas un mes o quieres partir limpio. "
                   "Esta acción no se puede deshacer.")

        _reg_actual = _cargar_registro()
        if not _reg_actual.empty:
            st.download_button(
                "⬇️ Descargar respaldo antes de vaciar (CSV)",
                data=_reg_actual.to_csv(index=False).encode("utf-8"),
                file_name=f"registro_compras_respaldo_{datetime.now().strftime('%Y-%m-%d')}.csv",
                mime="text/csv",
                key="descargar_respaldo_registro",
            )

        if not st.session_state.confirmar_vaciar:
            if st.button("🧹 Vaciar registro", key="btn_vaciar"):
                st.session_state.confirmar_vaciar = True
                _rerun()
        else:
            st.warning("⚠️ ¿Seguro? Se borrarán TODAS las compras anotadas y el listado quedará en cero.")
            vc1, vc2 = st.columns(2)
            if vc1.button("✅ Sí, vaciar todo", key="vaciar_ok"):
                if _vaciar_registro():
                    st.session_state.confirmar_vaciar = False
                    _rerun()
            if vc2.button("✖ Cancelar", key="vaciar_no"):
                st.session_state.confirmar_vaciar = False
                _rerun()


# ============================================================================
# FOOTER
# ============================================================================
st.markdown("---")
st.caption("🚗 Prototipo Modelo de Compras | Datos cargados desde archivos Excel | "
           "En la app definitiva, datos vendrán directamente de la API de Ailoo  ·  v1.7")
