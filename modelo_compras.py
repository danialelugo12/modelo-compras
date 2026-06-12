"""
================================================================================
MODELO DE COMPRAS - GENERADOR AUTOMÁTICO
================================================================================
Analiza demanda, cobertura, quiebres y sobrestock para generar un plan de
compras inteligente desde 4 archivos fuente (Maestro, Ventas, Fallidas, Compras).

Uso:
    python modelo_compras.py
    python modelo_compras.py --input ./INPUT --output ./OUTPUT
    python modelo_compras.py --config config.yaml

Autor: Analista Senior de Abastecimiento
Versión: 1.0
================================================================================
"""

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from config import CONFIG


# ============================================================================
# LOGGING
# ============================================================================
def setup_logger(output_dir: Path) -> logging.Logger:
    """Configura logger que escribe a archivo y consola."""
    log_file = output_dir / f"log_{datetime.now():%Y%m%d_%H%M%S}.txt"
    logger = logging.getLogger("modelo_compras")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()

    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%H:%M:%S")
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(fh)
    logger.addHandler(ch)
    return logger


# ============================================================================
# CARGA Y VALIDACIÓN
# ============================================================================
def conv_fecha_excel(v):
    """Convierte un valor a fecha, manejando seriales de Excel y strings."""
    if pd.isna(v):
        return pd.NaT
    if isinstance(v, (int, float)):
        try:
            return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))
        except Exception:
            return pd.NaT
    if isinstance(v, pd.Timestamp):
        return v
    return pd.to_datetime(v, errors="coerce")


def _normalizar(nombre: str) -> str:
    """Normaliza nombre de archivo: sin extensión, sin acentos, sin espacios/guiones,
    todo en mayúsculas. Permite hacer match flexible entre variaciones."""
    import unicodedata
    s = Path(nombre).stem  # quita extensión
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()  # quita acentos
    s = s.upper()
    # Quitar todo lo que no sea letra o número
    s = "".join(c for c in s if c.isalnum())
    return s


def _buscar_archivo(input_dir: Path, patrones: list, logger: logging.Logger):
    """Busca en input_dir un archivo .xlsx cuyo nombre normalizado contenga TODOS los
    patrones dados. Devuelve la primera coincidencia o None.

    Ejemplos de match para patrones ['VENTAS','6','MESES']:
      - VENTAS_6_MESES.xlsx       ✓
      - VENTAS 6 MESES.xlsx       ✓
      - ventas-6-meses.xlsx       ✓
      - Ventas6Meses.xlsx         ✓
    """
    patrones_norm = [_normalizar(p) for p in patrones]
    candidatos = []
    for f in input_dir.glob("*.xlsx"):
        if f.name.startswith("~$"):  # archivos temporales de Excel abiertos
            continue
        nombre_norm = _normalizar(f.name)
        if all(p in nombre_norm for p in patrones_norm):
            candidatos.append(f)

    if not candidatos:
        return None
    if len(candidatos) > 1:
        logger.warning(
            f"  ⚠️ Múltiples archivos coinciden con {patrones}, tomando el más reciente: "
            f"{[c.name for c in candidatos]}"
        )
        candidatos.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return candidatos[0]


def cargar_archivos(input_dir: Path, logger: logging.Logger) -> dict:
    """Carga los 4 archivos fuente buscándolos por patrones flexibles en el nombre.

    Acepta cualquier variación de mayúsculas/minúsculas, espacios, guiones bajos,
    guiones, acentos, etc. Por ejemplo, todos estos son equivalentes:
      - VENTAS_6_MESES.xlsx
      - Ventas 6 Meses.xlsx
      - VENTAS-6-MESES.xlsx
    """
    # Patrones que deben aparecer en el nombre normalizado de cada archivo
    patrones_busqueda = {
        "maestro":  ["MAESTRO"],
        "ventas":   ["VENTAS"],          # busca cualquier archivo con "VENTAS"...
        "fallidas": ["VENTA", "FALLIDA"],
        "compras":  ["COMPRAS"],
    }

    archivos = {}
    faltantes = []
    for clave, patrones in patrones_busqueda.items():
        # Evitar que VENTAS atrape al de fallidas: si VENTAS está en patrón, excluir los que tengan FALLIDA
        encontrado = _buscar_archivo(input_dir, patrones, logger)
        if encontrado is None:
            faltantes.append((clave, patrones))
        else:
            archivos[clave] = encontrado

    # Refinar: si el de "ventas" tiene "FALLIDA" en el nombre, en realidad es el de fallidas
    if "ventas" in archivos and "FALLIDA" in _normalizar(archivos["ventas"].name):
        # buscar otro archivo con VENTAS que NO tenga FALLIDA
        otros = [
            f for f in input_dir.glob("*.xlsx")
            if not f.name.startswith("~$")
            and "VENTAS" in _normalizar(f.name)
            and "FALLIDA" not in _normalizar(f.name)
        ]
        if otros:
            archivos["ventas"] = otros[0]
        else:
            del archivos["ventas"]
            faltantes.append(("ventas", ["VENTAS"]))

    if faltantes:
        logger.error(f"Archivos faltantes en {input_dir}:")
        logger.error("  El script no pudo encontrar archivos para:")
        for clave, patrones in faltantes:
            logger.error(f"  - {clave}: necesita un archivo .xlsx con {patrones} en el nombre")
        logger.error("  Archivos disponibles en INPUT:")
        for f in input_dir.glob("*.xlsx"):
            if not f.name.startswith("~$"):
                logger.error(f"    • {f.name}")
        raise FileNotFoundError(f"Faltan archivos: {[k for k, _ in faltantes]}")

    logger.info("Cargando archivos fuente...")
    for clave, ruta in archivos.items():
        logger.info(f"  ✓ {clave}: {ruta.name}")

    data = {}
    # Maestro: del COMPRAS.xlsx hoja "MAESTRO ..." si existe, sino del archivo maestro
    try:
        xls_compras = pd.ExcelFile(archivos["compras"])
        hoja_maestro = next(
            (h for h in xls_compras.sheet_names if "MAESTRO" in h.upper()), None
        )
        if hoja_maestro:
            data["maestro"] = pd.read_excel(archivos["compras"], sheet_name=hoja_maestro)
            logger.info(f"  Maestro tomado de {archivos['compras'].name} hoja '{hoja_maestro}'")
        else:
            data["maestro"] = pd.read_excel(archivos["maestro"])
            logger.info(f"  Maestro tomado de {archivos['maestro'].name}")
    except Exception as e:
        logger.warning(f"  No se pudo leer maestro de COMPRAS, usando archivo aparte: {e}")
        data["maestro"] = pd.read_excel(archivos["maestro"])

    data["ventas"] = pd.read_excel(archivos["ventas"])
    data["fallidas"] = pd.read_excel(archivos["fallidas"])

    # En compras: buscar hoja "COMPRA NACIONAL" o variantes
    xls_c = pd.ExcelFile(archivos["compras"])
    hoja_compra = next(
        (h for h in xls_c.sheet_names if "COMPRA" in h.upper() and "NACIONAL" in h.upper()),
        None,
    ) or next(
        (h for h in xls_c.sheet_names if "COMPRA" in h.upper() and "MAESTRO" not in h.upper()),
        xls_c.sheet_names[0],
    )
    data["compras"] = pd.read_excel(archivos["compras"], sheet_name=hoja_compra)
    logger.info(f"  Compras tomado de hoja '{hoja_compra}'")

    logger.info(f"  Maestro:  {len(data['maestro']):>7,} productos")
    logger.info(f"  Ventas:   {len(data['ventas']):>7,} transacciones")
    logger.info(f"  Fallidas: {len(data['fallidas']):>7,} filas")
    logger.info(f"  Compras:  {len(data['compras']):>7,} compras")

    return data


def validar_columnas(data: dict, logger: logging.Logger):
    """Valida que cada archivo tenga las columnas esperadas."""
    requisitos = {
        "maestro": ["Id", "Marca", "Producto", "SKU", "Precio AILOO", "Stock Total"],
        "ventas": ["Fecha", "Tipo", "Canal", "Tienda", "Id producto", "Cantidad"],
        "fallidas": ["FECHA", "ID AILOO", "CANTIDAD", "PRODUCTO"],
        "compras": ["FECHA", "ID AILOO", "PRECIO NETO UNITARIO", "PEDIDO ", "LLEGADO"],
    }
    errores = []
    for fuente, cols in requisitos.items():
        df = data[fuente]
        faltan = [c for c in cols if c not in df.columns]
        if faltan:
            errores.append(f"{fuente}: faltan columnas {faltan}")
    if errores:
        for e in errores:
            logger.error(f"  {e}")
        raise ValueError("Estructura de archivos inválida. Revisa columnas.")
    logger.info("  ✓ Estructura de archivos OK")


# ============================================================================
# PROCESAMIENTO
# ============================================================================
def procesar_ventas(df: pd.DataFrame, logger: logging.Logger) -> pd.DataFrame:
    """Calcula demanda real por producto (incluye boleta + factura, todas las
    tiendas; las notas de crédito se restan solas porque vienen con cantidad
    negativa). Excluye proformas y canales no comerciales."""
    df = df.copy()
    df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")

    filt = (
        df["Tipo"].isin(CONFIG["ventas"]["tipos_validos"])
        & ~df["Canal"].isin(CONFIG["ventas"]["canales_excluir"])
    )
    df_ok = df[filt]
    descartadas = len(df) - len(df_ok)

    demanda = df_ok.groupby("Id producto")["Cantidad"].sum().reset_index()
    demanda.columns = ["Id", "VENTAS_6M"]

    logger.info(f"  Ventas válidas: {len(df_ok):,} (descartadas {descartadas:,})")
    logger.info(f"  Productos con ventas: {len(demanda):,}")
    return demanda


def procesar_fallidas(df: pd.DataFrame, logger: logging.Logger) -> tuple:
    """Procesa quiebres. Devuelve (demanda_con_id, oportunidades_sin_id)."""
    df = df.copy()
    df["CANTIDAD_OK"] = pd.to_numeric(df["CANTIDAD"], errors="coerce").fillna(
        CONFIG["fallidas"]["cantidad_default"]
    )

    sin_id_mask = (
        df["ID AILOO"].astype(str).str.upper().str.strip().isin(["SIN ID", "NAN", ""])
    )

    # Con ID válido
    con_id = df[~sin_id_mask].copy()
    con_id["ID AILOO"] = pd.to_numeric(con_id["ID AILOO"], errors="coerce")
    con_id = con_id.dropna(subset=["ID AILOO"])
    con_id["ID AILOO"] = con_id["ID AILOO"].astype(int)
    demanda_fall = con_id.groupby("ID AILOO")["CANTIDAD_OK"].sum().reset_index()
    demanda_fall.columns = ["Id", "FALLIDAS_6M"]

    # Sin ID: oportunidad comercial
    sin_id = df[sin_id_mask].copy()
    sin_id["PROD"] = sin_id["PRODUCTO"].astype(str).str.strip().str.upper()
    # Excluir descripciones genéricas / placeholders del formulario
    EXCLUIR = {
        "SIN ID", "NO ENCUENTRA", "NO ENCUENTRO", "NO SE ENCUENTRA",
        "NO HAY", "NO TENEMOS", "NO EXISTE", "PRODUCTO NO ENCONTRADO",
        "SIN DESCRIPCION", "SIN DESCRIPCIÓN", "NAN", "NONE", "",
    }
    sin_id = sin_id[~sin_id["PROD"].isin(EXCLUIR)]
    sin_id = sin_id[sin_id["PROD"].str.len() > 2]  # filtrar demasiado cortos
    ranking = sin_id["PROD"].value_counts().reset_index()
    ranking.columns = ["PRODUCTO_DESCRIPCION", "VECES_SOLICITADO"]

    logger.info(f"  Productos con quiebre (ID válido): {len(demanda_fall):,}")
    logger.info(f"  Solicitudes SIN ID (oportunidad): {len(sin_id):,}")
    return demanda_fall, ranking


def procesar_compras(df: pd.DataFrame, logger: logging.Logger) -> tuple:
    """Devuelve (último_costo_por_producto, stock_en_tránsito)."""
    df = df.copy()
    df["FECHA_OK"] = df["FECHA"].apply(conv_fecha_excel)
    df["ID_OK"] = pd.to_numeric(df["ID AILOO"], errors="coerce")
    df = df.dropna(subset=["ID_OK"])
    df["ID_OK"] = df["ID_OK"].astype(int)
    df["PRECIO_NETO"] = pd.to_numeric(df["PRECIO NETO UNITARIO"], errors="coerce")

    # Último costo: tomar la compra más reciente con precio > 0
    df_costo = df[(df["PRECIO_NETO"] > 0) & df["FECHA_OK"].notna()]
    ultimo = (
        df_costo.sort_values("FECHA_OK")
        .groupby("ID_OK")
        .agg(
            ULTIMO_PRECIO_NETO=("PRECIO_NETO", "last"),
            ULTIMA_FECHA_COMPRA=("FECHA_OK", "last"),
            ULTIMO_PROVEEDOR=("PROVEEDOR", "last"),
        )
        .reset_index()
        .rename(columns={"ID_OK": "Id"})
    )

    # Stock en tránsito = PEDIDO - LLEGADO (solo positivos)
    df["PEDIDO_N"] = pd.to_numeric(df["PEDIDO "], errors="coerce").fillna(0)
    df["LLEGADO_N"] = pd.to_numeric(df["LLEGADO"], errors="coerce").fillna(0)
    df["PENDIENTE"] = (df["PEDIDO_N"] - df["LLEGADO_N"]).clip(lower=0)
    transito = df.groupby("ID_OK")["PENDIENTE"].sum().reset_index()
    transito.columns = ["Id", "EN_TRANSITO"]
    transito = transito[transito["EN_TRANSITO"] > 0]

    logger.info(f"  Productos con costo reciente: {len(ultimo):,}")
    logger.info(f"  Productos con stock en tránsito: {len(transito):,}")
    return ultimo, transito


def construir_modelo(
    maestro: pd.DataFrame,
    ventas: pd.DataFrame,
    fallidas: pd.DataFrame,
    costos: pd.DataFrame,
    transito: pd.DataFrame,
    logger: logging.Logger,
) -> pd.DataFrame:
    """Construye el modelo consolidado con ABC, cobertura, estado y compra
    ajustada según costo."""
    base = maestro[["Id", "Marca", "Producto", "SKU", "Precio AILOO", "Stock Total"]].copy()
    base["Stock Total"] = pd.to_numeric(base["Stock Total"], errors="coerce").fillna(0)
    base["STOCK_ACTUAL"] = base["Stock Total"].clip(lower=0)

    base = base.merge(ventas, on="Id", how="left")
    base = base.merge(fallidas, on="Id", how="left")
    base = base.merge(costos, on="Id", how="left")
    base = base.merge(transito, on="Id", how="left")

    for col in ["VENTAS_6M", "FALLIDAS_6M", "EN_TRANSITO"]:
        base[col] = base[col].fillna(0)

    base["DEMANDA_TOTAL"] = base["VENTAS_6M"] + base["FALLIDAS_6M"]
    base["DEMANDA_MENSUAL"] = base["DEMANDA_TOTAL"] / CONFIG["periodo_meses"]

    # Solo productos con demanda entran al análisis ABC
    con = base[base["DEMANDA_TOTAL"] > 0].copy()
    con = con.sort_values("DEMANDA_TOTAL", ascending=False).reset_index(drop=True)

    # ABC
    total = con["DEMANDA_TOTAL"].sum()
    con["PCT_PARTICIPACION"] = con["DEMANDA_TOTAL"] / total
    con["PCT_ACUMULADO"] = con["PCT_PARTICIPACION"].cumsum()
    abc_cuts = CONFIG["abc"]
    con["CLASIFICACION_ABC"] = con["PCT_ACUMULADO"].apply(
        lambda p: "A" if p <= abc_cuts["A"] else ("B" if p <= abc_cuts["B"] else "C")
    )

    # Cobertura y estado
    con["COBERTURA_MESES"] = np.where(
        con["DEMANDA_MENSUAL"] > 0,
        con["STOCK_ACTUAL"] / con["DEMANDA_MENSUAL"],
        np.inf,
    )
    con["COBERTURA_OBJETIVO"] = con["CLASIFICACION_ABC"].map(CONFIG["cobertura_obj"])

    def tipo_abast(r):
        if r["COBERTURA_MESES"] < 1:
            return "Urgente"
        elif r["COBERTURA_MESES"] < r["COBERTURA_OBJETIVO"]:
            return "Pronto"
        else:
            return "OK"

    con["TIPO_ABASTECIMIENTO"] = con.apply(tipo_abast, axis=1)
    con["CANT_SUGERIDA_BASE"] = (
        con["DEMANDA_MENSUAL"] * con["COBERTURA_OBJETIVO"]
        - con["STOCK_ACTUAL"]
        - con["EN_TRANSITO"]
    )

    def estado(r):
        if r["CANT_SUGERIDA_BASE"] <= 0:
            return "SOBRESTOCK"
        elif r["COBERTURA_MESES"] < 1:
            return "CRÍTICO"
        elif r["COBERTURA_MESES"] < r["COBERTURA_OBJETIVO"]:
            return "COMPRAR"
        else:
            return "OK"

    con["ESTADO_INVENTARIO"] = con.apply(estado, axis=1)

    # Ajuste por costo
    def meses_segun_costo(c):
        if pd.isna(c) or c == 0:
            return np.nan
        for cost_max, meses in CONFIG["reglas_costo"]:
            if c <= cost_max:
                return meses
        return CONFIG["reglas_costo"][-1][1]

    con["MESES_A_COMPRAR"] = con["ULTIMO_PRECIO_NETO"].apply(meses_segun_costo)

    # MESES_FINAL: si hay costo usa la regla por costo; si no, usa la cobertura objetivo del ABC
    con["MESES_FINAL"] = con["MESES_A_COMPRAR"].fillna(con["COBERTURA_OBJETIVO"])

    # COMPRA_AJUSTADA siempre se calcula (haya o no costo)
    con["COMPRA_AJUSTADA"] = np.ceil(
        (
            con["DEMANDA_MENSUAL"] * con["MESES_FINAL"]
            - con["STOCK_ACTUAL"]
            - con["EN_TRANSITO"]
        ).clip(lower=0)
    )

    # El monto solo se puede calcular si hay costo
    con["COSTO_COMPRA_AJUSTADA"] = con["COMPRA_AJUSTADA"] * con["ULTIMO_PRECIO_NETO"]

    # Observación para diferenciar las filas que no tienen costo
    con["OBSERVACION_COMPRA"] = np.where(
        con["ULTIMO_PRECIO_NETO"].isna() & (con["COMPRA_AJUSTADA"] > 0),
        "⚠️ Costo pendiente de levantar con proveedor",
        np.where(
            con["COMPRA_AJUSTADA"] > 0,
            "Listo para comprar",
            "",
        ),
    )

    logger.info(f"  Productos con demanda: {len(con):,}")
    logger.info(f"  Productos clase A: {(con['CLASIFICACION_ABC']=='A').sum():,}")
    logger.info(f"  Productos en CRÍTICO: {(con['ESTADO_INVENTARIO']=='CRÍTICO').sum():,}")
    logger.info(f"  Productos en SOBRESTOCK: {(con['ESTADO_INVENTARIO']=='SOBRESTOCK').sum():,}")
    return con


# ============================================================================
# EXPORTACIÓN EXCEL (delegada a un módulo aparte para mantener claridad)
# ============================================================================
from exportar_excel import generar_excel  # noqa: E402


# ============================================================================
# MAIN
# ============================================================================
def main():
    parser = argparse.ArgumentParser(description="Genera el modelo de compras mensual.")
    parser.add_argument("--input", default="INPUT", help="Carpeta con archivos fuente")
    parser.add_argument("--output", default="OUTPUT", help="Carpeta donde escribir resultados")
    args = parser.parse_args()

    input_dir = Path(args.input)
    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    logger = setup_logger(output_dir)
    logger.info("=" * 70)
    logger.info("MODELO DE COMPRAS - INICIO DE EJECUCIÓN")
    logger.info("=" * 70)

    try:
        # 1. Cargar
        data = cargar_archivos(input_dir, logger)
        validar_columnas(data, logger)

        # 2. Procesar cada fuente
        logger.info("\nProcesando ventas...")
        ventas = procesar_ventas(data["ventas"], logger)

        logger.info("\nProcesando ventas fallidas...")
        fallidas, ranking_sin_id = procesar_fallidas(data["fallidas"], logger)

        logger.info("\nProcesando compras (costos y tránsito)...")
        costos, transito = procesar_compras(data["compras"], logger)

        # 3. Construir modelo
        logger.info("\nConstruyendo modelo consolidado...")
        modelo = construir_modelo(
            data["maestro"], ventas, fallidas, costos, transito, logger
        )

        # 4. Exportar
        logger.info("\nGenerando archivo Excel...")
        fecha = datetime.now().strftime("%Y-%m-%d")
        output_file = output_dir / f"MODELO_COMPRAS_{fecha}.xlsx"
        generar_excel(modelo, ranking_sin_id, output_file, logger)

        logger.info("\n" + "=" * 70)
        logger.info("PROCESO COMPLETADO CON ÉXITO")
        logger.info(f"Archivo generado: {output_file}")
        logger.info("=" * 70)

    except Exception as e:
        logger.exception(f"Error en ejecución: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
