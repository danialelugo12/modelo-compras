"""
Configuración del modelo de compras.
Modifica aquí los parámetros del negocio sin tocar la lógica.
"""

CONFIG = {
    # --- Período de análisis ---
    "periodo_meses": 6,  # ventana de demanda histórica

    # --- Archivos esperados en la carpeta INPUT ---
    "files": {
        "maestro": "MAESTRO.xlsx",
        "ventas": "VENTAS_6_MESES.xlsx",
        "fallidas": "Planilla_Venta_Fallida.xlsx",
        "compras": "COMPRAS.xlsx",
    },

    # --- Reglas de procesamiento de ventas ---
    "ventas": {
        "tipos_validos": ["BOLETA", "FACTURA"],  # NOTA_CREDITO se incluye solo si viene con cantidad negativa
        "canales_excluir": ["Distribucion", "Todos"],
    },

    # --- Reglas de procesamiento de fallidas ---
    "fallidas": {
        "cantidad_default": 1,  # cuando viene vacía la columna CANTIDAD
    },

    # --- Cortes ABC (% acumulado de demanda) ---
    "abc": {
        "A": 0.80,
        "B": 0.95,
        # C: > 0.95
    },

    # --- Cobertura objetivo (meses) por clasificación ABC ---
    "cobertura_obj": {
        "A": 2.5,
        "B": 2.0,
        "C": 1.0,
    },

    # --- Reglas de meses a comprar según costo unitario (CLP) ---
    # Lista de tuplas (costo_max, meses). Se evalúa en orden, toma la primera que cumpla.
    "reglas_costo": [
        (15000, 3),    # ≤ $15.000 → comprar para 3 meses
        (99000, 2),    # $15.001 - $99.000 → comprar para 2 meses
        (float("inf"), 1),  # ≥ $100.000 → comprar para 1 mes (fraccionar)
    ],

    # --- Configuración del Excel de salida ---
    "excel": {
        "max_filas_oportunidad_sin_id": 200,  # cuántos productos sin ID mostrar
    },
}
