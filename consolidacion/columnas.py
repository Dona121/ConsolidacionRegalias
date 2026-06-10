"""
Definición de columnas, secciones y colores para las tres hojas del Excel.

Cada hoja tiene:
- Un orden de columnas (lista de strings).
- Una asignación columna → (color_celda, color_encabezado).
- Sets para clasificar columnas como fecha, número, días o con fórmula.

Esto se usa tanto para construir el DataFrame final como para escribir el Excel.

La paleta de colores y los formatos viven en `diseno.py`; aquí solo se importan
para construir el mapeo `columna → (color_celda, color_encabezado)`.
"""

# Paleta centralizada en diseno.py (se re-exporta por retrocompatibilidad).
from .diseno import (
    AZUL, AZUL_HEADER,
    NARANJA, NARANJA_HEADER,
    VERDE, VERDE_HEADER,
)


# ══════════════════════════════════════════════════════════════════════════════
# ── Hoja 1 — MatrizSeguimientoEvaluacion ──────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

columnas_datos_generales = [
    "BPIN", "ENTIDAD O SECRETARIA", "NOMBRE PROYECTO", "ALCANCE DEL PROYECTO",
    "SECTOR", "INDICADOR DE PRODUCTO MGA", "ESTADO PROYECTO", "ESTADO CONTRATO",
    "TIPO CONTRATO", "FUENTE DE FINANCIACIÓN", "VALOR SGR", "VALOR NACIÓN",
    "VALOR OTROS", "VALOR OTRAS FUENTES NO SUIFP", "VALOR TOTAL PROYECTO",
    "VALOR PAGOS", "ULTIMA FECHA PAGO", "FECHA DE MIGRACIÓN A GESPROY",
    "FECHA DE ASIGNACIÓN DE RECURSOS", "FECHA DE INCORPORACIÓN DE RECURSOS",
]

columnas_datos_calificacion = [
    "AVANCE FISICO", "AVANCE FINANCIERO", "CPI", "SPI",
    "FECHA APROBACIÓN PROYECTO", "FECHA DE APERTURA DEL PRIMER PROCESO",
    "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL",
    "FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO",
    "FECHA ACTA INICIO", "HORIZONTE DEL PROYECTO",
    "FECHA DE FINALIZACIÓN", "FECHA EN LA QUE PASO A ESTADO PARA CIERRE",
    "FECHA DE CORTE GESPROY",
    "INFORMACIÓN SOLICITADA", "INFORMACIÓN RECIBIDA",
    "FECHA DE RECIBO DE INFORMACIÓN",
    "DIAS DESDE LA APROBACIÓN HASTA APERTURA DEL PRIMER PROCESO",
    "DIAS DESDE LA APERTURA HASTA LA FIRMA DEL PRIMER CONTRATO",
    "DIAS DESDE LA FECHA DE SUSCRIPCIÓN HASTA LA FECHA DEL ACTA DE INICIO",
    "DESEMPEÑO EN EL CRONOGRAMA", "DESEMPEÑO EN EL COSTO",
    "COLUMNA APOYO", "BRECHA FISICO - FINANCIERA", "CONTROL EXTERNALIDADES",
]

columnas_evaluacion = [
    "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN",
    "CALIFICACIÓN INFORMACIÓN A TIEMPO",
    "CALIFICACIÓN CALIDAD INFORMACIÓN",
    "COLUMNA APOYO 2",
    "CALIFICACIÓN EJECUCIÓN DEL PROYECTO",
    # Índice que migra de la versión anterior (número decimal)
    "ÍNDICE DE EFICIENCIA EN LA TERMINACIÓN",
    "COMENTARIOS CALIFICACIÓN",
    # Columnas manuales nuevas (encabezado verde)
    "RESPONSABLE CARGUE EN GESPROY",
    "MUNICIPIOS",
]

todas_las_columnas = (
    columnas_datos_generales
    + columnas_datos_calificacion
    + columnas_evaluacion
)

# Mapping columna → (color_celda, color_encabezado).
# Para las columnas manuales (RESPONSABLE CARGUE EN GESPROY, MUNICIPIOS) se
# sobrescribe a verde aunque vivan dentro de la sección EVALUACIÓN (cuya barra
# de título sigue siendo azul).
color_por_columna = (
    {c: (AZUL, AZUL_HEADER) for c in columnas_datos_generales}
    | {c: (NARANJA, NARANJA_HEADER) for c in columnas_datos_calificacion}
    | {c: (AZUL, AZUL_HEADER) for c in columnas_evaluacion}
    | {"RESPONSABLE CARGUE EN GESPROY": (VERDE, VERDE_HEADER)}
    | {"MUNICIPIOS": (VERDE, VERDE_HEADER)}
)

columnas_fecha_h1 = {
    "ULTIMA FECHA PAGO", "FECHA APROBACIÓN PROYECTO",
    "FECHA DE APERTURA DEL PRIMER PROCESO",
    "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL",
    "FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO",
    "FECHA ACTA INICIO", "FECHA DE CORTE GESPROY",
    "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
    "FECHA DE INCORPORACIÓN DE RECURSOS",
    "HORIZONTE DEL PROYECTO", "FECHA DE FINALIZACIÓN",
    "FECHA EN LA QUE PASO A ESTADO PARA CIERRE",
    "FECHA DE RECIBO DE INFORMACIÓN",
}

columnas_numero_h1 = {
    "VALOR SGR", "VALOR NACIÓN", "VALOR OTROS", "VALOR OTRAS FUENTES NO SUIFP",
    "VALOR TOTAL PROYECTO", "VALOR PAGOS", "AVANCE FISICO", "AVANCE FINANCIERO",
    "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN",
    "CALIFICACIÓN CALIDAD INFORMACIÓN",
    "ÍNDICE DE EFICIENCIA EN LA TERMINACIÓN",
}

columnas_dias_h1 = {
    "DIAS DESDE LA APROBACIÓN HASTA APERTURA DEL PRIMER PROCESO",
    "DIAS DESDE LA APERTURA HASTA LA FIRMA DEL PRIMER CONTRATO",
    "DIAS DESDE LA FECHA DE SUSCRIPCIÓN HASTA LA FECHA DEL ACTA DE INICIO",
}

columnas_con_formula_h1 = {
    "DESEMPEÑO EN EL CRONOGRAMA", "DESEMPEÑO EN EL COSTO",
    "COLUMNA APOYO", "BRECHA FISICO - FINANCIERA",
    "CALIFICACIÓN INFORMACIÓN A TIEMPO", "COLUMNA APOYO 2",
    "CALIFICACIÓN EJECUCIÓN DEL PROYECTO",
}


# ══════════════════════════════════════════════════════════════════════════════
# ── Hoja 2 — OtrosEjecutoresDescentralizadas ──────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

cols_desc_generales = [
    "BPIN", "EJECUTOR", "NOMBRE DEL PROYECTO", "ALCANCE", "SECTOR", "FUENTE",
    "ESTADO PROYECTO", "ESTADO CONTRATO", "VALOR SGR", "VALOR OTROS",
    "VALOR TOTAL", "FECHA DE MIGRACIÓN A GESPROY",
    "FECHA DE ASIGNACIÓN DE RECURSOS",
    # Typo corregido: antes "RECUROS", ahora "RECURSOS"
    "FECHA DE INCORPORACIÓN DE RECURSOS",
]

cols_desc_calificacion = [
    "AVANCE FÍSICO", "AVANCE FINANCIERO",
    "CPI", "SPI",
    "FECHA APROBACIÓN PROYECTO", "FECHA DE APERTURA DEL PRIMER PROCESO",
    "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL",
    "FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO", "FECHA ACTA INICIO",
    # Nueva columna manual (no migra de Gesproy)
    "HORIZONTE DEL PROYECTO",
    # Fecha que migra de la versión anterior (estado para cierre)
    "FECHA EN LA QUE PASO A ESTADO PARA CIERRE",
    "FECHA DE CORTE GESPROY",
    "DIAS DESDE LA APROBACIÓN HASTA APERTURA DEL PRIMER PROCESO",
    "DIAS DESDE LA APERTURA HASTA LA FIRMA DEL PRIMER CONTRATO",
    "DIAS DESDE LA FECHA DE SUSCRIPCIÓN HASTA LA FECHA DEL ACTA DE INICIO",
    "DESEMPEÑO EN EL CRONOGRAMA", "DESEMPEÑO EN EL COSTO",
    "COLUMNA APOYO", "BRECHA FISICO - FINANCIERA", "CONTROL EXTERNALIDADES",
]

cols_desc_evaluacion = [
    "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN",
    "CALIFICACIÓN INFORMACIÓN A TIEMPO",
    "CALIFICACIÓN CALIDAD INFORMACIÓN",
    "COLUMNA APOYO 2",
    "CALIFICACIÓN EJECUCIÓN DEL PROYECTO",
    # Índice que migra de la versión anterior (número decimal)
    "ÍNDICE DE EFICIENCIA EN LA TERMINACIÓN",
    "COMENTARIOS CALIFICACIÓN",
    # Columna manual nueva (encabezado verde)
    "MUNICIPIOS",
]

todas_desc = cols_desc_generales + cols_desc_calificacion + cols_desc_evaluacion

color_desc = (
    {c: (AZUL, AZUL_HEADER) for c in cols_desc_generales}
    | {c: (NARANJA, NARANJA_HEADER) for c in cols_desc_calificacion}
    | {c: (AZUL, AZUL_HEADER) for c in cols_desc_evaluacion}
    | {"MUNICIPIOS": (VERDE, VERDE_HEADER)}
)

columnas_fecha_desc = {
    "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
    # Typo corregido
    "FECHA DE INCORPORACIÓN DE RECURSOS", "FECHA APROBACIÓN PROYECTO",
    "FECHA DE APERTURA DEL PRIMER PROCESO",
    "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL",
    "FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO",
    "FECHA ACTA INICIO", "FECHA DE CORTE GESPROY",
    # Nueva columna manual
    "HORIZONTE DEL PROYECTO",
    # Fecha que migra de la versión anterior (estado para cierre)
    "FECHA EN LA QUE PASO A ESTADO PARA CIERRE",
}

columnas_numero_desc = {
    "VALOR SGR", "VALOR OTROS", "VALOR TOTAL",
    "AVANCE FÍSICO", "AVANCE FINANCIERO",
    "ÍNDICE DE EFICIENCIA EN LA TERMINACIÓN",
}

columnas_dias_desc = {
    "DIAS DESDE LA APROBACIÓN HASTA APERTURA DEL PRIMER PROCESO",
    "DIAS DESDE LA APERTURA HASTA LA FIRMA DEL PRIMER CONTRATO",
    "DIAS DESDE LA FECHA DE SUSCRIPCIÓN HASTA LA FECHA DEL ACTA DE INICIO",
}

columnas_con_formula_desc = {
    "DESEMPEÑO EN EL CRONOGRAMA", "DESEMPEÑO EN EL COSTO",
    "COLUMNA APOYO", "BRECHA FISICO - FINANCIERA",
    "COLUMNA APOYO 2", "CALIFICACIÓN EJECUCIÓN DEL PROYECTO",
    # CALIFICACIÓN INFORMACIÓN A TIEMPO: es manual en esta hoja.
}


# ══════════════════════════════════════════════════════════════════════════════
# ── Hoja 3 — OtrosEjecutoresMunicipios ────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

cols_mun = [
    "BPIN", "EJECUTOR", "NOMBRE DEL PROYECTO", "ALCANCE", "SECTOR", "FUENTE",
    "ESTADO PROYECTO", "ESTADO CONTRATO", "VALOR SGR", "VALOR OTROS",
    "VALOR TOTAL", "FECHA APROBACIÓN PROYECTO",
    "FECHA DE ASIGNACIÓN DE RECURSOS",
    # Typo corregido
    "FECHA DE INCORPORACIÓN DE RECURSOS",
    "FECHA ACTA INICIO", "AVANCE FÍSICO", "AVANCE FINANCIERO", "COMENTARIOS",
    # Columna manual nueva (encabezado verde)
    "MUNICIPIOS",
]

color_mun = (
    {c: (AZUL, AZUL_HEADER) for c in cols_mun}
    | {"MUNICIPIOS": (VERDE, VERDE_HEADER)}
)

columnas_fecha_mun = {
    "FECHA APROBACIÓN PROYECTO", "FECHA DE ASIGNACIÓN DE RECURSOS",
    # Typo corregido
    "FECHA DE INCORPORACIÓN DE RECURSOS", "FECHA ACTA INICIO",
}

columnas_numero_mun = {
    "VALOR SGR", "VALOR OTROS", "VALOR TOTAL",
    "AVANCE FÍSICO", "AVANCE FINANCIERO",
}


# ══════════════════════════════════════════════════════════════════════════════
# ── Índices de columnas (para fórmulas Excel) ─────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

idx_h1   = {col: todas_las_columnas.index(col) for col in todas_las_columnas}
idx_desc = {col: todas_desc.index(col)          for col in todas_desc}


__all__ = [
    # Colores
    "AZUL", "AZUL_HEADER", "NARANJA", "NARANJA_HEADER", "VERDE", "VERDE_HEADER",
    # Hoja 1
    "columnas_datos_generales", "columnas_datos_calificacion",
    "columnas_evaluacion", "todas_las_columnas", "color_por_columna",
    "columnas_fecha_h1", "columnas_numero_h1", "columnas_dias_h1",
    "columnas_con_formula_h1", "idx_h1",
    # Hoja 2
    "cols_desc_generales", "cols_desc_calificacion", "cols_desc_evaluacion",
    "todas_desc", "color_desc", "columnas_fecha_desc", "columnas_numero_desc",
    "columnas_dias_desc", "columnas_con_formula_desc", "idx_desc",
    # Hoja 3
    "cols_mun", "color_mun", "columnas_fecha_mun", "columnas_numero_mun",
]
