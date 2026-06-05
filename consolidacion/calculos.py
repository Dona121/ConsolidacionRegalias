"""
Cálculos en Python (no fórmulas Excel) para las columnas derivadas que
aplican tanto a la hoja principal como a la de descentralizadas:

- DIAS DESDE LA APROBACIÓN HASTA APERTURA DEL PRIMER PROCESO
- DIAS DESDE LA APERTURA HASTA LA FIRMA DEL PRIMER CONTRATO
- DIAS DESDE LA FECHA DE SUSCRIPCIÓN HASTA LA FECHA DEL ACTA DE INICIO

Nota: CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN ya NO se calcula aquí; ahora
migra de la versión anterior de la Matriz (ver `procesamiento.consolidar_h1`).
"""

from __future__ import annotations

import polars as pl


# ── Días transcurridos ────────────────────────────────────────────────────────

def dias_desde_aprobacion_hasta_primer_proceso(
    estado_proyecto: str,
    fecha_aprobacion: str,
    fecha_corte_gesproy: str,
):
    """Solo aplica a proyectos en estado 'SIN CONTRATAR'."""
    condicion = (
        (pl.col(estado_proyecto) == "SIN CONTRATAR")
        & pl.col(fecha_aprobacion).is_not_null()
        & pl.col(fecha_corte_gesproy).is_not_null()
    )
    return (
        pl.when(condicion)
        .then((pl.col(fecha_corte_gesproy) - pl.col(fecha_aprobacion)).dt.total_days())
        .otherwise(None)
    )


def dias_desde_apertura_hasta_primer_contrato(
    estado_proyecto: str,
    fecha_acta_inicio: str,
    fecha_primer_proceso: str,
):
    """Solo aplica a proyectos en estado 'SIN CONTRATAR' con fecha de apertura."""
    condicion = (
        (pl.col(estado_proyecto) == "SIN CONTRATAR")
        & pl.col(fecha_primer_proceso).is_not_null()
    )
    return (
        pl.when(condicion)
        .then((pl.col(fecha_acta_inicio) - pl.col(fecha_primer_proceso)).dt.total_days())
        .otherwise(None)
    )


def dias_desde_suscripcion_hasta_fecha_acta_inicio(
    estado_proyecto: str,
    fecha_corte_gesproy: str,
    fecha_suscripcion: str,
):
    """Solo aplica a proyectos en estado 'CONTRATADO SIN ACTA DE INICIO'."""
    condicion = pl.col(estado_proyecto) == "CONTRATADO SIN ACTA DE INICIO"
    return (
        pl.when(condicion)
        .then((pl.col(fecha_corte_gesproy) - pl.col(fecha_suscripcion)).dt.total_days())
        .otherwise(None)
    )


# ── Agregar todos los cálculos a un DataFrame ────────────────────────────────

def agregar_calculos(df: pl.DataFrame) -> pl.DataFrame:
    """
    Agrega las tres columnas de días calculadas en Python al DataFrame.

    Se usa tanto para la Hoja 1 como para la Hoja 2 (Descentralizadas), ya
    que ambas comparten estos cálculos y los mismos nombres de columna.

    CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN ya no se calcula aquí: ahora
    migra de la versión anterior de la Matriz.
    """
    return df.with_columns(
        dias_desde_aprobacion_hasta_primer_proceso(
            "ESTADO PROYECTO", "FECHA APROBACIÓN PROYECTO", "FECHA DE CORTE GESPROY"
        ).alias("DIAS DESDE LA APROBACIÓN HASTA APERTURA DEL PRIMER PROCESO"),
        dias_desde_apertura_hasta_primer_contrato(
            "ESTADO PROYECTO", "FECHA ACTA INICIO", "FECHA DE APERTURA DEL PRIMER PROCESO"
        ).alias("DIAS DESDE LA APERTURA HASTA LA FIRMA DEL PRIMER CONTRATO"),
        dias_desde_suscripcion_hasta_fecha_acta_inicio(
            "ESTADO PROYECTO", "FECHA DE CORTE GESPROY",
            "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL"
        ).alias("DIAS DESDE LA FECHA DE SUSCRIPCIÓN HASTA LA FECHA DEL ACTA DE INICIO"),
    )


__all__ = [
    "dias_desde_aprobacion_hasta_primer_proceso",
    "dias_desde_apertura_hasta_primer_contrato",
    "dias_desde_suscripcion_hasta_fecha_acta_inicio",
    "agregar_calculos",
]
