"""
Lógica de consolidación: aplica los joins por BPIN, prioriza contratos y
fechas, y produce los tres DataFrames listos para escribir al Excel.

Esta es la pieza central del aplicativo. El resto de módulos solo entregan
funciones helper; aquí se orquesta el flujo:

    1. Filtra `regalias_proyectos` (excluye CERRADO y DESAPROBADO).
    2. Selecciona un contrato representativo por BPIN.
    3. Construye `_fecha_final` priorizando Gesproy sobre fechas manuales.
    4. Hace los left joins con BPIN como llave.
    5. Agrega cálculos en Python (días, calificación contratación).
    6. Devuelve los DataFrames finales y la lista de fechas conservadas.
"""

from __future__ import annotations

from datetime import datetime

import polars as pl

from .calculos import agregar_calculos
from .esquemas import (
    ESQUEMA_GESPROY_PROYECTOS,
    ESQUEMA_GESPROY_CONTRATOS,
    ESQUEMA_GESPROY_CARGUE,
    ESQUEMA_MATRIZ_H1,
)
from .lectura import normalizar_fecha


# ══════════════════════════════════════════════════════════════════════════════
# ── Helpers ───────────────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

# Tipos de contrato priorizados al elegir un contrato por BPIN
_TIPOS_PRIORITARIOS = {
    "Obra pública", "Consultoría",
    "Convenios de Cooperación", "Interadministrativos",
}
_TIPOS_SECUNDARIOS = {
    "Suministro",
    "Contratos o convenios con entidades sin ánimo de lucro",
}


def _seleccionar_contrato_por_bpin(regalias_contratos: pl.DataFrame) -> pl.DataFrame:
    """
    Selecciona un contrato representativo por BPIN. Dentro de cada BPIN se
    aplica priorización por tipo y, en caso de empate, mayor valor SGR.
    """
    contratos = []
    for (bpin,), grupo in regalias_contratos.group_by("BPIN"):
        grupo = grupo.sort("VALOR TOTAL FUENTES SGR", descending=True)
        prioritarios = grupo.filter(pl.col("TIPO CONTRATO").is_in(_TIPOS_PRIORITARIOS))
        if len(prioritarios) > 0:
            contratos.append(prioritarios.head(1))
            continue
        secundarios = grupo.filter(pl.col("TIPO CONTRATO").is_in(_TIPOS_SECUNDARIOS))
        contratos.append(secundarios.head(1) if len(secundarios) > 0 else grupo.head(1))
    if not contratos:
        return regalias_contratos.clear()  # DataFrame vacío con el mismo esquema
    return pl.concat(contratos)


def _corregir_typo_recursos(df: pl.DataFrame) -> pl.DataFrame:
    """
    Si el archivo viene con el typo "RECUROS" (sin S) y NO tiene la columna
    correcta "RECURSOS", se renombra para evitar romper la validación durante
    el periodo de transición.
    """
    typo_col    = "FECHA DE INCORPORACIÓN DE RECUROS"
    correcta    = "FECHA DE INCORPORACIÓN DE RECURSOS"
    if typo_col in df.columns and correcta not in df.columns:
        return df.rename({typo_col: correcta})
    return df


def _renombrar_fecha_suscripcion(df: pl.DataFrame) -> pl.DataFrame:
    """
    Transición: la columna "FECHA SUSCRIPCION" pasó a llamarse
    "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL". Si el archivo de versión
    anterior aún trae el nombre viejo (y no el nuevo), se renombra para no
    romper la validación ni la migración durante el periodo de transición.
    """
    viejo = "FECHA SUSCRIPCION"
    nuevo = "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL"
    if viejo in df.columns and nuevo not in df.columns:
        return df.rename({viejo: nuevo})
    return df


# ══════════════════════════════════════════════════════════════════════════════
# ── Preparación de fuentes Gesproy ────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def preparar_proyectos(df_proy_raw: pl.DataFrame) -> pl.DataFrame:
    """Filtra los proyectos activos y castea los valores numéricos."""
    return (
        df_proy_raw
        .select(list(ESQUEMA_GESPROY_PROYECTOS.keys()))
        .filter(~pl.col("ESTADO PROYECTO").is_in(["CERRADO", "DESAPROBADO"]))
        .with_columns(
            pl.col(
                "VALOR NACIÓN", "VALOR OTROS", "VALOR OTRAS FUENTES NO SUIFP",
                "VALOR TOTAL PROYECTO", "VALOR SGR", "VALOR PAGOS",
            ).cast(pl.Float64, strict=False)
        )
    )


def preparar_contratos(df_cttos_raw: pl.DataFrame) -> pl.DataFrame:
    """Normaliza fechas, renombra columnas y selecciona un contrato por BPIN."""
    regalias_contratos = normalizar_fecha(
        df_cttos_raw.select(list(ESQUEMA_GESPROY_CONTRATOS.keys())),
        ["FECHA INICIAL", "FECHA ACT ADTIVO APERT", "FECHA SUSCRIPCION", "ULTIMA FECHA PAGO"],
    ).rename({
        "FECHA ACT ADTIVO APERT": "FECHA DE APERTURA DEL PRIMER PROCESO",
        "FECHA INICIAL":          "FECHA ACTA INICIO",
    })
    # Fecha del PRIMER contrato suscrito por BPIN: mínimo de FECHA SUSCRIPCION
    # sobre TODOS los contratos del BPIN, calculado ANTES de elegir el contrato
    # principal (el principal puede no ser el primero que se suscribió). Luego se
    # renombra FECHA SUSCRIPCION → la del contrato principal que quede tras la
    # selección.
    regalias_contratos = regalias_contratos.with_columns(
        pl.col("FECHA SUSCRIPCION").min().over("BPIN")
        .alias("FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO")
    ).rename({"FECHA SUSCRIPCION": "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL"})
    return _seleccionar_contrato_por_bpin(regalias_contratos)


def preparar_cargue(df_carga_raw: pl.DataFrame) -> pl.DataFrame:
    """Castea avances a Float64 y normaliza la fecha de aprobación."""
    return normalizar_fecha(
        df_carga_raw
        .select(list(ESQUEMA_GESPROY_CARGUE.keys()))
        .with_columns(
            pl.col("AVANCE FISICO", "AVANCE FINANCIERO").cast(pl.Float64, strict=False)
        ),
        ["FECHA APROBACIÓN PROYECTO"],
    )


# ══════════════════════════════════════════════════════════════════════════════
# ── Preparación de la versión anterior ────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

# Columnas opcionales de H1 (pueden no estar en archivos generados con versiones
# anteriores del aplicativo). Se agregan con valor nulo/vacío si faltan, para
# garantizar retrocompatibilidad. Tipos:
_COLS_OPCIONALES_H1 = {
    "CPI":                              pl.lit("").cast(pl.String),
    "SPI":                              pl.lit("").cast(pl.String),
    "INFORMACIÓN SOLICITADA":           pl.lit("").cast(pl.String),
    "INFORMACIÓN RECIBIDA":             pl.lit("").cast(pl.String),
    "FECHA DE RECIBO DE INFORMACIÓN":   pl.lit(None).cast(pl.Date),
    "CONTROL EXTERNALIDADES":           pl.lit(None).cast(pl.Float64),
    "FECHA DE CORTE GESPROY":           pl.lit(None).cast(pl.Date),
    # Ahora migra de la versión anterior (ya no se calcula en consolidación)
    "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN": pl.lit(None).cast(pl.Float64),
    "CALIFICACIÓN CALIDAD INFORMACIÓN": pl.lit(None).cast(pl.Float64),
    # Fecha que migra de la versión anterior (estado para cierre)
    "FECHA EN LA QUE PASO A ESTADO PARA CIERRE": pl.lit(None).cast(pl.Date),
    "COMENTARIOS CALIFICACIÓN":         pl.lit("").cast(pl.String),
    # Nuevas columnas manuales; archivos antiguos no las tienen → cadena vacía.
    "RESPONSABLE CARGUE EN GESPROY":    pl.lit("").cast(pl.String),
    "MUNICIPIOS":                       pl.lit("").cast(pl.String),
}

# Columnas opcionales de Descentralizadas (la nueva HORIZONTE DEL PROYECTO no
# migra de Gesproy; si la versión anterior no la tiene, se crea vacía).
_COLS_OPCIONALES_DESC = {
    "HORIZONTE DEL PROYECTO": pl.lit(None).cast(pl.Date),
    # Fecha que migra de la versión anterior (estado para cierre)
    "FECHA EN LA QUE PASO A ESTADO PARA CIERRE": pl.lit(None).cast(pl.Date),
    # Fecha del primer contrato: en Desc migra de la versión anterior. Es
    # opcional para no romper la primera generación si el archivo aún no la
    # trae (se crea vacía y el equipo la diligencia en adelante).
    "FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO": pl.lit(None).cast(pl.Date),
    # Nueva columna manual; se conserva desde la versión anterior.
    "MUNICIPIOS":             pl.lit("").cast(pl.String),
}

# Columnas opcionales de Municipios (nueva columna manual).
_COLS_OPCIONALES_MUN = {
    "MUNICIPIOS": pl.lit("").cast(pl.String),
}

_FECHAS_H1 = [
    "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
    "FECHA DE INCORPORACIÓN DE RECURSOS", "FECHA DE FINALIZACIÓN",
    "HORIZONTE DEL PROYECTO", "FECHA DE RECIBO DE INFORMACIÓN",
    "FECHA EN LA QUE PASO A ESTADO PARA CIERRE",
    "FECHA DE CORTE GESPROY",
]

_FECHAS_DESC = [
    "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
    "FECHA DE INCORPORACIÓN DE RECURSOS",
    "FECHA APROBACIÓN PROYECTO", "FECHA DE APERTURA DEL PRIMER PROCESO",
    "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL",
    "FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO", "FECHA ACTA INICIO",
    "FECHA DE CORTE GESPROY",
    "HORIZONTE DEL PROYECTO",
    "FECHA EN LA QUE PASO A ESTADO PARA CIERRE",
]

_FECHAS_MUN = [
    "FECHA APROBACIÓN PROYECTO", "FECHA DE ASIGNACIÓN DE RECURSOS",
    "FECHA DE INCORPORACIÓN DE RECURSOS", "FECHA ACTA INICIO",
]

# Fechas manuales que viven en H1 y que se priorizan contra Gesproy (Gesproy gana).
_FECHAS_MANUALES_H1 = [
    "BPIN",
    "FECHA APROBACIÓN PROYECTO",
    "FECHA DE APERTURA DEL PRIMER PROCESO",
    "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL",
    "FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO",
    "FECHA ACTA INICIO",
    "ULTIMA FECHA PAGO",
]


def _completar_columnas_opcionales(df: pl.DataFrame, opcionales: dict) -> pl.DataFrame:
    """Agrega columnas opcionales que no existan, con valor nulo/vacío."""
    faltantes = {col: expr for col, expr in opcionales.items() if col not in df.columns}
    if faltantes:
        df = df.with_columns([expr.alias(col) for col, expr in faltantes.items()])
    return df


def preparar_version_anterior(
    df_h1_raw: pl.DataFrame,
    df_desc_raw: pl.DataFrame,
    df_mun_raw: pl.DataFrame,
) -> tuple[pl.DataFrame, pl.DataFrame, pl.DataFrame, pl.DataFrame]:
    """
    Normaliza las tres tablas de la versión anterior y resuelve typos /
    columnas opcionales.

    Devuelve:
        bpines_version_anterior : DataFrame con las columnas de contexto H1
        otros_desc              : DataFrame de descentralizadas normalizado
        otros_mun               : DataFrame de municipios normalizado
        fechas_manuales_h1      : DataFrame con las fechas manuales del H1
                                  (para el coalesce posterior contra Gesproy)
    """
    # Typo "RECUROS" → "RECURSOS" en Desc y Mun (transición)
    df_desc_raw = _corregir_typo_recursos(df_desc_raw)
    df_mun_raw  = _corregir_typo_recursos(df_mun_raw)

    # "FECHA SUSCRIPCION" → "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL"
    # (transición). En H1 ayuda al fallback manual; en Desc es la columna que
    # se migra y valida.
    df_h1_raw   = _renombrar_fecha_suscripcion(df_h1_raw)
    df_desc_raw = _renombrar_fecha_suscripcion(df_desc_raw)

    # Completar columnas que pueden faltar en archivos viejos
    df_h1_raw   = _completar_columnas_opcionales(df_h1_raw,   _COLS_OPCIONALES_H1)
    df_desc_raw = _completar_columnas_opcionales(df_desc_raw, _COLS_OPCIONALES_DESC)
    df_mun_raw  = _completar_columnas_opcionales(df_mun_raw,  _COLS_OPCIONALES_MUN)

    # Normalizar fechas en cada tabla
    bpines_version_anterior = normalizar_fecha(
        df_h1_raw.select(list(ESQUEMA_MATRIZ_H1.keys())),
        _FECHAS_H1,
    )
    otros_desc = normalizar_fecha(df_desc_raw, _FECHAS_DESC)
    otros_mun  = normalizar_fecha(df_mun_raw,  _FECHAS_MUN)

    # Fechas manuales (subset de H1 usado para resolver Gesproy > manual)
    cols_presentes = [c for c in _FECHAS_MANUALES_H1 if c in df_h1_raw.columns]
    fechas_manuales_h1 = normalizar_fecha(
        df_h1_raw.select(cols_presentes),
        [c for c in _FECHAS_MANUALES_H1[1:] if c in cols_presentes],
    )

    return bpines_version_anterior, otros_desc, otros_mun, fechas_manuales_h1


# ══════════════════════════════════════════════════════════════════════════════
# ── Priorización Gesproy > manual ─────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def construir_fecha_final(
    regalias_proyectos: pl.DataFrame,
    regalias_contratos: pl.DataFrame,
    regalias_cargue: pl.DataFrame,
    fechas_manuales_h1: pl.DataFrame,
) -> tuple[pl.DataFrame, list[dict]]:
    """
    Para cada fecha clave, combina la información de Gesproy con la fecha
    ingresada manualmente: Gesproy tiene prioridad; si no hay valor, se
    conserva la fecha manual.

    Devuelve:
        _fecha_final          : DataFrame con BPIN + las 5 fechas consolidadas
        fechas_conservadas    : Lista de dicts con las fechas que vinieron de
                                 la versión anterior porque Gesproy no las tenía,
                                 para mostrarlas en pantalla al usuario.
    """
    fuentes_gesproy = {
        "FECHA APROBACIÓN PROYECTO":                   regalias_cargue.select("BPIN", "FECHA APROBACIÓN PROYECTO"),
        "FECHA DE APERTURA DEL PRIMER PROCESO":        regalias_contratos.select("BPIN", "FECHA DE APERTURA DEL PRIMER PROCESO"),
        "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL": regalias_contratos.select("BPIN", "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL"),
        "FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO":    regalias_contratos.select("BPIN", "FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO"),
        "FECHA ACTA INICIO":                           regalias_contratos.select("BPIN", "FECHA ACTA INICIO"),
        "ULTIMA FECHA PAGO":                           regalias_contratos.select("BPIN", "ULTIMA FECHA PAGO"),
    }

    _fecha_final = regalias_proyectos.select("BPIN")
    fechas_conservadas: list[dict] = []
    bpin_nombre = regalias_proyectos.select("BPIN", "NOMBRE PROYECTO")

    for col, df_gesproy in fuentes_gesproy.items():
        # Para cada fecha:
        #   1. left join todos_bpin × Gesproy   → _gesproy (null si no existe)
        #   2. left join todos_bpin × manuales  → _manual  (null si no estaba)
        #   3. coalesce(_gesproy, _manual)      → prioriza Gesproy
        todos_bpin = _fecha_final.select("BPIN")

        con_gesproy = todos_bpin.join(
            df_gesproy.rename({col: "_gesproy"}),
            on="BPIN", how="left",
        )

        if col in fechas_manuales_h1.columns:
            con_manual = todos_bpin.join(
                fechas_manuales_h1.select("BPIN", pl.col(col).alias("_manual")),
                on="BPIN", how="left",
            )
        else:
            con_manual = todos_bpin.with_columns(pl.lit(None).cast(pl.Date).alias("_manual"))

        combinado = (
            con_gesproy
            .join(con_manual, on="BPIN", how="left")
            .select(
                "BPIN",
                pl.coalesce([pl.col("_gesproy"), pl.col("_manual")]).alias(col),
                pl.col("_manual").alias(f"_manual_{col}"),
                pl.col("_gesproy").alias(f"_gesproy_{col}"),
            )
        )

        # Detectar fechas que vienen de la versión anterior (Gesproy no las tenía)
        sin_gesproy_con_manual = combinado.filter(
            pl.col(f"_gesproy_{col}").is_null()
            & pl.col(f"_manual_{col}").is_not_null()
        )
        if len(sin_gesproy_con_manual) > 0:
            con_nombre = sin_gesproy_con_manual.join(bpin_nombre, on="BPIN", how="left")
            for row in con_nombre.iter_rows(named=True):
                fechas_conservadas.append({
                    "BPIN":             row["BPIN"],
                    "Nombre proyecto":  row.get("NOMBRE PROYECTO", ""),
                    "Columna":          col,
                    "Fecha conservada": str(row[f"_manual_{col}"]),
                })

        _fecha_final = _fecha_final.join(
            combinado.select("BPIN", col),
            on="BPIN", how="left",
        )

    return _fecha_final, fechas_conservadas


# ══════════════════════════════════════════════════════════════════════════════
# ── Consolidación final ───────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def consolidar_h1(
    regalias_proyectos: pl.DataFrame,
    bpines_version_anterior: pl.DataFrame,
    _fecha_final: pl.DataFrame,
    regalias_contratos: pl.DataFrame,
    regalias_cargue: pl.DataFrame,
    fecha_corte: pl.Expr,
) -> pl.DataFrame:
    """
    Construye el DataFrame para la Hoja 1 (MatrizSeguimientoEvaluacion).

    Los `pl.lit("")` para las columnas con fórmulas Excel son placeholders;
    el motor de xlsxwriter escribe la fórmula real al exportar.
    """
    df_consolidado = (
        regalias_proyectos
        .join(bpines_version_anterior, on="BPIN", how="left")
        .join(_fecha_final,            on="BPIN", how="left")
        .join(regalias_contratos,      on="BPIN", how="left")
        .join(regalias_cargue,         on="BPIN", how="left")
    )

    return df_consolidado.select(
        "BPIN",
        "ENTIDAD O SECRETARIA",
        "NOMBRE PROYECTO",
        "ALCANCE DEL PROYECTO",
        "SECTOR",
        "INDICADOR DE PRODUCTO MGA",
        "ESTADO PROYECTO",
        "ESTADO CONTRATO",
        "TIPO CONTRATO",
        "FUENTE DE FINANCIACIÓN",
        "VALOR SGR",
        "VALOR NACIÓN",
        "VALOR OTROS",
        "VALOR OTRAS FUENTES NO SUIFP",
        "VALOR TOTAL PROYECTO",
        "VALOR PAGOS",
        "ULTIMA FECHA PAGO",
        "FECHA DE MIGRACIÓN A GESPROY",
        "FECHA DE ASIGNACIÓN DE RECURSOS",
        "FECHA DE INCORPORACIÓN DE RECURSOS",
        "AVANCE FISICO",
        "AVANCE FINANCIERO",
        "CPI",
        "SPI",
        # Fechas de contrato consolidadas (Gesproy > manual)
        "FECHA APROBACIÓN PROYECTO",
        "FECHA DE APERTURA DEL PRIMER PROCESO",
        "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL",
        "FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO",
        "FECHA ACTA INICIO",
        "HORIZONTE DEL PROYECTO",
        "FECHA DE FINALIZACIÓN",
        # Fecha que migra de la versión anterior (estado para cierre)
        pl.coalesce([pl.col("FECHA EN LA QUE PASO A ESTADO PARA CIERRE"), pl.lit(None).cast(pl.Date)]).alias("FECHA EN LA QUE PASO A ESTADO PARA CIERRE"),
        pl.coalesce([pl.col("FECHA DE CORTE GESPROY"), fecha_corte]).alias("FECHA DE CORTE GESPROY"),
        "INFORMACIÓN SOLICITADA",
        "INFORMACIÓN RECIBIDA",
        "FECHA DE RECIBO DE INFORMACIÓN",
        pl.lit("").alias("DESEMPEÑO EN EL CRONOGRAMA"),
        pl.lit("").alias("DESEMPEÑO EN EL COSTO"),
        pl.lit("").alias("COLUMNA APOYO"),
        pl.lit("").alias("BRECHA FISICO - FINANCIERA"),
        "CONTROL EXTERNALIDADES",
        # Ahora migra de la versión anterior (ya no se calcula en consolidación)
        pl.coalesce([pl.col("CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN"), pl.lit(None).cast(pl.Float64)]).alias("CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN"),
        pl.lit("").alias("CALIFICACIÓN INFORMACIÓN A TIEMPO"),
        pl.coalesce([pl.col("CALIFICACIÓN CALIDAD INFORMACIÓN"), pl.lit(None).cast(pl.Float64)]).alias("CALIFICACIÓN CALIDAD INFORMACIÓN"),
        pl.lit("").alias("COLUMNA APOYO 2"),
        pl.lit("").alias("CALIFICACIÓN EJECUCIÓN DEL PROYECTO"),
        pl.coalesce([pl.col("COMENTARIOS CALIFICACIÓN"), pl.lit("")]).alias("COMENTARIOS CALIFICACIÓN"),
        # Columnas manuales nuevas: si no existen, llegan como cadena vacía
        # gracias a `_completar_columnas_opcionales`.
        pl.coalesce([pl.col("RESPONSABLE CARGUE EN GESPROY"), pl.lit("")]).alias("RESPONSABLE CARGUE EN GESPROY"),
        pl.coalesce([pl.col("MUNICIPIOS"), pl.lit("")]).alias("MUNICIPIOS"),
    )


def consolidar_desc(otros_desc: pl.DataFrame, fecha_corte: pl.Expr) -> pl.DataFrame:
    """
    Aplica el coalesce de FECHA DE CORTE GESPROY contra la fecha calculada,
    y crea como placeholders (cadena vacía) las columnas de fórmula que no
    existan en el archivo. Las fórmulas Excel reales se escriben después
    en `escribir_hoja`, por lo que aquí solo importa que la columna exista.
    """
    cols_formula_placeholders = [
        "DESEMPEÑO EN EL CRONOGRAMA",
        "DESEMPEÑO EN EL COSTO",
        "COLUMNA APOYO",
        "BRECHA FISICO - FINANCIERA",
        "COLUMNA APOYO 2",
        "CALIFICACIÓN EJECUCIÓN DEL PROYECTO",
    ]
    faltantes = [
        pl.lit("").alias(c) for c in cols_formula_placeholders if c not in otros_desc.columns
    ]
    if faltantes:
        otros_desc = otros_desc.with_columns(faltantes)

    return otros_desc.with_columns(
        pl.coalesce([pl.col("FECHA DE CORTE GESPROY"), fecha_corte]).alias("FECHA DE CORTE GESPROY")
    )


# ══════════════════════════════════════════════════════════════════════════════
# ── Pipeline completo ─────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def consolidar(
    df_proy_raw: pl.DataFrame,
    df_cttos_raw: pl.DataFrame,
    df_carga_raw: pl.DataFrame,
    df_h1_raw: pl.DataFrame,
    df_desc_raw: pl.DataFrame,
    df_mun_raw: pl.DataFrame,
    fecha_corte: pl.Expr | None = None,
) -> tuple[pl.DataFrame, pl.DataFrame, pl.DataFrame, list[dict]]:
    """
    Pipeline completo de consolidación. Devuelve:

        bpines_version_anterior : DataFrame para la hoja H1 (con cálculos aplicados)
        otros_desc              : DataFrame para Descentralizadas (con cálculos)
        otros_mun               : DataFrame para Municipios
        fechas_conservadas      : Lista de dicts (BPIN, columna, fecha) que se
                                  conservaron desde la versión anterior porque
                                  Gesproy no las traía.
    """
    if fecha_corte is None:
        fecha_corte = pl.date(datetime.now().year, datetime.now().month, 15)

    # ── Fuentes Gesproy ──────────────────────────────────────────────────────
    regalias_proyectos = preparar_proyectos(df_proy_raw)
    regalias_contratos = preparar_contratos(df_cttos_raw)
    regalias_cargue    = preparar_cargue(df_carga_raw)

    # ── Versión anterior ─────────────────────────────────────────────────────
    bpines_version_anterior, otros_desc, otros_mun, fechas_manuales_h1 = (
        preparar_version_anterior(df_h1_raw, df_desc_raw, df_mun_raw)
    )

    # ── Fechas consolidadas (Gesproy > manual) ───────────────────────────────
    _fecha_final, fechas_conservadas = construir_fecha_final(
        regalias_proyectos, regalias_contratos, regalias_cargue, fechas_manuales_h1,
    )

    # ── Hoja 1 ───────────────────────────────────────────────────────────────
    bpines_version_anterior = consolidar_h1(
        regalias_proyectos, bpines_version_anterior, _fecha_final,
        regalias_contratos, regalias_cargue, fecha_corte,
    )

    # ── Hoja 2 ───────────────────────────────────────────────────────────────
    otros_desc = consolidar_desc(otros_desc, fecha_corte)

    # ── Cálculos en Python (días + calificación contratación) ────────────────
    bpines_version_anterior = agregar_calculos(bpines_version_anterior)
    otros_desc              = agregar_calculos(otros_desc)

    return bpines_version_anterior, otros_desc, otros_mun, fechas_conservadas


__all__ = [
    "preparar_proyectos",
    "preparar_contratos",
    "preparar_cargue",
    "preparar_version_anterior",
    "construir_fecha_final",
    "consolidar_h1",
    "consolidar_desc",
    "consolidar",
]
