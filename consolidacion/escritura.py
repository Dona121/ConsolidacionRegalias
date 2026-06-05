"""
Escritura de las hojas Excel con xlsxwriter.

`escribir_hoja` escribe una hoja con sus secciones, encabezados de colores,
datos y fórmulas. Todo lo relativo al diseño visual (paleta, formatos y
constantes de layout) vive en `diseno.py`; aquí solo se aplica.

Notas técnicas importantes (no cambiar sin entender):
- No usar `header_format` dentro de `add_table()` — corrompe styles.xml.
  Los colores de encabezado se aplican manualmente con `ws.write()` después.
- `style=None` en `add_table()` para que la tabla no aplique zebra ni colores.
- Columnas auxiliares COLUMNA APOYO / COLUMNA APOYO 2 se ocultan vía
  `set_column(..., {"hidden": True})`.
"""

from __future__ import annotations

import datetime as _dt

import pandas
import xlsxwriter

from .diseno import (
    ALTO_FILA, ALTO_FILA_SECCION, ANCHO_COLUMNA,
    FILA_SECCION, FILA_HEADER, FILA_DATOS,
    formato_encabezado,
)


# ══════════════════════════════════════════════════════════════════════════════
# ── Helpers para celdas (independientes del tipo pandas/python) ───────────────
# ══════════════════════════════════════════════════════════════════════════════

def _es_nulo(valor) -> bool:
    """True si el valor es None, NaN, NaT, o cadena vacía."""
    if valor is None:
        return True
    try:
        return bool(pandas.isnull(valor))
    except (TypeError, ValueError):
        return False


def _a_datetime(valor):
    """
    Convierte a `datetime.datetime` para xlsxwriter `write_datetime`.

    Acepta:
        - pandas.Timestamp        -> .to_pydatetime()
        - datetime.datetime       -> tal cual
        - datetime.date           -> datetime(year, month, day)
        - str ISO 'YYYY-MM-DD'    -> parse
    """
    if hasattr(valor, "to_pydatetime"):
        return valor.to_pydatetime()
    if isinstance(valor, _dt.datetime):
        return valor
    if isinstance(valor, _dt.date):
        return _dt.datetime(valor.year, valor.month, valor.day)
    if isinstance(valor, str):
        return _dt.datetime.fromisoformat(valor[:19])
    return valor


# ══════════════════════════════════════════════════════════════════════════════
# ── Escritura de una hoja ─────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def escribir_hoja(
    workbook: xlsxwriter.Workbook,
    ws_hoja,
    nombre_tabla: str,
    datos_hoja: pandas.DataFrame,
    secciones_hoja: list,        # [(titulo, [cols], fmt_titulo), ...]
    color_col: dict,             # col -> (color_celda, color_encabezado)
    col_fecha: set,
    col_numero: set,
    col_dias: set,
    col_con_formula: set,
    fn_formulas,                 # callable(r: int) -> dict | None
    col_ocultar: list,
    formatos: dict,              # resultado de crear_formatos()
) -> None:
    """
    Escribe la hoja Excel con sus secciones, encabezados de color, datos y
    fórmulas. `formatos` debe venir de `crear_formatos()`.
    """
    fmt_celda  = formatos["fmt_celda"]
    fmt_fecha  = formatos["fmt_fecha"]
    fmt_numero = formatos["fmt_numero"]
    fmt_dias   = formatos["fmt_dias"]

    todas  = [c for _, cols, _ in secciones_hoja for c in cols]
    n_cols = len(todas)
    n_rows = len(datos_hoja)

    # Tamaños fijos
    for ci in range(n_cols):
        ws_hoja.set_column(ci, ci, ANCHO_COLUMNA)
    ws_hoja.set_row(FILA_SECCION, ALTO_FILA_SECCION)
    ws_hoja.set_row(FILA_HEADER, ALTO_FILA)
    for ri in range(n_rows):
        ws_hoja.set_row(FILA_DATOS + ri, ALTO_FILA)

    # Títulos de sección con merge
    col_off = 0
    for titulo, cols, ft in secciones_hoja:
        n = len(cols)
        ws_hoja.merge_range(FILA_SECCION, col_off, FILA_SECCION, col_off + n - 1, titulo, ft)
        col_off += n

    # Tabla xlsxwriter — sin header_format para evitar el bug de styles.xml
    ws_hoja.add_table(
        FILA_HEADER, 0, FILA_HEADER + n_rows, n_cols - 1,
        {
            "name": nombre_tabla,
            "style": None,
            "autofilter": True,
            "header_row": True,
            "columns": [{"header": c} for c in todas],
        },
    )

    # Color de encabezados aplicado manualmente
    for ci, col in enumerate(todas):
        fmt_hdr = formato_encabezado(color_col[col][0], formatos)
        ws_hoja.write(FILA_HEADER, ci, col, fmt_hdr)

    # Columnas auxiliares ocultas
    for col in col_ocultar:
        if col in todas:
            ci = todas.index(col)
            ws_hoja.set_column(ci, ci, ANCHO_COLUMNA, None, {"hidden": True})

    # Datos fila a fila
    for row_idx in range(n_rows):
        excel_row = row_idx + FILA_DATOS + 1  # 1-based para referencias de fórmulas
        formulas_fila = fn_formulas(excel_row) if fn_formulas else {}

        for col_idx, col_name in enumerate(todas):
            er = row_idx + FILA_DATOS
            ec = col_idx

            if col_name in col_con_formula and col_name in formulas_fila:
                ws_hoja.write_formula(er, ec, formulas_fila[col_name], fmt_celda)
                continue

            valor = datos_hoja.iloc[row_idx, col_idx]

            if col_name in col_fecha:
                if _es_nulo(valor):
                    ws_hoja.write_blank(er, ec, None, fmt_fecha)
                else:
                    ws_hoja.write_datetime(er, ec, _a_datetime(valor), fmt_fecha)
            elif col_name in col_numero:
                if _es_nulo(valor):
                    ws_hoja.write_blank(er, ec, None, fmt_numero)
                else:
                    ws_hoja.write_number(er, ec, float(valor), fmt_numero)
            elif col_name in col_dias:
                if _es_nulo(valor):
                    ws_hoja.write_blank(er, ec, None, fmt_dias)
                else:
                    ws_hoja.write_number(er, ec, int(valor), fmt_dias)
            else:
                ws_hoja.write(er, ec, None if _es_nulo(valor) else valor, fmt_celda)


__all__ = [
    "escribir_hoja",
]
