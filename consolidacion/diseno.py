"""
Diseño visual del Excel: paleta de colores, formatos de celda y constantes
de layout (alto/ancho, filas de sección/encabezado/datos).

Este módulo centraliza TODO lo relacionado con la apariencia del libro Excel,
de modo que `columnas.py` se ocupe solo de la estructura de datos (qué columna
va dónde y de qué color) y `escritura.py` solo de la mecánica de escritura.

Convención de colores
----------------------
Cada "tema" tiene dos tonos:

    <COLOR>         → tono claro: fondo de la barra de título de sección.
    <COLOR>_HEADER  → tono oscuro: fondo del encabezado de cada columna.

El mapeo `columna → (color_celda, color_encabezado)` vive en `columnas.py`;
aquí solo se definen los valores y los `Format` de xlsxwriter que los aplican.
"""

from __future__ import annotations

import xlsxwriter


# ══════════════════════════════════════════════════════════════════════════════
# ── Paleta de colores ─────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

AZUL           = "#B7DEE8"
AZUL_HEADER    = "#31869B"
NARANJA        = "#F1A983"
NARANJA_HEADER = "#BE5014"
VERDE          = "#D8E4BC"
VERDE_HEADER   = "#9BBB59"


# ══════════════════════════════════════════════════════════════════════════════
# ── Constantes de layout ──────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

ALTO_FILA          = 70   # alto de la fila de encabezado y de cada fila de datos
ALTO_FILA_SECCION  = 40   # alto de la fila de títulos de sección
ANCHO_COLUMNA      = 25   # ancho uniforme de todas las columnas

FILA_SECCION = 0          # fila 0: títulos de sección (merge)
FILA_HEADER  = 1          # fila 1: encabezados de columna
FILA_DATOS   = 2          # fila 2 en adelante: datos


# ══════════════════════════════════════════════════════════════════════════════
# ── Construcción de formatos compartidos ──────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def crear_formatos(workbook: xlsxwriter.Workbook) -> dict:
    """
    Crea todos los formatos necesarios y los devuelve en un diccionario.

    Devuelve un dict con:
        fmt_titulo_*   : formatos para la fila de título de sección
        fmt_header_*   : formatos para la fila de encabezados de columna
        fmt_celda      : formato genérico de celda
        fmt_fecha      : celda con número de formato dd/mm/yyyy
        fmt_numero     : celda con número de formato #,##0.000
        fmt_dias       : celda numérica para días (entero)
    """
    base = dict(font_name="Roboto", font_size=10, border=1, text_wrap=True)
    f    = lambda extra: workbook.add_format({**base, **extra})

    return {
        # Títulos de sección (texto blanco, fondo claro)
        "fmt_titulo_azul":    f({"bold": True, "font_color": "white", "bg_color": AZUL,    "align": "center", "valign": "vcenter", "font_size": 12}),
        "fmt_titulo_naranja": f({"bold": True, "font_color": "white", "bg_color": NARANJA, "align": "center", "valign": "vcenter", "font_size": 12}),
        "fmt_titulo_verde":   f({"bold": True, "font_color": "white", "bg_color": VERDE,   "align": "center", "valign": "vcenter", "font_size": 12}),

        # Encabezados de columna (texto blanco, fondo oscuro)
        "fmt_header_azul":    f({"bold": True, "font_color": "white", "bg_color": AZUL_HEADER,    "align": "center", "valign": "vcenter"}),
        "fmt_header_naranja": f({"bold": True, "font_color": "white", "bg_color": NARANJA_HEADER, "align": "center", "valign": "vcenter"}),
        "fmt_header_verde":   f({"bold": True, "font_color": "white", "bg_color": VERDE_HEADER,   "align": "center", "valign": "vcenter"}),

        # Celdas
        "fmt_celda":  f({"valign": "vcenter", "align": "center"}),
        "fmt_fecha":  f({"valign": "vcenter", "align": "center", "num_format": "dd/mm/yyyy"}),
        "fmt_numero": f({"valign": "vcenter", "align": "center", "num_format": "#,##0.000"}),
        "fmt_dias":   f({"valign": "vcenter", "align": "center"}),
    }


def formato_encabezado(color_celda: str, formatos: dict):
    """
    Devuelve el `Format` de encabezado que corresponde al color de celda de una
    columna. Si el color no está mapeado, usa el encabezado azul por defecto.

    `color_celda` es el primer elemento de la tupla `(color_celda, color_header)`
    definida en `columnas.py`.
    """
    header_por_color = {
        AZUL:    formatos["fmt_header_azul"],
        NARANJA: formatos["fmt_header_naranja"],
        VERDE:   formatos["fmt_header_verde"],
    }
    return header_por_color.get(color_celda, formatos["fmt_header_azul"])


__all__ = [
    # Paleta
    "AZUL", "AZUL_HEADER", "NARANJA", "NARANJA_HEADER", "VERDE", "VERDE_HEADER",
    # Layout
    "ALTO_FILA", "ALTO_FILA_SECCION", "ANCHO_COLUMNA",
    "FILA_SECCION", "FILA_HEADER", "FILA_DATOS",
    # Formatos
    "crear_formatos", "formato_encabezado",
]
