"""
Validación de columnas y tipos de dato para los archivos cargados.

`validar_columnas` recorre las columnas declaradas en un esquema y reporta
problemas concretos (columna faltante, tipo incorrecto, formato inválido).
La función auxiliar `mostrar_errores_validacion` los renderiza en Streamlit.
"""

from __future__ import annotations

import polars as pl
import pandas
import streamlit as st


# ══════════════════════════════════════════════════════════════════════════════
# ── Tipos y formatos aceptados ────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

_TIPOS_NUMERICOS = {
    pl.Int8, pl.Int16, pl.Int32, pl.Int64,
    pl.UInt8, pl.UInt16, pl.UInt32, pl.UInt64,
    pl.Float32, pl.Float64,
}

_FORMATOS_FECHA_VALIDOS = [
    "%d/%m/%Y",           # 01/03/2026  — formato de exportación actual
    "%Y-%m-%d",           # 2026-03-01  — formato legacy
    "%Y-%m-%d %H:%M:%S",  # 2026-03-01 00:00:00  — Datetime serializado
]

_COMO_CORREGIR = {
    "texto_es_numero": (
        "Esta columna debe contener **texto** (letras o combinación de letras y números), "
        "pero Excel la convirtió a **número**. "
        "Para corregirlo: selecciona la columna en Excel, clic derecho, "
        "*Formato de celdas*, elige **Texto** y vuelve a escribir o pegar los valores."
    ),
    "numero_es_texto": (
        "Esta columna debe contener **números**, pero tiene valores que no son numéricos. "
        "Revisa que no haya letras, espacios o símbolos de moneda (como $ o %) mezclados en las celdas."
    ),
    "fecha_es_numero": (
        "Esta columna debe contener **fechas**, pero Excel la guardó como **número** (número de serie). "
        "Selecciona la columna, abre *Formato de celdas* y elige **Fecha** con el formato **dd/mm/yyyy**."
    ),
    "fecha_formato_invalido": (
        "Esta columna debe contener **fechas** en formato dd/mm/yyyy (p.ej. 01/03/2026), "
        "pero los valores no corresponden a ese formato. "
        "Verifica que la columna esté en formato *Fecha* con el patrón **dd/mm/yyyy** en Excel."
    ),
    "columna_faltante": (
        "Abre el archivo en Excel y verifica que el encabezado de esa columna "
        "esté escrito exactamente como se indica aquí, sin espacios al inicio ni al final "
        "y con las mismas mayúsculas."
    ),
}


# ══════════════════════════════════════════════════════════════════════════════
# ── Helpers internos ──────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def dtype_es_fecha(dtype) -> bool:
    """True si el dtype es Date o cualquier variante de Datetime."""
    if dtype == pl.Date:
        return True
    try:
        return isinstance(dtype, pl.Datetime)
    except Exception:
        return hasattr(dtype, "base_type") and dtype.base_type() == pl.Datetime


def _es_fecha_valida_str(serie: pl.Series) -> bool:
    """True si todos los valores no-vacíos de la serie se pueden parsear como fecha."""
    no_vacios = serie.drop_nulls()
    no_vacios = no_vacios.filter(no_vacios != "")
    if len(no_vacios) == 0:
        return True
    for fmt in _FORMATOS_FECHA_VALIDOS:
        try:
            if "%S" in fmt:
                parseados = no_vacios.str.to_datetime(fmt, strict=False)
            else:
                parseados = no_vacios.str.to_date(fmt, strict=False)
            if parseados.drop_nulls().len() == len(no_vacios):
                return True
        except Exception:
            continue
    return False


def _ejemplos_valores(serie: pl.Series, n: int = 3) -> str:
    """Hasta n valores no-nulos de la serie, formateados como texto legible."""
    como_str  = serie.drop_nulls().cast(pl.String)
    no_vacios = como_str.filter(como_str != "").head(n).to_list()
    if not no_vacios:
        return "*(columna vacía)*"
    return ", ".join(f"`{v}`" for v in no_vacios)


# ══════════════════════════════════════════════════════════════════════════════
# ── API pública ───────────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def validar_columnas(df: pl.DataFrame, esquema: dict, _nombre_ignorado: str = "") -> list[dict]:
    """
    Verifica presencia y tipo de cada columna del esquema.

    Retorna una lista de dicts con las claves:
        tipo, col, titulo, detalle, como_corregir

    `_nombre_ignorado` se mantiene por compatibilidad con la firma anterior.
    """
    problemas = []
    columnas_df = set(df.columns)

    for col, (tipo_legible, categoria) in esquema.items():

        # ── Columna faltante ───────────────────────────────────────────────────
        if col not in columnas_df:
            problemas.append({
                "tipo": "faltante",
                "col": col,
                "titulo": f"Columna `{col}` no encontrada",
                "detalle": (
                    f"Se esperaba una columna llamada exactamente **`{col}`** "
                    f"(tipo: *{tipo_legible}*) pero no existe en el archivo. "
                    "Puede ser un error de nombre, mayúsculas o espacios extra."
                ),
                "como_corregir": _COMO_CORREGIR["columna_faltante"],
            })
            continue

        if categoria == "libre":
            continue

        serie = df[col]
        dtype = serie.dtype

        # ── Texto ──────────────────────────────────────────────────────────────
        if categoria == "texto":
            if dtype in _TIPOS_NUMERICOS:
                problemas.append({
                    "tipo": "tipo_incorrecto",
                    "col": col,
                    "titulo": f"Columna `{col}` está guardada como número (debe ser texto)",
                    "detalle": (
                        f"Valores encontrados: {_ejemplos_valores(serie)}. "
                        "El sistema necesita esta columna como texto porque puede contener "
                        "ceros al inicio u otros caracteres que Excel elimina al convertirla a número."
                    ),
                    "como_corregir": _COMO_CORREGIR["texto_es_numero"],
                })

        # ── Número ─────────────────────────────────────────────────────────────
        elif categoria == "numero":
            if dtype == pl.String:
                no_vacios = serie.drop_nulls()
                no_vacios = no_vacios.filter(no_vacios != "")
                if len(no_vacios) > 0:
                    convertidos = no_vacios.cast(pl.Float64, strict=False)
                    invalidos   = no_vacios.filter(convertidos.is_null())
                    if len(invalidos) > 0:
                        ejemplos = ", ".join(f"`{v}`" for v in invalidos.head(3).to_list())
                        problemas.append({
                            "tipo": "tipo_incorrecto",
                            "col": col,
                            "titulo": f"Columna `{col}` tiene valores que no son números",
                            "detalle": (
                                f"Se encontraron {len(invalidos)} valor(es) no numérico(s). "
                                f"Ejemplos: {ejemplos}."
                            ),
                            "como_corregir": _COMO_CORREGIR["numero_es_texto"],
                        })
            elif dtype_es_fecha(dtype):
                problemas.append({
                    "tipo": "tipo_incorrecto",
                    "col": col,
                    "titulo": f"Columna `{col}` está guardada como fecha (debe ser número)",
                    "detalle": "El sistema esperaba un número en esta columna pero encontró fechas.",
                    "como_corregir": _COMO_CORREGIR["numero_es_texto"],
                })

        # ── Fecha ──────────────────────────────────────────────────────────────
        elif categoria == "fecha":
            if dtype in _TIPOS_NUMERICOS:
                problemas.append({
                    "tipo": "tipo_incorrecto",
                    "col": col,
                    "titulo": f"Columna `{col}` está guardada como número (debe ser fecha)",
                    "detalle": f"Valores encontrados: {_ejemplos_valores(serie)}.",
                    "como_corregir": _COMO_CORREGIR["fecha_es_numero"],
                })
            elif dtype == pl.String:
                if not _es_fecha_valida_str(serie):
                    no_vacios = serie.drop_nulls()
                    no_vacios = no_vacios.filter(no_vacios != "")
                    ejemplos  = ", ".join(f"`{v}`" for v in no_vacios.head(3).to_list())
                    problemas.append({
                        "tipo": "tipo_incorrecto",
                        "col": col,
                        "titulo": f"Columna `{col}` tiene fechas en formato no reconocido",
                        "detalle": (
                            f"Ejemplos encontrados: {ejemplos}. "
                            "El sistema acepta los formatos **dd/mm/yyyy** o **yyyy-mm-dd**."
                        ),
                        "como_corregir": _COMO_CORREGIR["fecha_formato_invalido"],
                    })
            # pl.Date / pl.Datetime → correcto, sin problema

    return problemas


def mostrar_errores_validacion(
    problemas_por_fuente: list[tuple[str, list[dict]]],
) -> None:
    """Renderiza los problemas de validación en Streamlit, agrupados por fuente."""
    total = sum(len(p) for _, p in problemas_por_fuente)
    st.error(
        f"Se encontraron **{total} problema(s)** en los archivos cargados. "
        "Corrígelos en Excel y vuelve a cargar los archivos."
    )
    for nombre_fuente, problemas in problemas_por_fuente:
        if not problemas:
            continue
        st.markdown("---")
        st.markdown(f"### {nombre_fuente}")
        for p in problemas:
            prefijo = "Columna faltante" if p["tipo"] == "faltante" else "Tipo incorrecto"
            with st.expander(f"{prefijo}: {p['titulo']}", expanded=True):
                st.markdown(f"**Qué pasó:** {p['detalle']}")
                st.markdown(f"**Cómo corregirlo:** {p['como_corregir']}")
    st.info(
        "Despliega la sección **Referencia: columnas esperadas por archivo** (más arriba) "
        "para ver el listado completo de columnas requeridas con sus tipos."
    )


def mostrar_esquema(esquema: dict) -> None:
    """Tabla de columnas esperadas y sus tipos para la referencia visible al usuario."""
    filas = [{"Columna": col, "Tipo esperado": tipo} for col, (tipo, _) in esquema.items()]
    st.dataframe(pandas.DataFrame(filas), use_container_width=True, hide_index=True)


__all__ = [
    "dtype_es_fecha",
    "validar_columnas",
    "mostrar_errores_validacion",
    "mostrar_esquema",
]
