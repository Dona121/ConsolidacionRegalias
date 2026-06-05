import streamlit as st
import polars as pl
import pandas
import xlsxwriter
import io
from datetime import datetime


st.set_page_config(
    page_title="Matriz de Seguimiento y Evaluación",
    layout="wide",
)

st.title("Matriz de Seguimiento y Evaluación — Regalías SGR")
st.markdown("Carga los archivos fuente para generar la matriz mensual.")


# ══════════════════════════════════════════════════════════════════════════════
# ── Esquemas esperados por archivo / tabla ────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════
#
# Cada entrada es:  columna → (tipo_legible, categoria)
# categoria indica qué validación de contenido aplicar:
#   "texto"  → debe ser String; si es numérico se reporta error (ej. BPIN)
#   "numero" → debe poder convertirse a Float64
#   "fecha"  → debe poder convertirse a Date
#   "libre"  → solo se valida presencia, sin validar contenido
# ──────────────────────────────────────────────────────────────────────────────

ESQUEMA_GESPROY_PROYECTOS = {
    "BPIN":                        ("Texto",  "texto"),
    "NOMBRE PROYECTO":             ("Texto",  "texto"),
    "SECTOR":                      ("Texto",  "texto"),
    "ESTADO PROYECTO":             ("Texto",  "texto"),
    "VALOR SGR":                   ("Número", "numero"),
    "VALOR NACIÓN":                ("Número", "numero"),
    "VALOR OTROS":                 ("Número", "numero"),
    "VALOR OTRAS FUENTES NO SUIFP":("Número", "numero"),
    "VALOR TOTAL PROYECTO":        ("Número", "numero"),
    "VALOR PAGOS":                 ("Número", "numero"),
}

ESQUEMA_GESPROY_CONTRATOS = {
    "BPIN":                   ("Texto",  "texto"),
    "FECHA ACT ADTIVO APERT": ("Fecha",  "fecha"),
    "ESTADO CONTRATO":        ("Texto",  "texto"),
    "FECHA INICIAL":          ("Fecha",  "fecha"),
    "FECHA SUSCRIPCION":      ("Fecha",  "fecha"),
    "ULTIMA FECHA PAGO":      ("Fecha",  "fecha"),
    "TIPO CONTRATO":          ("Texto",  "texto"),
    "VALOR TOTAL FUENTES SGR":("Número", "numero"),
}

ESQUEMA_GESPROY_CARGUE = {
    "BPIN":                      ("Texto",  "texto"),
    "FECHA APROBACIÓN PROYECTO": ("Fecha",  "fecha"),
    "AVANCE FISICO":             ("Número", "numero"),
    "AVANCE FINANCIERO":         ("Número", "numero"),
}

ESQUEMA_MATRIZ_H1 = {
    "BPIN":                               ("Texto",  "texto"),
    "ALCANCE DEL PROYECTO":               ("Texto",  "libre"),
    "FUENTE DE FINANCIACIÓN":             ("Texto",  "libre"),
    "ENTIDAD O SECRETARIA":               ("Texto",  "libre"),
    "INDICADOR DE PRODUCTO MGA":          ("Texto",  "libre"),
    "FECHA DE MIGRACIÓN A GESPROY":       ("Fecha",  "fecha"),
    "FECHA DE ASIGNACIÓN DE RECURSOS":    ("Fecha",  "fecha"),
    "FECHA DE INCORPORACIÓN DE RECURSOS": ("Fecha",  "fecha"),
    "HORIZONTE DEL PROYECTO":             ("Fecha",  "fecha"),
    "FECHA DE FINALIZACIÓN":              ("Fecha",  "fecha"),
    "CPI":                                ("Número", "numero"),
    "SPI":                                ("Número", "numero"),
    "INFORMACIÓN SOLICITADA":             ("Texto",  "libre"),
    "INFORMACIÓN RECIBIDA":               ("Texto",  "libre"),
    "FECHA DE RECIBO DE INFORMACIÓN":     ("Fecha",  "fecha"),
    "CONTROL EXTERNALIDADES":             ("Número", "numero"),
    "FECHA DE CORTE GESPROY":             ("Fecha",  "fecha"),
    "CALIFICACIÓN CALIDAD INFORMACIÓN":   ("Número", "numero"),
    "COMENTARIOS CALIFICACIÓN":           ("Texto",  "libre"),
}

ESQUEMA_MATRIZ_DESC = {
    "BPIN":                                        ("Texto",  "texto"),
    "EJECUTOR":                                    ("Texto",  "libre"),
    "NOMBRE DEL PROYECTO":                         ("Texto",  "libre"),
    "ALCANCE":                                     ("Texto",  "libre"),
    "SECTOR":                                      ("Texto",  "libre"),
    "FUENTE":                                      ("Texto",  "libre"),
    "ESTADO PROYECTO":                             ("Texto",  "texto"),
    "ESTADO CONTRATO":                             ("Texto",  "libre"),
    "VALOR SGR":                                   ("Número", "numero"),
    "VALOR OTROS ":                                ("Número", "numero"),
    "VALOR TOTAL":                                 ("Número", "numero"),
    "FECHA DE MIGRACIÓN A GESPROY":                ("Fecha",  "fecha"),
    "FECHA DE ASIGNACIÓN DE RECURSOS":             ("Fecha",  "fecha"),
    "FECHA DE INCORPORACIÓN DE RECUROS":           ("Fecha",  "fecha"),
    "AVANCE FÍSICO":                               ("Número", "numero"),
    "AVANCE FINANCIERO":                           ("Número", "numero"),
    "CPI":                                         ("Número", "numero"),
    "SPI":                                         ("Número", "numero"),
    "FECHA APROBACIÓN PROYECTO":                   ("Fecha",  "fecha"),
    "FECHA DE APERTURA DEL PRIMER PROCESO":        ("Fecha",  "fecha"),
    "FECHA SUSCRIPCION":                           ("Fecha",  "fecha"),
    "FECHA ACTA INICIO":                           ("Fecha",  "fecha"),
    "FECHA DE CORTE GESPROY":                      ("Fecha",  "fecha"),
    "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN":   ("Número", "numero"),
    "CALIFICACIÓN INFORMACIÓN A TIEMPO":           ("Número", "numero"),
    "CALIFICACIÓN CALIDAD INFORMACIÓN":            ("Número", "numero"),
    "CONTROL EXTERNALIDADES":                      ("Número", "numero"),
    "COMENTARIOS CALIFICACIÓN":                    ("Texto",  "libre"),
}

ESQUEMA_MATRIZ_MUN = {
    "BPIN":                              ("Texto",  "texto"),
    "EJECUTOR":                          ("Texto",  "libre"),
    "NOMBRE DEL PROYECTO":               ("Texto",  "libre"),
    "ALCANCE":                           ("Texto",  "libre"),
    "SECTOR":                            ("Texto",  "libre"),
    "FUENTE":                            ("Texto",  "libre"),
    "ESTADO PROYECTO":                   ("Texto",  "texto"),
    "ESTADO CONTRATO":                   ("Texto",  "libre"),
    "VALOR SGR":                         ("Número", "numero"),
    "VALOR OTROS":                       ("Número", "numero"),
    "VALOR TOTAL":                       ("Número", "numero"),
    "FECHA APROBACIÓN PROYECTO":         ("Fecha",  "fecha"),
    "FECHA DE ASIGNACIÓN DE RECURSOS":   ("Fecha",  "fecha"),
    "FECHA DE INCORPORACIÓN DE RECUROS": ("Fecha",  "fecha"),
    "FECHA ACTA INICIO":                 ("Fecha",  "fecha"),
    "AVANCE FÍSICO":                     ("Número", "numero"),
    "AVANCE FINANCIERO":                 ("Número", "numero"),
    "COMENTARIOS ":                      ("Texto",  "libre"),
}


# ══════════════════════════════════════════════════════════════════════════════
# ── Funciones de validación ───────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

_TIPOS_NUMERICOS = {
    pl.Int8, pl.Int16, pl.Int32, pl.Int64,
    pl.UInt8, pl.UInt16, pl.UInt32, pl.UInt64,
    pl.Float32, pl.Float64,
}

_FORMATOS_FECHA_VALIDOS = [
    "%d/%m/%Y",           # 01/03/2026  — formato de exportación actual
    "%Y-%m-%d",           # 2026-03-01  — formato legacy
    "%Y-%m-%d %H:%M:%S",  # 2026-03-01 00:00:00  — Datetime serializado como string
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


def _dtype_es_fecha(dtype) -> bool:
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


def validar_columnas(df: pl.DataFrame, esquema: dict, _nombre_ignorado: str) -> list[dict]:
    """
    Verifica presencia y tipo de cada columna del esquema.
    Retorna lista de dicts con claves: tipo, col, titulo, detalle, como_corregir.
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
                no_vacios   = serie.drop_nulls()
                no_vacios   = no_vacios.filter(no_vacios != "")
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
            elif _dtype_es_fecha(dtype):
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
    """Renderiza los problemas de validación en Streamlit, agrupados por archivo/tabla."""
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
    """Tabla de columnas esperadas y sus tipos."""
    filas = [{"Columna": col, "Tipo esperado": tipo} for col, (tipo, _) in esquema.items()]
    st.dataframe(pandas.DataFrame(filas), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# ── Lectura y normalización ───────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def ruta_reciente(archivos: list[str], prefijo: str) -> str | None:
    """Devuelve el nombre del archivo con el prefijo dado que tenga la fecha más reciente."""
    filtrados = [f for f in archivos if f.startswith(prefijo)]
    if not filtrados:
        return None
    filas = []
    for nombre in filtrados:
        fecha = pl.Series([nombre]).str.extract(r"(\d{8})").item()
        if fecha is not None:
            filas.append({"Nombre archivo": nombre, "Fecha": fecha})
    if not filas:
        # Ningún archivo tiene fecha en el nombre: devolver el primero encontrado
        return filtrados[0]
    return (
        pl.DataFrame(filas)
        .sort("Fecha", descending=True)
        .row(0, named=True)["Nombre archivo"]
    )


def leer_excel_regalias(contenido_bytes: bytes) -> pl.DataFrame:
    """
    Lee un reporte Gesproy donde la fila 0 es metadato, la fila 1 son los
    encabezados y los datos comienzan en la fila 2. Todas las columnas se
    leen como texto (infer_schema_length=0).
    """
    df = pl.read_excel(
        io.BytesIO(contenido_bytes),
        has_header=False,
        infer_schema_length=0,
    )
    if df.height < 2:
        raise ValueError(
            "El archivo no tiene suficientes filas. "
            "Se esperan al menos 2 (fila de metadato + fila de encabezados)."
        )
    encabezados = dict(zip(df.columns, df.row(1)))
    return (
        df
        .rename(encabezados)
        .slice(2)
        .select(pl.all().name.map(lambda x: x.strip()))
    )


def normalizar_fecha(df: pl.DataFrame, columnas: list[str]) -> pl.DataFrame:
    """
    Convierte columnas de fecha de forma robusta independientemente del dtype
    con que Polars las haya leído:

    - pl.Date     → sin cambio
    - pl.Datetime → cast a pl.Date (elimina la hora)
    - String      → intenta dd/mm/yyyy primero, luego yyyy-mm-dd
    - Otros tipos → sin cambio (el error ya fue reportado por validar_columnas)
    """
    exprs = []
    for col in columnas:
        if col not in df.columns:
            continue
        dtype = df.schema[col]
        if dtype == pl.Date:
            exprs.append(pl.col(col))
        elif _dtype_es_fecha(dtype):
            exprs.append(pl.col(col).cast(pl.Date, strict=False))
        elif dtype in _TIPOS_NUMERICOS:
            # Número de serie de Excel (días desde 1899-12-30)
            exprs.append(
                pl.when(pl.col(col).is_not_null())
                .then(
                    pl.col(col).cast(pl.Int32, strict=False)
                    .map_elements(
                        lambda n: __import__("datetime").date(1899, 12, 30)
                        + __import__("datetime").timedelta(days=int(n)),
                        return_dtype=pl.Date,
                    )
                )
                .otherwise(None)
                .alias(col)
            )
        elif dtype == pl.String:
            exprs.append(
                pl.when(pl.col(col).str.contains(r"^\d{2}/\d{2}/\d{4}$", literal=False))
                .then(pl.col(col).str.to_date("%d/%m/%Y", strict=False))
                .when(pl.col(col).str.contains(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$", literal=False))
                .then(pl.col(col).str.to_datetime("%Y-%m-%d %H:%M:%S", strict=False).cast(pl.Date))
                .otherwise(pl.col(col).str.to_date("%Y-%m-%d", strict=False))
                .alias(col)
            )
        # Otros tipos sin cambio
    if exprs:
        df = df.with_columns(exprs)
    return df


# ══════════════════════════════════════════════════════════════════════════════
# ── Funciones de cálculo ──────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def dias_desde_aprobacion_hasta_primer_proceso(
    estado_proyecto, fecha_aprobacion, fecha_corte_gesproy
):
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
    estado_proyecto, fecha_acta_inicio, fecha_primer_proceso
):
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
    estado_proyecto, fecha_corte_gesproy, fecha_suscripcion
):
    condicion = pl.col(estado_proyecto) == "CONTRATADO SIN ACTA DE INICIO"
    return (
        pl.when(condicion)
        .then((pl.col(fecha_corte_gesproy) - pl.col(fecha_suscripcion)).dt.total_days())
        .otherwise(None)
    )


def calificacion_desempeño_contratacion(
    estado_proyecto, fecha_aprobacion, fecha_apertura_primer_proceso,
    fecha_corte_gesproy, fecha_suscripcion, fecha_acta_inicio
):
    sin_contratar       = pl.col(estado_proyecto) == "SIN CONTRATAR"
    contratado_sin_acta = pl.col(estado_proyecto) == "CONTRATADO SIN ACTA DE INICIO"

    esc_2 = pl.col(fecha_aprobacion).is_not_null() & pl.col(fecha_apertura_primer_proceso).is_null()
    esc_3 = pl.col(fecha_suscripcion).is_null()    & pl.col(fecha_apertura_primer_proceso).is_not_null()
    esc_4 = pl.col(fecha_suscripcion).is_not_null()

    d_aprobacion  = (pl.col(fecha_corte_gesproy) - pl.col(fecha_aprobacion)).dt.total_days()
    d_apertura    = (pl.col(fecha_corte_gesproy) - pl.col(fecha_apertura_primer_proceso)).dt.total_days()
    d_suscripcion = (pl.col(fecha_corte_gesproy) - pl.col(fecha_suscripcion)).dt.total_days()

    cal_esc3 = (180 - d_apertura)   / (180 - 121)
    cal_esc4 = (60  - d_suscripcion) / 29 * 100

    return (
        pl.when(sin_contratar & esc_2 & (d_aprobacion  >= 0)  & (d_aprobacion  <= 180)).then(pl.lit(100))
        .when(sin_contratar & esc_2 & (d_aprobacion  > 180)).then(pl.lit(0))
        .when(sin_contratar & esc_3 & (d_apertura    >= 0)  & (d_apertura    <= 120)).then(pl.lit(100))
        .when(sin_contratar & esc_3 & (d_apertura    > 120) & (d_apertura    <= 180)).then(cal_esc3)
        .when(sin_contratar & esc_3 & (d_apertura    > 180)).then(pl.lit(0))
        .when(contratado_sin_acta & esc_4 & (d_suscripcion >= 0)  & (d_suscripcion <= 30)).then(pl.lit(100))
        .when(contratado_sin_acta & esc_4 & (d_suscripcion > 30)  & (d_suscripcion <= 60)).then(cal_esc4)
        .when(contratado_sin_acta & esc_4 & (d_suscripcion > 60)).then(pl.lit(0))
        .otherwise(pl.lit(None))
    )


# ══════════════════════════════════════════════════════════════════════════════
# ── Definición de columnas y colores ─────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

AZUL           = "#B7DEE8"
AZUL_HEADER    = "#31869B"
NARANJA        = "#F1A983"
NARANJA_HEADER = "#BE5014"

# ── Hoja 1 ────────────────────────────────────────────────────────────────────
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
    "FECHA SUSCRIPCION", "FECHA ACTA INICIO", "HORIZONTE DEL PROYECTO",
    "FECHA DE FINALIZACIÓN", "FECHA DE CORTE GESPROY",
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
    "COMENTARIOS CALIFICACIÓN",
]
todas_las_columnas = columnas_datos_generales + columnas_datos_calificacion + columnas_evaluacion

color_por_columna = (
    {c: (AZUL, AZUL_HEADER) for c in columnas_datos_generales}
    | {c: (NARANJA, NARANJA_HEADER) for c in columnas_datos_calificacion}
    | {c: (AZUL, AZUL_HEADER) for c in columnas_evaluacion}
)
columnas_fecha_h1 = {
    "ULTIMA FECHA PAGO", "FECHA APROBACIÓN PROYECTO",
    "FECHA DE APERTURA DEL PRIMER PROCESO", "FECHA SUSCRIPCION",
    "FECHA ACTA INICIO", "FECHA DE CORTE GESPROY",
    "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
    "FECHA DE INCORPORACIÓN DE RECURSOS",
    "HORIZONTE DEL PROYECTO", "FECHA DE FINALIZACIÓN",
    "FECHA DE RECIBO DE INFORMACIÓN",
}
columnas_numero_h1 = {
    "VALOR SGR", "VALOR NACIÓN", "VALOR OTROS", "VALOR OTRAS FUENTES NO SUIFP",
    "VALOR TOTAL PROYECTO", "VALOR PAGOS", "AVANCE FISICO", "AVANCE FINANCIERO",
    "CALIFICACIÓN CALIDAD INFORMACIÓN",
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

# ── Hoja 2 ────────────────────────────────────────────────────────────────────
cols_desc_generales = [
    "BPIN", "EJECUTOR", "NOMBRE DEL PROYECTO", "ALCANCE", "SECTOR", "FUENTE",
    "ESTADO PROYECTO", "ESTADO CONTRATO", "VALOR SGR", "VALOR OTROS ",
    "VALOR TOTAL", "FECHA DE MIGRACIÓN A GESPROY",
    "FECHA DE ASIGNACIÓN DE RECURSOS", "FECHA DE INCORPORACIÓN DE RECUROS",
]
cols_desc_calificacion = [
    "AVANCE FÍSICO", "AVANCE FINANCIERO",
    "CPI", "SPI",
    "FECHA APROBACIÓN PROYECTO", "FECHA DE APERTURA DEL PRIMER PROCESO",
    "FECHA SUSCRIPCION", "FECHA ACTA INICIO", "FECHA DE CORTE GESPROY",
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
    "COMENTARIOS CALIFICACIÓN",
]
todas_desc = cols_desc_generales + cols_desc_calificacion + cols_desc_evaluacion

color_desc = (
    {c: (AZUL, AZUL_HEADER) for c in cols_desc_generales}
    | {c: (NARANJA, NARANJA_HEADER) for c in cols_desc_calificacion}
    | {c: (AZUL, AZUL_HEADER) for c in cols_desc_evaluacion}
)
columnas_fecha_desc = {
    "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
    "FECHA DE INCORPORACIÓN DE RECUROS", "FECHA APROBACIÓN PROYECTO",
    "FECHA DE APERTURA DEL PRIMER PROCESO", "FECHA SUSCRIPCION",
    "FECHA ACTA INICIO", "FECHA DE CORTE GESPROY",
}
columnas_numero_desc = {
    "VALOR SGR", "VALOR OTROS ", "VALOR TOTAL",
    "AVANCE FÍSICO", "AVANCE FINANCIERO",
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
    # CALIFICACIÓN INFORMACIÓN A TIEMPO: es manual en esta hoja
}

# ── Hoja 3 ────────────────────────────────────────────────────────────────────
cols_mun = [
    "BPIN", "EJECUTOR", "NOMBRE DEL PROYECTO", "ALCANCE", "SECTOR", "FUENTE",
    "ESTADO PROYECTO", "ESTADO CONTRATO", "VALOR SGR", "VALOR OTROS",
    "VALOR TOTAL", "FECHA APROBACIÓN PROYECTO",
    "FECHA DE ASIGNACIÓN DE RECURSOS", "FECHA DE INCORPORACIÓN DE RECUROS",
    "FECHA ACTA INICIO", "AVANCE FÍSICO", "AVANCE FINANCIERO", "COMENTARIOS ",
]
color_mun          = {c: (AZUL, AZUL_HEADER) for c in cols_mun}
columnas_fecha_mun = {
    "FECHA APROBACIÓN PROYECTO", "FECHA DE ASIGNACIÓN DE RECURSOS",
    "FECHA DE INCORPORACIÓN DE RECUROS", "FECHA ACTA INICIO",
}
columnas_numero_mun = {
    "VALOR SGR", "VALOR OTROS", "VALOR TOTAL",
    "AVANCE FÍSICO", "AVANCE FINANCIERO",
}


# ══════════════════════════════════════════════════════════════════════════════
# ── Generadores de fórmulas Excel ────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def col_letter(i: int) -> str:
    """Convierte índice de columna (0-based) a letra(s) Excel (A, B, …, AA, …)."""
    result = ""
    while True:
        result = chr(i % 26 + 65) + result
        i = i // 26 - 1
        if i < 0:
            break
    return result


idx_h1   = {col: todas_las_columnas.index(col) for col in todas_las_columnas}
idx_desc = {col: todas_desc.index(col) for col in todas_desc}


def formulas_para_fila_h1(r: int) -> dict:
    def ref(c): return f"${col_letter(idx_h1[c])}{r}"

    estado       = ref("ESTADO PROYECTO")
    spi          = ref("SPI")
    cpi          = ref("CPI")
    af           = ref("AVANCE FINANCIERO")
    afis         = ref("AVANCE FISICO")
    col_apoyo    = ref("COLUMNA APOYO")
    col_apoyo2   = ref("COLUMNA APOYO 2")
    info_sol     = ref("INFORMACIÓN SOLICITADA")
    cron         = ref("DESEMPEÑO EN EL CRONOGRAMA")
    cost         = ref("DESEMPEÑO EN EL COSTO")
    brecha       = ref("BRECHA FISICO - FINANCIERA")
    ext          = ref("CONTROL EXTERNALIDADES")
    fecha_recibo = ref("FECHA DE RECIBO DE INFORMACIÓN")
    en_ejecucion = f'{estado}="CONTRATADO EN EJECUCI\u00d3N"'

    def f_indicador(v):
        p = f"(({v}-0.38)/(0.84-0.38))*100"
        return (
            f'=IF(AND({v}>1.3,{en_ejecucion}),0,'
            f'IF(AND({v}>1.25,{v}<=1.3,{en_ejecucion}),30,'
            f'IF(AND({v}>1.2,{v}<=1.25,{en_ejecucion}),90,'
            f'IF(AND({v}>=0.84,{v}<1.2,{en_ejecucion}),100,'
            f'IF(AND({v}>=0.38,{v}<0.84,{en_ejecucion}),{p},'
            f'IF(AND({v}<0.38,{en_ejecucion}),0,""))))))'
        )

    b = f"({af}-{afis})"
    f_apoyo = (
        f'=IF(AND({af}<60,{b}<=50),100,IF(AND({af}<60,{b}>50),0,'
        f'IF(AND({af}>=60,{af}<70,{b}<=40),100,IF(AND({af}>=60,{af}<70,{b}>40),0,'
        f'IF(AND({af}>=70,{af}<80,{b}<=30),100,IF(AND({af}>=70,{af}<80,{b}>30),0,'
        f'IF(AND({af}>=80,{af}<90,{b}<=20),100,IF(AND({af}>=80,{af}<90,{b}>20),0,'
        f'IF(AND({af}>=90,{af}<=100,{b}<=10),100,'
        f'IF(AND({af}>=90,{af}<=100,{b}>10),0,""))))))))))'
    )
    f_brecha = (
        f'=IF(AND({afis}>{af},{en_ejecucion}),100,'
        f'IF(AND({af}>{afis},{b}>50,{en_ejecucion}),0,'
        f'IF(AND({af}>{afis},{b}<=50,{af}<=60),100,'
        f'IF(AND({af}>{afis},{b}<=50,{af}>60),{col_apoyo},""))))'
    )
    f_info_tiempo = (
        f'=IF(ISNUMBER({fecha_recibo}),'
        f'IF(DAY({fecha_recibo})=10,100,'
        f'IF(DAY({fecha_recibo})=11,80,'
        f'IF(DAY({fecha_recibo})=12,50,0))),"")'
    )
    f_apoyo2 = (
        f'=IF({ext}=0,100,IF({ext}=1,90,IF({ext}=2,75,'
        f'IF({ext}=3,60,IF({ext}=4,50,IF({ext}=5,25,'
        f'IF({ext}>=6,0,"")))))))'
    )
    ponderado   = f"({cron}*0.4+{cost}*0.2+{brecha}*0.4)"
    calculo     = f'IFERROR(IF(ISTEXT({info_sol}),{ponderado},""),"")'
    f_ejecucion = (
        f'=IFERROR(IF(AND({en_ejecucion},{ext}>=1),'
        f'({calculo})*{col_apoyo2}/100,{calculo}),0)'
    )
    return {
        "DESEMPEÑO EN EL CRONOGRAMA":          f_indicador(spi),
        "DESEMPEÑO EN EL COSTO":               f_indicador(cpi),
        "COLUMNA APOYO":                       f_apoyo,
        "BRECHA FISICO - FINANCIERA":          f_brecha,
        "CALIFICACIÓN INFORMACIÓN A TIEMPO":   f_info_tiempo,
        "COLUMNA APOYO 2":                     f_apoyo2,
        "CALIFICACIÓN EJECUCIÓN DEL PROYECTO": f_ejecucion,
    }


def formulas_para_fila_desc(r: int) -> dict:
    def ref(c): return f"${col_letter(idx_desc[c])}{r}"

    estado       = ref("ESTADO PROYECTO")
    spi          = ref("SPI")
    cpi          = ref("CPI")
    af           = ref("AVANCE FINANCIERO")
    afis         = ref("AVANCE FÍSICO")
    col_apoyo    = ref("COLUMNA APOYO")
    col_apoyo2   = ref("COLUMNA APOYO 2")
    info_sol     = ref("CALIFICACIÓN INFORMACIÓN A TIEMPO")
    cron         = ref("DESEMPEÑO EN EL CRONOGRAMA")
    cost         = ref("DESEMPEÑO EN EL COSTO")
    brecha       = ref("BRECHA FISICO - FINANCIERA")
    ext          = ref("CONTROL EXTERNALIDADES")
    en_ejecucion = f'{estado}="CONTRATADO EN EJECUCI\u00d3N"'

    def f_indicador(v):
        p = f"(({v}-0.38)/(0.84-0.38))*100"
        return (
            f'=IF(AND({v}>1.3,{en_ejecucion}),0,'
            f'IF(AND({v}>1.25,{v}<=1.3,{en_ejecucion}),30,'
            f'IF(AND({v}>1.2,{v}<=1.25,{en_ejecucion}),90,'
            f'IF(AND({v}>=0.84,{v}<1.2,{en_ejecucion}),100,'
            f'IF(AND({v}>=0.38,{v}<0.84,{en_ejecucion}),{p},'
            f'IF(AND({v}<0.38,{en_ejecucion}),0,""))))))'
        )

    b = f"({af}-{afis})"
    f_apoyo = (
        f'=IF(AND({af}<60,{b}<=50),100,IF(AND({af}<60,{b}>50),0,'
        f'IF(AND({af}>=60,{af}<70,{b}<=40),100,IF(AND({af}>=60,{af}<70,{b}>40),0,'
        f'IF(AND({af}>=70,{af}<80,{b}<=30),100,IF(AND({af}>=70,{af}<80,{b}>30),0,'
        f'IF(AND({af}>=80,{af}<90,{b}<=20),100,IF(AND({af}>=80,{af}<90,{b}>20),0,'
        f'IF(AND({af}>=90,{af}<=100,{b}<=10),100,'
        f'IF(AND({af}>=90,{af}<=100,{b}>10),0,""))))))))))'
    )
    f_brecha = (
        f'=IF(AND({afis}>{af},{en_ejecucion}),100,'
        f'IF(AND({af}>{afis},{b}>50,{en_ejecucion}),0,'
        f'IF(AND({af}>{afis},{b}<=50,{af}<=60),100,'
        f'IF(AND({af}>{afis},{b}<=50,{af}>60),{col_apoyo},""))))'
    )
    f_apoyo2 = (
        f'=IF({ext}=0,100,IF({ext}=1,90,IF({ext}=2,75,'
        f'IF({ext}=3,60,IF({ext}=4,50,IF({ext}=5,25,'
        f'IF({ext}>=6,0,"")))))))'
    )
    ponderado   = f"({cron}*0.4+{cost}*0.2+{brecha}*0.4)"
    calculo     = f'IFERROR(IF(ISNUMBER({info_sol}),{ponderado},""),"")'
    f_ejecucion = (
        f'=IFERROR(IF(AND({en_ejecucion},{ext}>=1),'
        f'({calculo})*{col_apoyo2}/100,{calculo}),0)'
    )
    return {
        "DESEMPEÑO EN EL CRONOGRAMA":          f_indicador(spi),
        "DESEMPEÑO EN EL COSTO":               f_indicador(cpi),
        "COLUMNA APOYO":                       f_apoyo,
        "BRECHA FISICO - FINANCIERA":          f_brecha,
        "COLUMNA APOYO 2":                     f_apoyo2,
        "CALIFICACIÓN EJECUCIÓN DEL PROYECTO": f_ejecucion,
    }


# ══════════════════════════════════════════════════════════════════════════════
# ── Escritura de hojas Excel ──────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def escribir_hoja(
    workbook,
    ws_hoja,
    nombre_tabla: str,
    datos_hoja: pandas.DataFrame,
    secciones_hoja: list,
    color_col: dict,
    col_fecha: set,
    col_numero: set,
    col_dias: set,
    col_con_formula: set,
    fn_formulas,        # callable(r: int) -> dict | None
    col_ocultar: list,
    fmt_titulo_azul, fmt_titulo_naranja,
    fmt_header_azul, fmt_header_naranja,
    fmt_celda, fmt_fecha, fmt_numero, fmt_dias,
):
    todas        = [c for _, cols, _ in secciones_hoja for c in cols]
    n_cols       = len(todas)
    n_rows       = len(datos_hoja)
    ROW_HEIGHT   = 70
    COL_WIDTH    = 25
    FILA_SECCION = 0
    FILA_HEADER  = 1
    FILA_DATOS   = 2

    for ci in range(n_cols):
        ws_hoja.set_column(ci, ci, COL_WIDTH)
    ws_hoja.set_row(FILA_SECCION, 40)
    ws_hoja.set_row(FILA_HEADER, ROW_HEIGHT)
    for ri in range(n_rows):
        ws_hoja.set_row(FILA_DATOS + ri, ROW_HEIGHT)

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
        fmt_hdr = fmt_header_azul if color_col[col][0] == AZUL else fmt_header_naranja
        ws_hoja.write(FILA_HEADER, ci, col, fmt_hdr)

    # Columnas auxiliares ocultas
    for col in col_ocultar:
        if col in todas:
            ci = todas.index(col)
            ws_hoja.set_column(ci, ci, COL_WIDTH, None, {"hidden": True})

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
                if pandas.isnull(valor):
                    ws_hoja.write_blank(er, ec, None, fmt_fecha)
                else:
                    ws_hoja.write_datetime(er, ec, valor.to_pydatetime(), fmt_fecha)
            elif col_name in col_numero:
                if pandas.isnull(valor):
                    ws_hoja.write_blank(er, ec, None, fmt_numero)
                else:
                    ws_hoja.write_number(er, ec, float(valor), fmt_numero)
            elif col_name in col_dias:
                if pandas.isnull(valor):
                    ws_hoja.write_blank(er, ec, None, fmt_dias)
                else:
                    ws_hoja.write_number(er, ec, int(valor), fmt_dias)
            else:
                ws_hoja.write(er, ec, None if pandas.isnull(valor) else valor, fmt_celda)


# ══════════════════════════════════════════════════════════════════════════════
# ── UI: Carga de archivos ─────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

st.header("Carga de archivos fuente")

col1, col2 = st.columns(2)
with col1:
    st.subheader("Reportes Gesproy")
    uploads_gesproy = st.file_uploader(
        "Sube los archivos CG-proy, CG-cttos y CG-carga",
        type=["xlsx", "xls"],
        accept_multiple_files=True,
        key="gesproy",
    )
with col2:
    st.subheader("Versión anterior de la Matriz")
    upload_version_anterior = st.file_uploader(
        "Sube el archivo de la versión anterior (.xlsx)",
        type=["xlsx", "xls"],
        accept_multiple_files=False,
        key="version_anterior",
    )

# ── Referencia de columnas esperadas ──────────────────────────────────────────
st.divider()
with st.expander("Referencia: columnas esperadas por archivo"):
    st.caption(
        "Consulta esta tabla para verificar que los archivos tengan exactamente "
        "los nombres de columna y tipos de dato requeridos."
    )
    t1, t2, t3, t4, t5, t6 = st.tabs([
        "CG-proy", "CG-cttos", "CG-carga",
        "Matriz — Hoja 1", "Matriz — Descentralizadas", "Matriz — Municipios",
    ])
    with t1: mostrar_esquema(ESQUEMA_GESPROY_PROYECTOS)
    with t2: mostrar_esquema(ESQUEMA_GESPROY_CONTRATOS)
    with t3: mostrar_esquema(ESQUEMA_GESPROY_CARGUE)
    with t4: mostrar_esquema(ESQUEMA_MATRIZ_H1)
    with t5: mostrar_esquema(ESQUEMA_MATRIZ_DESC)
    with t6: mostrar_esquema(ESQUEMA_MATRIZ_MUN)

with st.expander("Cómo funciona la consolidación"):
    st.markdown("""
### Visión general

El aplicativo toma tres reportes de Gesproy (proyectos, contratos y cargue) más el archivo
de la versión anterior de la Matriz, los cruza por BPIN y produce un nuevo Excel con las
tres hojas de siempre. El resultado refleja el estado actual de Gesproy, complementado con
la información que el equipo ha ingresado manualmente en la Matriz cuando Gesproy aún no
tiene esa información disponible.

---

### Fuentes de información y qué aporta cada una

| Archivo | Qué contiene | Qué aporta a la consolidación |
|---|---|---|
| **CG-proy** | Listado de todos los proyectos SGR | Base de proyectos activos. Se excluyen los estados *Cerrado* y *Desaprobado*. |
| **CG-cttos** | Contratos de cada proyecto | Estado del contrato, tipo, fechas (apertura, suscripción, acta de inicio, última fecha de pago) y valor. Solo aparecen proyectos que ya tienen contrato. |
| **CG-carga** | Avance físico y financiero | Porcentajes de avance y fecha de aprobación del proyecto. |
| **Versión anterior** | Matriz del mes pasado | Información de contexto del proyecto (alcance, fuente, indicador MGA), datos ingresados manualmente (CPI, SPI, información solicitada/recibida, control externalidades, calificaciones) y fechas ingresadas cuando Gesproy no las tenía todavía. |

---

### Cómo se cruzan los datos (joins)

El cruce se hace siempre por **BPIN**, que es el código único de cada proyecto.

1. **Base:** se parte del listado de proyectos activos de CG-proy.
2. Se le agrega la información de contexto de la versión anterior (alcance, entidad, indicador MGA, etc.).
3. Se le agrega el contrato más representativo de cada BPIN (ver priorización abajo).
4. Se le agrega el avance físico, financiero y fecha de aprobación del cargue.

Si un proyecto no tiene contrato en Gesproy, las columnas de contrato quedan en blanco —
pero si el usuario tenía fechas ingresadas manualmente, esas se conservan (ver más abajo).

---

### Priorización de contratos

Cuando un proyecto tiene más de un contrato en Gesproy, el aplicativo selecciona uno solo
siguiendo este orden de prioridad:

1. Obra pública, Consultoría, Convenios de Cooperación, Interadministrativos
2. Suministro, Contratos con entidades sin ánimo de lucro
3. Cualquier otro tipo (el de mayor valor)

Dentro de cada categoría se toma el de **mayor valor SGR**.

---

### Priorización de fechas: Gesproy vs. fechas manuales

Esta es la regla más importante para entender cómo se manejan las fechas de contratación
(`FECHA APROBACIÓN PROYECTO`, `FECHA DE APERTURA DEL PRIMER PROCESO`, `FECHA SUSCRIPCION`,
`FECHA ACTA INICIO`, `ULTIMA FECHA PAGO`):

| Situación | Qué hace el aplicativo |
|---|---|
| El BPIN **no aparece** en el archivo de contratos (proyecto sin contratar) | Conserva la fecha ingresada manualmente en la Matriz. |
| El BPIN **sí aparece** en contratos, pero la fecha está **vacía** | Conserva la fecha ingresada manualmente en la Matriz. |
| El BPIN **sí aparece** en contratos **y la fecha está registrada** | Usa la fecha de Gesproy (tiene prioridad). |

Esto garantiza que las fechas que el equipo ingresó no se pierdan mientras el proyecto
avanza en Gesproy. Cuando Gesproy finalmente registre la fecha oficial, el aplicativo la
tomará automáticamente en la próxima consolidación.

Al finalizar la generación, el aplicativo muestra una tabla con todas las fechas que se
conservaron desde la versión anterior porque Gesproy aún no las tenía.

---

### Indicadores calculados automáticamente

El aplicativo calcula los siguientes campos para la hoja principal y la de descentralizadas:

**Días transcurridos** (solo para estados específicos):

- **Días desde la aprobación hasta apertura del primer proceso** — se calcula únicamente
  para proyectos en estado *Sin contratar*, como la diferencia entre la fecha de corte y
  la fecha de aprobación del proyecto.
- **Días desde la apertura hasta la firma del primer contrato** — se calcula únicamente
  para proyectos en estado *Sin contratar* que ya tienen fecha de apertura de proceso.
- **Días desde la suscripción hasta el acta de inicio** — se calcula únicamente para
  proyectos en estado *Contratado sin acta de inicio*.

**Calificación desempeño en la contratación** — evalúa qué tan oportuno fue el proceso de
contratación según el estado del proyecto:

- *Sin contratar, con aprobación pero sin apertura de proceso:* 100 puntos si han pasado
  180 días o menos desde la aprobación; 0 si superan los 180 días.
- *Sin contratar, con apertura de proceso pero sin suscripción:* 100 puntos si la apertura
  tiene 120 días o menos; escala proporcional entre 120 y 180 días; 0 si supera los 180.
- *Contratado sin acta de inicio:* 100 puntos si la suscripción tiene 30 días o menos;
  escala proporcional entre 30 y 60 días; 0 si supera los 60 días.

**Fórmulas en Excel** (calculadas dentro del archivo, no en Python):

- *Desempeño en el cronograma (SPI)* y *Desempeño en el costo (CPI):* escala basada en el
  valor del indicador. Por encima de 1.3 se penaliza; entre 0.84 y 1.2 es la zona óptima;
  por debajo de 0.38 es 0.
- *Brecha físico-financiera:* compara avance físico y avance financiero, considera el
  estado del proyecto y aplica rangos de tolerancia según el nivel de avance.
- *Calificación ejecución del proyecto:* pondera cronograma (40%), costo (20%) y brecha
  físico-financiera (40%). Si hay externalidades registradas, aplica un factor de ajuste.
- *Calificación información a tiempo* (solo hoja principal): evalúa el día en que se
  recibió la información solicitada (día 10 = 100 pts, día 11 = 80, día 12 = 50, otro = 0).

---

### Manejo de la fecha de corte

La fecha de corte se establece automáticamente como el **día 15 del mes actual**. Si la
versión anterior ya tiene una fecha de corte registrada para un proyecto, esa se conserva;
si no, se usa la calculada.

---

### Restricciones aplicadas

- Solo se incluyen proyectos en estado **distinto** a *Cerrado* y *Desaprobado*.
- Las columnas de fórmulas (desempeño, brecha, calificaciones) se **recalculan siempre**
  al abrir el Excel — no se guardan valores fijos.
- Las columnas auxiliares *Columna Apoyo* y *Columna Apoyo 2* están **ocultas** en el
  Excel; son intermedias del cálculo de la brecha y la calificación de ejecución.
- La hoja de **Municipios** no tiene fórmulas automáticas; toda su calificación es manual.
- La hoja de **Descentralizadas** tiene fórmulas para desempeño y ejecución, pero la
  *Calificación información a tiempo* es manual (la columna existe pero no tiene fórmula).
""")


st.divider()
st.header("Generar Matriz")

if st.button("Generar Matriz", type="primary", use_container_width=True):

    # ── Validación de archivos presentes ──────────────────────────────────────
    errores_archivos = []
    nombre_proy = nombre_cttos = nombre_carga = None

    if not uploads_gesproy:
        errores_archivos.append(
            "Faltan los archivos de Gesproy (CG-proy, CG-cttos y CG-carga)."
        )
    else:
        nombres_gesproy = [f.name for f in uploads_gesproy]
        nombre_proy  = ruta_reciente(nombres_gesproy, "CG-proy")
        nombre_cttos = ruta_reciente(nombres_gesproy, "CG-cttos")
        nombre_carga = ruta_reciente(nombres_gesproy, "CG-carga")
        if not nombre_proy:
            errores_archivos.append(
                "No se encontró el archivo **CG-proy**. "
                "El nombre debe comenzar con ese prefijo, p.ej. `CG-proy_20260301.xlsx`."
            )
        if not nombre_cttos:
            errores_archivos.append(
                "No se encontró el archivo **CG-cttos**. "
                "El nombre debe comenzar con ese prefijo, p.ej. `CG-cttos_20260301.xlsx`."
            )
        if not nombre_carga:
            errores_archivos.append(
                "No se encontró el archivo **CG-carga**. "
                "El nombre debe comenzar con ese prefijo, p.ej. `CG-carga_20260301.xlsx`."
            )

    if not upload_version_anterior:
        errores_archivos.append(
            "Falta el archivo de la **versión anterior de la Matriz**."
        )

    if errores_archivos:
        st.error("No se puede generar la Matriz. Revisa los archivos cargados:")
        for e in errores_archivos:
            st.markdown(f"- {e}")
        st.stop()

    gesproy_bytes    = {f.name: f.read() for f in uploads_gesproy}
    version_anterior = upload_version_anterior.read()

    progress = st.progress(0, text="Iniciando...")

    try:
        # ══════════════════════════════════════════════════════════════════════
        # ── Lectura de todos los archivos ─────────────────────────────────────
        # ══════════════════════════════════════════════════════════════════════

        progress.progress(8, text="Leyendo proyectos...")
        try:
            df_proy_raw = leer_excel_regalias(gesproy_bytes[nombre_proy])
        except Exception as e:
            progress.empty()
            st.error(f"No se pudo leer el archivo **CG-proy** (`{nombre_proy}`): {e}")
            st.stop()

        progress.progress(18, text="Leyendo versión anterior...")
        try:
            df_h1_raw = pl.read_excel(
                io.BytesIO(version_anterior),
                table_name="MatrizSeguimientoEvaluacion",
            )
        except Exception as e:
            progress.empty()
            st.error(
                "No se pudo leer la tabla **MatrizSeguimientoEvaluacion** "
                "del archivo de versión anterior. "
                "Verifica que el archivo tenga una tabla (no solo una hoja) "
                f"con ese nombre exacto. Detalle: {e}"
            )
            st.stop()

        try:
            df_desc_raw = pl.read_excel(
                io.BytesIO(version_anterior),
                table_name="OtrosEjecutoresDescentralizadas",
            )
        except Exception as e:
            progress.empty()
            st.error(
                "No se pudo leer la tabla **OtrosEjecutoresDescentralizadas**. "
                f"Detalle: {e}"
            )
            st.stop()

        try:
            df_mun_raw = pl.read_excel(
                io.BytesIO(version_anterior),
                table_name="OtrosEjecutoresMunicipios",
            )
        except Exception as e:
            progress.empty()
            st.error(
                "No se pudo leer la tabla **OtrosEjecutoresMunicipios**. "
                f"Detalle: {e}"
            )
            st.stop()

        progress.progress(30, text="Leyendo contratos...")
        try:
            df_cttos_raw = leer_excel_regalias(gesproy_bytes[nombre_cttos])
        except Exception as e:
            progress.empty()
            st.error(f"No se pudo leer el archivo **CG-cttos** (`{nombre_cttos}`): {e}")
            st.stop()

        progress.progress(40, text="Leyendo cargue...")
        try:
            df_carga_raw = leer_excel_regalias(gesproy_bytes[nombre_carga])
        except Exception as e:
            progress.empty()
            st.error(f"No se pudo leer el archivo **CG-carga** (`{nombre_carga}`): {e}")
            st.stop()

        # ── Validación de columnas y tipos ─────────────────────────────────────
        todos_los_problemas = [
            (f"CG-proy ({nombre_proy})",
             validar_columnas(df_proy_raw,  ESQUEMA_GESPROY_PROYECTOS, "")),
            (f"CG-cttos ({nombre_cttos})",
             validar_columnas(df_cttos_raw, ESQUEMA_GESPROY_CONTRATOS, "")),
            (f"CG-carga ({nombre_carga})",
             validar_columnas(df_carga_raw, ESQUEMA_GESPROY_CARGUE, "")),
            ("Versión anterior — MatrizSeguimientoEvaluacion",
             validar_columnas(df_h1_raw,   ESQUEMA_MATRIZ_H1, "")),
            ("Versión anterior — OtrosEjecutoresDescentralizadas",
             validar_columnas(df_desc_raw, ESQUEMA_MATRIZ_DESC, "")),
            ("Versión anterior — OtrosEjecutoresMunicipios",
             validar_columnas(df_mun_raw,  ESQUEMA_MATRIZ_MUN, "")),
        ]
        if sum(len(p) for _, p in todos_los_problemas) > 0:
            progress.empty()
            mostrar_errores_validacion(todos_los_problemas)
            st.stop()

        # ══════════════════════════════════════════════════════════════════════
        # ── Procesamiento ─────────────────────────────────────────────────────
        # ══════════════════════════════════════════════════════════════════════

        progress.progress(50, text="Procesando datos...")
        regalias_proyectos = (
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

        # Contratos: normalizar_fecha maneja String o Date sin distinción
        regalias_contratos = normalizar_fecha(
            df_cttos_raw.select(list(ESQUEMA_GESPROY_CONTRATOS.keys())),
            ["FECHA INICIAL", "FECHA ACT ADTIVO APERT", "FECHA SUSCRIPCION", "ULTIMA FECHA PAGO"],
        ).rename({
            "FECHA ACT ADTIVO APERT": "FECHA DE APERTURA DEL PRIMER PROCESO",
            "FECHA INICIAL":          "FECHA ACTA INICIO",
        })

        # Un contrato representativo por BPIN según prioridad de tipo
        tipos_prioritarios = {"Obra pública", "Consultoría", "Convenios de Cooperación", "Interadministrativos"}
        tipos_secundarios  = {"Suministro", "Contratos o convenios con entidades sin ánimo de lucro"}
        contratos = []
        for (bpin,), grupo in regalias_contratos.group_by("BPIN"):
            grupo = grupo.sort("VALOR TOTAL FUENTES SGR", descending=True)
            prioritarios = grupo.filter(pl.col("TIPO CONTRATO").is_in(tipos_prioritarios))
            if len(prioritarios) > 0:
                contratos.append(prioritarios.head(1))
                continue
            secundarios = grupo.filter(pl.col("TIPO CONTRATO").is_in(tipos_secundarios))
            contratos.append(secundarios.head(1) if len(secundarios) > 0 else grupo.head(1))
        if not contratos:
            regalias_contratos = regalias_contratos.clear()  # DataFrame vacío con mismo esquema
        else:
            regalias_contratos = pl.concat(contratos)

        # Cargue: normalizar_fecha maneja String o Date sin distinción
        regalias_cargue = normalizar_fecha(
            df_carga_raw
            .select(list(ESQUEMA_GESPROY_CARGUE.keys()))
            .with_columns(
                pl.col("AVANCE FISICO", "AVANCE FINANCIERO").cast(pl.Float64, strict=False)
            ),
            ["FECHA APROBACIÓN PROYECTO"],
        )

        # Versión anterior: normalizar fechas
        otros_ejecutores_descentralizadas = normalizar_fecha(
            df_desc_raw,
            [
                "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
                "FECHA DE INCORPORACIÓN DE RECUROS", "FECHA APROBACIÓN PROYECTO",
                "FECHA DE APERTURA DEL PRIMER PROCESO", "FECHA SUSCRIPCION", "FECHA ACTA INICIO",
                "FECHA DE CORTE GESPROY",
            ],
        )
        otros_ejecutores_municipios = normalizar_fecha(
            df_mun_raw,
            [
                "FECHA APROBACIÓN PROYECTO", "FECHA DE ASIGNACIÓN DE RECURSOS",
                "FECHA DE INCORPORACIÓN DE RECUROS", "FECHA ACTA INICIO",
            ],
        )
        # Columnas que pueden no existir en versiones anteriores del archivo:
        # se agregan vacías si no están presentes para garantizar compatibilidad.
        _cols_opcionales_h1 = {
            "CPI":                            pl.lit("").cast(pl.String),
            "SPI":                            pl.lit("").cast(pl.String),
            "INFORMACIÓN SOLICITADA":         pl.lit("").cast(pl.String),
            "INFORMACIÓN RECIBIDA":           pl.lit("").cast(pl.String),
            "FECHA DE RECIBO DE INFORMACIÓN": pl.lit(None).cast(pl.Date),
            "CONTROL EXTERNALIDADES":         pl.lit(None).cast(pl.Float64),
            "FECHA DE CORTE GESPROY":         pl.lit(None).cast(pl.Date),
            "CALIFICACIÓN CALIDAD INFORMACIÓN": pl.lit(None).cast(pl.Float64),
            "COMENTARIOS CALIFICACIÓN":       pl.lit("").cast(pl.String),
        }
        _cols_faltantes_h1 = {
            col: expr
            for col, expr in _cols_opcionales_h1.items()
            if col not in df_h1_raw.columns
        }
        if _cols_faltantes_h1:
            df_h1_raw = df_h1_raw.with_columns(
                [expr.alias(col) for col, expr in _cols_faltantes_h1.items()]
            )

        BPINes_version_anterior = normalizar_fecha(
            df_h1_raw.select(list(ESQUEMA_MATRIZ_H1.keys())),
            [
                "FECHA DE MIGRACIÓN A GESPROY", "FECHA DE ASIGNACIÓN DE RECURSOS",
                "FECHA DE INCORPORACIÓN DE RECURSOS", "FECHA DE FINALIZACIÓN",
                "HORIZONTE DEL PROYECTO", "FECHA DE RECIBO DE INFORMACIÓN",
                "FECHA DE CORTE GESPROY",
            ],
        )

        # ══════════════════════════════════════════════════════════════════════
        # ── Consolidación ─────────────────────────────────────────────────────
        # ══════════════════════════════════════════════════════════════════════
        progress.progress(60, text="Consolidando datos...")
        fecha_corte = pl.date(datetime.now().year, datetime.now().month, 15)

        # ── Leer fechas manuales de la versión anterior ──────────────────────
        # La tabla MatrizSeguimientoEvaluacion contiene fechas de contratación
        # ingresadas manualmente. Estas no forman parte de ESQUEMA_MATRIZ_H1
        # (que solo carga columnas de contexto del proyecto), pero sí existen en
        # el archivo. Las leemos aparte para poder hacer el coalesce correcto.
        _FECHAS_MANUALES_COLS = [
            "BPIN",
            "FECHA APROBACIÓN PROYECTO",
            "FECHA DE APERTURA DEL PRIMER PROCESO",
            "FECHA SUSCRIPCION",
            "FECHA ACTA INICIO",
            "ULTIMA FECHA PAGO",
        ]
        # Solo leer las que realmente existen en el archivo
        _cols_presentes = [c for c in _FECHAS_MANUALES_COLS if c in df_h1_raw.columns]
        _fechas_manuales_h1 = normalizar_fecha(
            df_h1_raw.select(_cols_presentes),
            [c for c in _FECHAS_MANUALES_COLS[1:] if c in _cols_presentes],
        )

        # ── Fuentes de Gesproy para cada fecha ────────────────────────────────
        _fuentes_gesproy = {
            "FECHA APROBACIÓN PROYECTO":            regalias_cargue.select("BPIN", "FECHA APROBACIÓN PROYECTO"),
            "FECHA DE APERTURA DEL PRIMER PROCESO": regalias_contratos.select("BPIN", "FECHA DE APERTURA DEL PRIMER PROCESO"),
            "FECHA SUSCRIPCION":                    regalias_contratos.select("BPIN", "FECHA SUSCRIPCION"),
            "FECHA ACTA INICIO":                    regalias_contratos.select("BPIN", "FECHA ACTA INICIO"),
            "ULTIMA FECHA PAGO":                    regalias_contratos.select("BPIN", "ULTIMA FECHA PAGO"),
        }

        # ── Construir tabla de fechas consolidadas (Gesproy > manual) ─────────
        # Base: TODOS los BPINs activos de Gesproy (incluyendo proyectos nuevos).
        # Para cada fecha: Gesproy tiene prioridad; si no tiene valor, se usa
        # la fecha ingresada manualmente en la versión anterior.
        # Proyectos nuevos (no están en la versión anterior) obtienen sus
        # fechas directamente de Gesproy.
        _fecha_final = regalias_proyectos.select("BPIN")
        filas_fechas_manuales = []
        _bpin_nombre = regalias_proyectos.select("BPIN", "NOMBRE PROYECTO")

        for col, df_gesproy in _fuentes_gesproy.items():
            # _fecha_final parte de todos los BPINs activos.
            # Para cada fecha, se hace:
            #   1. left join de todos BPINs con Gesproy  → fecha_gesproy (null si no existe)
            #   2. left join de todos BPINs con manuales → fecha_manual  (null si no existía en versión anterior)
            #   3. coalesce(gesproy, manual)              → prioriza Gesproy; si no, conserva manual
            #
            # Proyectos nuevos: manual=null → resultado = Gesproy (o null si tampoco está en Gesproy)
            # Proyectos sin contratar: gesproy=null, manual=valor → resultado = manual
            # Proyectos contratados: gesproy=valor → resultado = Gesproy

            todos_bpin = _fecha_final.select("BPIN")

            # Fecha de Gesproy para todos los BPINs
            con_gesproy = todos_bpin.join(
                df_gesproy.rename({col: "_gesproy"}),
                on="BPIN", how="left",
            )

            # Fecha manual para todos los BPINs (null si el BPIN no estaba en la versión anterior
            # o si la columna no existe en la versión anterior)
            if col in _fechas_manuales_h1.columns:
                con_manual = todos_bpin.join(
                    _fechas_manuales_h1.select("BPIN", pl.col(col).alias("_manual")),
                    on="BPIN", how="left",
                )
            else:
                con_manual = todos_bpin.with_columns(pl.lit(None).cast(pl.Date).alias("_manual"))

            # Combinar: coalesce(gesproy, manual)
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

            # Detectar casos donde Gesproy no tiene fecha pero el usuario sí ingresó una
            sin_gesproy_con_manual = combinado.filter(
                pl.col(f"_gesproy_{col}").is_null() &
                pl.col(f"_manual_{col}").is_not_null()
            )
            if len(sin_gesproy_con_manual) > 0:
                con_nombre = sin_gesproy_con_manual.join(_bpin_nombre, on="BPIN", how="left")
                for row in con_nombre.iter_rows(named=True):
                    filas_fechas_manuales.append({
                        "BPIN":             row["BPIN"],
                        "Nombre proyecto":  row.get("NOMBRE PROYECTO", ""),
                        "Columna":          col,
                        "Fecha conservada": str(row[f"_manual_{col}"]),
                    })

            _fecha_final = _fecha_final.join(
                combinado.select("BPIN", col),
                on="BPIN", how="left",
            )

        # ── Join principal ────────────────────────────────────────────────────
        df_consolidado = (
            regalias_proyectos
            .join(BPINes_version_anterior, on="BPIN", how="left")
            .join(_fecha_final,            on="BPIN", how="left")
            .join(regalias_contratos,      on="BPIN", how="left")
            .join(regalias_cargue,         on="BPIN", how="left")
        )

        # En df_consolidado, las fechas de contrato ya vienen de _fecha_final
        # (con la lógica Gesproy > manual aplicada). Los joins con contratos y
        # cargue agregan columnas adicionales (ESTADO CONTRATO, TIPO CONTRATO, etc.)
        # pero generan sufijos _right para las columnas de fecha que ya existen
        # en _fecha_final. Se descartan esos _right usando los nombres sin sufijo.
        BPINes_version_anterior = df_consolidado.select(
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
                # Fechas de contrato: vienen de _fecha_final (ya priorizadas)
                "FECHA APROBACIÓN PROYECTO",
                "FECHA DE APERTURA DEL PRIMER PROCESO",
                "FECHA SUSCRIPCION",
                "FECHA ACTA INICIO",
                "HORIZONTE DEL PROYECTO",
                "FECHA DE FINALIZACIÓN",
                pl.coalesce([pl.col("FECHA DE CORTE GESPROY"), fecha_corte]).alias("FECHA DE CORTE GESPROY"),
                "INFORMACIÓN SOLICITADA",
                "INFORMACIÓN RECIBIDA",
                "FECHA DE RECIBO DE INFORMACIÓN",
                pl.lit("").alias("DESEMPEÑO EN EL CRONOGRAMA"),
                pl.lit("").alias("DESEMPEÑO EN EL COSTO"),
                pl.lit("").alias("COLUMNA APOYO"),
                pl.lit("").alias("BRECHA FISICO - FINANCIERA"),
                "CONTROL EXTERNALIDADES",
                pl.lit("").alias("CALIFICACIÓN INFORMACIÓN A TIEMPO"),
                pl.coalesce([pl.col("CALIFICACIÓN CALIDAD INFORMACIÓN"), pl.lit(None).cast(pl.Float64)]).alias("CALIFICACIÓN CALIDAD INFORMACIÓN"),
                pl.lit("").alias("COLUMNA APOYO 2"),
                pl.lit("").alias("CALIFICACIÓN EJECUCIÓN DEL PROYECTO"),
                pl.coalesce([pl.col("COMENTARIOS CALIFICACIÓN"), pl.lit("")]).alias("COMENTARIOS CALIFICACIÓN"),
        )

        otros_ejecutores_descentralizadas = otros_ejecutores_descentralizadas.with_columns(
            pl.coalesce([pl.col("FECHA DE CORTE GESPROY"), fecha_corte]).alias("FECHA DE CORTE GESPROY")
        )

        # ══════════════════════════════════════════════════════════════════════
        # ── Cálculo de días y calificación de contratación ────────────────────
        # ══════════════════════════════════════════════════════════════════════
        progress.progress(70, text="Calculando indicadores...")

        def agregar_calculos(df: pl.DataFrame) -> pl.DataFrame:
            return df.with_columns(
                dias_desde_aprobacion_hasta_primer_proceso(
                    "ESTADO PROYECTO", "FECHA APROBACIÓN PROYECTO", "FECHA DE CORTE GESPROY"
                ).alias("DIAS DESDE LA APROBACIÓN HASTA APERTURA DEL PRIMER PROCESO"),
                dias_desde_apertura_hasta_primer_contrato(
                    "ESTADO PROYECTO", "FECHA ACTA INICIO", "FECHA DE APERTURA DEL PRIMER PROCESO"
                ).alias("DIAS DESDE LA APERTURA HASTA LA FIRMA DEL PRIMER CONTRATO"),
                dias_desde_suscripcion_hasta_fecha_acta_inicio(
                    "ESTADO PROYECTO", "FECHA DE CORTE GESPROY", "FECHA SUSCRIPCION"
                ).alias("DIAS DESDE LA FECHA DE SUSCRIPCIÓN HASTA LA FECHA DEL ACTA DE INICIO"),
                calificacion_desempeño_contratacion(
                    "ESTADO PROYECTO", "FECHA APROBACIÓN PROYECTO",
                    "FECHA DE APERTURA DEL PRIMER PROCESO", "FECHA DE CORTE GESPROY",
                    "FECHA SUSCRIPCION", "FECHA ACTA INICIO",
                ).alias("CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN"),
            )

        BPINes_version_anterior           = agregar_calculos(BPINes_version_anterior)
        otros_ejecutores_descentralizadas = agregar_calculos(otros_ejecutores_descentralizadas)

        # ══════════════════════════════════════════════════════════════════════
        # ── Vista previa ──────────────────────────────────────────────────────
        # ══════════════════════════════════════════════════════════════════════
        progress.progress(80, text="Preparando vista previa...")
        st.divider()
        st.header("Vista previa — Matriz principal")

        # ── Aviso de fechas manuales conservadas ──────────────────────────────
        if filas_fechas_manuales:
            st.warning(
                f"Se encontraron **{len(filas_fechas_manuales)} fecha(s)** que no están en los reportes "
                "de Gesproy y fueron conservadas desde la versión anterior. "
                "Revisa si estas fechas fueron ingresadas manualmente y si siguen siendo válidas."
            )
            st.dataframe(
                pandas.DataFrame(filas_fechas_manuales),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.success("Todas las fechas provienen de los reportes de Gesproy.")

        cols_preview = [
            "BPIN", "ENTIDAD O SECRETARIA", "NOMBRE PROYECTO", "ESTADO PROYECTO",
            "VALOR SGR", "AVANCE FISICO", "AVANCE FINANCIERO",
            "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN",
        ]
        st.caption(f"{len(BPINes_version_anterior)} proyectos")
        st.dataframe(
            BPINes_version_anterior
            .select([c for c in cols_preview if c in BPINes_version_anterior.columns])
            .to_pandas(),
            use_container_width=True,
            height=300,
        )

        # ══════════════════════════════════════════════════════════════════════
        # ── Preparar DataFrames para Excel ────────────────────────────────────
        # ══════════════════════════════════════════════════════════════════════
        progress.progress(86, text="Preparando datos para Excel...")
        datos_h1   = BPINes_version_anterior.select(todas_las_columnas).to_pandas()
        datos_desc = otros_ejecutores_descentralizadas.select(todas_desc).to_pandas()
        datos_mun  = otros_ejecutores_municipios.select(cols_mun).to_pandas()

        # ══════════════════════════════════════════════════════════════════════
        # ── Generar Excel ─────────────────────────────────────────────────────
        # ══════════════════════════════════════════════════════════════════════
        progress.progress(90, text="Generando Excel...")

        output_buffer = io.BytesIO()
        workbook      = xlsxwriter.Workbook(output_buffer, {"in_memory": True})

        _base_fmt = dict(font_name="Roboto", font_size=10, border=1, text_wrap=True)
        def _fmt(extra: dict): return workbook.add_format({**_base_fmt, **extra})

        fmt_titulo_azul    = _fmt({"bold": True, "font_color": "white", "bg_color": AZUL,           "align": "center", "valign": "vcenter", "font_size": 12})
        fmt_titulo_naranja = _fmt({"bold": True, "font_color": "white", "bg_color": NARANJA,         "align": "center", "valign": "vcenter", "font_size": 12})
        fmt_header_azul    = _fmt({"bold": True, "font_color": "white", "bg_color": AZUL_HEADER,    "align": "center", "valign": "vcenter"})
        fmt_header_naranja = _fmt({"bold": True, "font_color": "white", "bg_color": NARANJA_HEADER, "align": "center", "valign": "vcenter"})
        fmt_celda          = _fmt({"valign": "vcenter", "align": "center"})
        fmt_fecha_fmt      = _fmt({"valign": "vcenter", "align": "center", "num_format": "dd/mm/yyyy"})
        fmt_numero_fmt     = _fmt({"valign": "vcenter", "align": "center", "num_format": "#,##0.000"})
        fmt_dias_fmt       = _fmt({"valign": "vcenter", "align": "center"})

        fmt_kwargs = dict(
            fmt_titulo_azul=fmt_titulo_azul, fmt_titulo_naranja=fmt_titulo_naranja,
            fmt_header_azul=fmt_header_azul, fmt_header_naranja=fmt_header_naranja,
            fmt_celda=fmt_celda, fmt_fecha=fmt_fecha_fmt,
            fmt_numero=fmt_numero_fmt, fmt_dias=fmt_dias_fmt,
        )

        ws1 = workbook.add_worksheet("MatrizSeguimientoEvaluacion")
        escribir_hoja(
            workbook=workbook, ws_hoja=ws1,
            nombre_tabla="MatrizSeguimientoEvaluacion",
            datos_hoja=datos_h1,
            secciones_hoja=[
                ("DATOS GENERALES",         columnas_datos_generales,    fmt_titulo_azul),
                ("DATOS PARA CALIFICACIÓN", columnas_datos_calificacion, fmt_titulo_naranja),
                ("EVALUACIÓN",              columnas_evaluacion,         fmt_titulo_azul),
            ],
            color_col=color_por_columna,
            col_fecha=columnas_fecha_h1, col_numero=columnas_numero_h1,
            col_dias=columnas_dias_h1,   col_con_formula=columnas_con_formula_h1,
            fn_formulas=formulas_para_fila_h1,
            col_ocultar=["COLUMNA APOYO", "COLUMNA APOYO 2"],
            **fmt_kwargs,
        )

        ws2 = workbook.add_worksheet("OtrosEjecutoresDescentralizadas")
        escribir_hoja(
            workbook=workbook, ws_hoja=ws2,
            nombre_tabla="OtrosEjecutoresDescentralizadas",
            datos_hoja=datos_desc,
            secciones_hoja=[
                ("DATOS GENERALES",         cols_desc_generales,    fmt_titulo_azul),
                ("DATOS PARA CALIFICACIÓN", cols_desc_calificacion, fmt_titulo_naranja),
                ("EVALUACIÓN",              cols_desc_evaluacion,   fmt_titulo_azul),
            ],
            color_col=color_desc,
            col_fecha=columnas_fecha_desc, col_numero=columnas_numero_desc,
            col_dias=columnas_dias_desc,   col_con_formula=columnas_con_formula_desc,
            fn_formulas=formulas_para_fila_desc,
            col_ocultar=["COLUMNA APOYO", "COLUMNA APOYO 2"],
            **fmt_kwargs,
        )

        ws3 = workbook.add_worksheet("OtrosEjecutoresMunicipios")
        escribir_hoja(
            workbook=workbook, ws_hoja=ws3,
            nombre_tabla="OtrosEjecutoresMunicipios",
            datos_hoja=datos_mun,
            secciones_hoja=[("DATOS GENERALES", cols_mun, fmt_titulo_azul)],
            color_col=color_mun,
            col_fecha=columnas_fecha_mun, col_numero=columnas_numero_mun,
            col_dias=set(),              col_con_formula=set(),
            fn_formulas=None,            col_ocultar=[],
            **fmt_kwargs,
        )

        workbook.close()
        output_buffer.seek(0)
        progress.progress(100, text="Listo.")

        nombre_archivo = f"MatrizSeguimientoEvaluacion_{datetime.now():%Y%m%d_%H%M}.xlsx"
        st.success(f"Matriz generada con **{len(BPINes_version_anterior)} proyectos**.")
        st.divider()
        st.header("Descargar")
        st.download_button(
            label="Descargar Matriz Excel",
            data=output_buffer,
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

    except Exception as e:
        progress.empty()
        st.error(f"Error inesperado durante el procesamiento: {e}")
        st.exception(e)
