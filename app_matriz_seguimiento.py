"""
App Streamlit — Matriz de Seguimiento y Evaluación SGR.

Orquesta los módulos del paquete `consolidacion`:

    esquemas       → diccionarios de validación
    columnas       → orden de columnas, colores, sets de tipo
    validacion     → validar_columnas + helpers para la UI
    lectura        → leer_excel_regalias, normalizar_fecha, ruta_reciente
    calculos       → días y calificación contratación (en Python)
    formulas       → fórmulas Excel para H1 y Desc
    escritura      → crear_formatos + escribir_hoja
    procesamiento  → pipeline consolidar()
"""

from __future__ import annotations

import io
from datetime import datetime

import pandas
import polars as pl
import streamlit as st
import xlsxwriter

from consolidacion.esquemas import (
    ESQUEMA_GESPROY_PROYECTOS,
    ESQUEMA_GESPROY_CONTRATOS,
    ESQUEMA_GESPROY_CARGUE,
    ESQUEMA_MATRIZ_H1,
    ESQUEMA_MATRIZ_DESC,
    ESQUEMA_MATRIZ_MUN,
)
from consolidacion.columnas import (
    todas_las_columnas, columnas_datos_generales, columnas_datos_calificacion,
    columnas_evaluacion, color_por_columna, columnas_fecha_h1,
    columnas_numero_h1, columnas_dias_h1, columnas_con_formula_h1,
    todas_desc, cols_desc_generales, cols_desc_calificacion, cols_desc_evaluacion,
    color_desc, columnas_fecha_desc, columnas_numero_desc, columnas_dias_desc,
    columnas_con_formula_desc,
    cols_mun, color_mun, columnas_fecha_mun, columnas_numero_mun,
)
from consolidacion.validacion import (
    validar_columnas, mostrar_errores_validacion, mostrar_esquema,
)
from consolidacion.lectura import leer_excel_regalias, leer_tabla_excel, ruta_reciente
from consolidacion.formulas import formulas_para_fila_h1, formulas_para_fila_desc
from consolidacion.diseno import crear_formatos
from consolidacion.escritura import escribir_hoja
from consolidacion.procesamiento import consolidar


# ══════════════════════════════════════════════════════════════════════════════
# ── Configuración general ─────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="Matriz de Seguimiento y Evaluación",
    layout="wide",
)
st.title("Matriz de Seguimiento y Evaluación — Regalías SGR")
st.markdown("Carga los archivos fuente para generar la matriz mensual.")


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
(`FECHA APROBACIÓN PROYECTO`, `FECHA DE APERTURA DEL PRIMER PROCESO`,
`FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL`, `FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO`,
`FECHA ACTA INICIO`, `ULTIMA FECHA PAGO`):

En la Matriz principal, `FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO` es el mínimo de las fechas de
suscripción de todos los contratos del BPIN en CG-cttos (el contrato principal puede no ser el
primero suscrito). En Descentralizadas esta columna se conserva de la versión anterior.

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

**Calificación desempeño en la contratación** — ya **no se calcula** durante la
consolidación: ahora **migra de la versión anterior** de la Matriz, igual que las demás
columnas de contexto. El equipo la edita manualmente y el aplicativo la conserva entre
versiones.

**Fórmulas en Excel** (calculadas dentro del archivo, no en Python):

- *Desempeño en el cronograma (SPI)* y *Desempeño en el costo (CPI):* escala basada en el
  valor del indicador. Por encima de 1.3 se penaliza; entre 0.84 y 1.2 es la zona óptima;
  por debajo de 0.38 es 0.
- *Brecha físico-financiera:* compara avance físico y avance financiero, considera el
  estado del proyecto y aplica rangos de tolerancia según el nivel de avance.
- *Calificación ejecución del proyecto:* pondera cronograma (40%), costo (20%) y brecha
  físico-financiera (40%). Si hay externalidades registradas, aplica un factor de ajuste.
  Si el proyecto sigue *Contratado en ejecución* pero su **horizonte ya venció** respecto a
  la fecha de corte, la calificación es 0.
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
- Las columnas **RESPONSABLE CARGUE EN GESPROY** (hoja principal) y **HORIZONTE DEL PROYECTO**
  (Descentralizadas) son **manuales**: no migran de Gesproy y se conservan tal cual entre
  versiones de la Matriz.
- La columna **MUNICIPIOS** (en las tres hojas, al final, con encabezado verde) también es
  **manual**: no migra de Gesproy y se conserva desde la versión anterior de la Matriz.
""")


st.divider()
st.header("Generar Matriz")


# ══════════════════════════════════════════════════════════════════════════════
# ── Acción principal ──────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

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
        # ── Lectura ──────────────────────────────────────────────────────────
        progress.progress(8, text="Leyendo proyectos...")
        try:
            df_proy_raw = leer_excel_regalias(gesproy_bytes[nombre_proy])
        except Exception as e:
            progress.empty()
            st.error(f"No se pudo leer el archivo **CG-proy** (`{nombre_proy}`): {e}")
            st.stop()

        progress.progress(18, text="Leyendo versión anterior...")
        try:
            df_h1_raw = leer_tabla_excel(version_anterior, "MatrizSeguimientoEvaluacion")
        except Exception as e:
            progress.empty()
            st.error(
                "No se pudo leer la tabla **MatrizSeguimientoEvaluacion** "
                "del archivo de versión anterior. "
                "Verifica que el archivo tenga una tabla (objeto Table de Excel, "
                f"no solo una hoja) con ese nombre exacto. Detalle: {e}"
            )
            st.stop()

        try:
            df_desc_raw = leer_tabla_excel(version_anterior, "OtrosEjecutoresDescentralizadas")
        except Exception as e:
            progress.empty()
            st.error(
                "No se pudo leer la tabla **OtrosEjecutoresDescentralizadas**. "
                f"Detalle: {e}"
            )
            st.stop()

        try:
            df_mun_raw = leer_tabla_excel(version_anterior, "OtrosEjecutoresMunicipios")
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

        # ── Validación de columnas y tipos ───────────────────────────────────
        # Se valida el DataFrame de la versión anterior DESPUÉS de corregir el
        # typo "RECUROS"/"RECURSOS" para no marcar como faltante una columna
        # que existe con el nombre antiguo. La corrección en línea aquí
        # garantiza retrocompatibilidad sin duplicar la lógica.
        for _typo, _correcto in [
            ("FECHA DE INCORPORACIÓN DE RECUROS", "FECHA DE INCORPORACIÓN DE RECURSOS"),
        ]:
            if _typo in df_desc_raw.columns and _correcto not in df_desc_raw.columns:
                df_desc_raw = df_desc_raw.rename({_typo: _correcto})
            if _typo in df_mun_raw.columns and _correcto not in df_mun_raw.columns:
                df_mun_raw = df_mun_raw.rename({_typo: _correcto})

        # Transición: "FECHA SUSCRIPCION" se renombró a
        # "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL". Si la versión anterior
        # aún trae el nombre viejo, se renombra antes de validar para no marcar
        # la columna como faltante.
        def _renombrar_susc(_df):
            _viejo, _nuevo = "FECHA SUSCRIPCION", "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL"
            if _viejo in _df.columns and _nuevo not in _df.columns:
                return _df.rename({_viejo: _nuevo})
            return _df

        df_desc_raw = _renombrar_susc(df_desc_raw)
        df_h1_raw   = _renombrar_susc(df_h1_raw)

        todos_los_problemas = [
            (f"CG-proy ({nombre_proy})",
             validar_columnas(df_proy_raw,  ESQUEMA_GESPROY_PROYECTOS)),
            (f"CG-cttos ({nombre_cttos})",
             validar_columnas(df_cttos_raw, ESQUEMA_GESPROY_CONTRATOS)),
            (f"CG-carga ({nombre_carga})",
             validar_columnas(df_carga_raw, ESQUEMA_GESPROY_CARGUE)),
            ("Versión anterior — MatrizSeguimientoEvaluacion",
             validar_columnas(df_h1_raw,   ESQUEMA_MATRIZ_H1)),
            ("Versión anterior — OtrosEjecutoresDescentralizadas",
             validar_columnas(df_desc_raw, ESQUEMA_MATRIZ_DESC)),
            ("Versión anterior — OtrosEjecutoresMunicipios",
             validar_columnas(df_mun_raw,  ESQUEMA_MATRIZ_MUN)),
        ]
        if sum(len(p) for _, p in todos_los_problemas) > 0:
            progress.empty()
            mostrar_errores_validacion(todos_los_problemas)
            st.stop()

        # ── Consolidación ────────────────────────────────────────────────────
        progress.progress(60, text="Consolidando datos...")
        (
            bpines_version_anterior,
            otros_desc,
            otros_mun,
            fechas_conservadas,
        ) = consolidar(
            df_proy_raw   = df_proy_raw,
            df_cttos_raw  = df_cttos_raw,
            df_carga_raw  = df_carga_raw,
            df_h1_raw     = df_h1_raw,
            df_desc_raw   = df_desc_raw,
            df_mun_raw    = df_mun_raw,
        )

        # ── Vista previa ─────────────────────────────────────────────────────
        progress.progress(80, text="Preparando vista previa...")
        st.divider()
        st.header("Vista previa — Matriz principal")

        if fechas_conservadas:
            st.warning(
                f"Se encontraron **{len(fechas_conservadas)} fecha(s)** que no están en los reportes "
                "de Gesproy y fueron conservadas desde la versión anterior. "
                "Revisa si estas fechas fueron ingresadas manualmente y si siguen siendo válidas."
            )
            st.dataframe(
                pandas.DataFrame(fechas_conservadas),
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
        st.caption(f"{len(bpines_version_anterior)} proyectos")
        try:
            _vista = bpines_version_anterior.select(
                [c for c in cols_preview if c in bpines_version_anterior.columns]
            ).to_pandas()
        except (ImportError, ModuleNotFoundError):
            _vista = pandas.DataFrame(
                bpines_version_anterior.select(
                    [c for c in cols_preview if c in bpines_version_anterior.columns]
                ).to_dicts()
            )
        st.dataframe(_vista, use_container_width=True, height=300)

        # ── Preparación de DataFrames para Excel ─────────────────────────────
        progress.progress(86, text="Preparando datos para Excel...")

        # `to_pandas()` requiere pyarrow en polars 1.x. Usamos un helper que
        # intenta to_pandas y, si falla, hace fallback via to_dicts().
        def _polars_a_pandas(df_pl):
            try:
                return df_pl.to_pandas()
            except (ImportError, ModuleNotFoundError):
                return pandas.DataFrame(df_pl.to_dicts())

        datos_h1   = _polars_a_pandas(bpines_version_anterior.select(todas_las_columnas))
        datos_desc = _polars_a_pandas(otros_desc.select(todas_desc))
        datos_mun  = _polars_a_pandas(otros_mun.select(cols_mun))

        # ── Generación del Excel ─────────────────────────────────────────────
        progress.progress(90, text="Generando Excel...")

        output_buffer = io.BytesIO()
        workbook      = xlsxwriter.Workbook(output_buffer, {"in_memory": True})
        formatos      = crear_formatos(workbook)

        # Hoja 1
        ws1 = workbook.add_worksheet("MatrizSeguimientoEvaluacion")
        escribir_hoja(
            workbook=workbook, ws_hoja=ws1,
            nombre_tabla="MatrizSeguimientoEvaluacion",
            datos_hoja=datos_h1,
            secciones_hoja=[
                ("DATOS GENERALES",         columnas_datos_generales,    formatos["fmt_titulo_azul"]),
                ("DATOS PARA CALIFICACIÓN", columnas_datos_calificacion, formatos["fmt_titulo_naranja"]),
                ("EVALUACIÓN",              columnas_evaluacion,         formatos["fmt_titulo_azul"]),
            ],
            color_col=color_por_columna,
            col_fecha=columnas_fecha_h1, col_numero=columnas_numero_h1,
            col_dias=columnas_dias_h1,   col_con_formula=columnas_con_formula_h1,
            fn_formulas=formulas_para_fila_h1,
            col_ocultar=["COLUMNA APOYO", "COLUMNA APOYO 2"],
            formatos=formatos,
        )

        # Hoja 2
        ws2 = workbook.add_worksheet("OtrosEjecutoresDescentralizadas")
        escribir_hoja(
            workbook=workbook, ws_hoja=ws2,
            nombre_tabla="OtrosEjecutoresDescentralizadas",
            datos_hoja=datos_desc,
            secciones_hoja=[
                ("DATOS GENERALES",         cols_desc_generales,    formatos["fmt_titulo_azul"]),
                ("DATOS PARA CALIFICACIÓN", cols_desc_calificacion, formatos["fmt_titulo_naranja"]),
                ("EVALUACIÓN",              cols_desc_evaluacion,   formatos["fmt_titulo_azul"]),
            ],
            color_col=color_desc,
            col_fecha=columnas_fecha_desc, col_numero=columnas_numero_desc,
            col_dias=columnas_dias_desc,   col_con_formula=columnas_con_formula_desc,
            fn_formulas=formulas_para_fila_desc,
            col_ocultar=["COLUMNA APOYO", "COLUMNA APOYO 2"],
            formatos=formatos,
        )

        # Hoja 3
        ws3 = workbook.add_worksheet("OtrosEjecutoresMunicipios")
        escribir_hoja(
            workbook=workbook, ws_hoja=ws3,
            nombre_tabla="OtrosEjecutoresMunicipios",
            datos_hoja=datos_mun,
            secciones_hoja=[("DATOS GENERALES", cols_mun, formatos["fmt_titulo_azul"])],
            color_col=color_mun,
            col_fecha=columnas_fecha_mun, col_numero=columnas_numero_mun,
            col_dias=set(),              col_con_formula=set(),
            fn_formulas=None,            col_ocultar=[],
            formatos=formatos,
        )

        workbook.close()
        output_buffer.seek(0)
        progress.progress(100, text="Listo.")

        # ── Descarga ─────────────────────────────────────────────────────────
        nombre_archivo = f"MatrizSeguimientoEvaluacion_{datetime.now():%Y%m%d_%H%M}.xlsx"
        st.success(f"Matriz generada con **{len(bpines_version_anterior)} proyectos**.")
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
