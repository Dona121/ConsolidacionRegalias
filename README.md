# Matriz de Seguimiento y Evaluación — Regalías SGR

Aplicación web (Streamlit) que **consolida los reportes mensuales de Gesproy** del
Sistema General de Regalías (SGR) y genera la *Matriz de Seguimiento y Evaluación*
en un único archivo Excel, con formato, fórmulas e indicadores listos para usar.

El equipo descarga tres reportes de Gesproy más la Matriz del mes anterior, los carga
en la app, y obtiene una Matriz nueva que refleja el estado actual de Gesproy sin perder
la información que se ingresó manualmente.

---

## Tabla de contenidos

- [Qué hace](#qué-hace)
- [Archivos de entrada](#archivos-de-entrada)
- [Salida](#salida)
- [Cómo funciona la consolidación](#cómo-funciona-la-consolidación)
- [Indicadores y calificaciones](#indicadores-y-calificaciones)
- [Arquitectura del código](#arquitectura-del-código)
- [Instalación y ejecución](#instalación-y-ejecución)
- [Despliegue en Railway](#despliegue-en-railway)
- [Notas técnicas](#notas-técnicas)

---

## Qué hace

1. Recibe tres reportes de Gesproy (proyectos, contratos, cargue) y la Matriz del mes anterior.
2. Valida que cada archivo tenga las columnas y tipos de dato esperados, con mensajes de error
   en lenguaje claro que explican **cómo corregirlos en Excel**.
3. Cruza todo por **BPIN** (código único del proyecto).
4. Calcula los indicadores de **días** en Python (la *calificación de contratación* ya no se
   calcula: migra de la versión anterior).
5. Genera un Excel con tres hojas (tablas con nombre), encabezados de colores, columnas
   con formato de fecha/número y fórmulas Excel que se recalculan al abrir el archivo.
6. Permite descargar el resultado.

---

## Archivos de entrada

| Archivo | Prefijo del nombre | Qué contiene | Qué aporta |
|---|---|---|---|
| **CG-proy** | `CG-proy…` | Listado de todos los proyectos SGR | Base de proyectos activos (excluye estados *Cerrado* y *Desaprobado*). |
| **CG-cttos** | `CG-cttos…` | Contratos de cada proyecto | Estado y tipo de contrato, fechas (apertura, suscripción, acta de inicio, último pago) y valor. |
| **CG-carga** | `CG-carga…` | Cargue de avances | Avance físico, financiero y fecha de aprobación del proyecto. |
| **Versión anterior** | — | Matriz del mes pasado (`.xlsx`) | Contexto del proyecto, datos manuales (CPI, SPI, externalidades, calificaciones) y fechas ingresadas a mano. |

**Sobre los nombres:** los tres reportes de Gesproy se identifican por su prefijo. Si subes varios
con el mismo prefijo, la app toma el de fecha más reciente (detectada por el patrón `YYYYMMDD` del nombre).

**Sobre la versión anterior:** debe contener tres *tablas con nombre* (objetos *Table* de Excel,
no solo hojas): `MatrizSeguimientoEvaluacion`, `OtrosEjecutoresDescentralizadas` y
`OtrosEjecutoresMunicipios`.

> La referencia completa de columnas esperadas por archivo está disponible dentro de la app,
> en el desplegable *"Referencia: columnas esperadas por archivo"*.

---

## Salida

Un Excel (`MatrizSeguimientoEvaluacion_YYYYMMDD_HHMM.xlsx`) con tres hojas/tablas:

1. **MatrizSeguimientoEvaluacion** — hoja principal, con secciones *Datos Generales*,
   *Datos para Calificación* y *Evaluación*. Incluye todas las fórmulas automáticas.
2. **OtrosEjecutoresDescentralizadas** — misma estructura; la *Calificación información a tiempo*
   es manual en esta hoja.
3. **OtrosEjecutoresMunicipios** — solo datos generales, **sin fórmulas automáticas**
   (toda la calificación es manual).

---

## Cómo funciona la consolidación

El cruce se hace siempre por **BPIN**:

1. **Base:** proyectos activos de CG-proy.
2. Se agrega el contexto de la versión anterior (alcance, entidad, indicador MGA, etc.).
3. Se agrega el **contrato más representativo** de cada BPIN.
4. Se agrega el avance físico/financiero y la fecha de aprobación del cargue.

### Priorización de contratos

Cuando un BPIN tiene varios contratos, se elige uno solo por orden de prioridad de tipo y,
dentro de cada categoría, el de **mayor VALOR SGR**:

1. Obra pública, Consultoría, Convenios de Cooperación, Interadministrativos
2. Suministro, Contratos/convenios con entidades sin ánimo de lucro
3. Cualquier otro tipo

### Priorización de fechas: Gesproy vs. fechas manuales

Para las fechas de contratación (`FECHA APROBACIÓN PROYECTO`, `FECHA DE APERTURA DEL PRIMER PROCESO`,
`FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL`, `FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO`,
`FECHA ACTA INICIO`, `ULTIMA FECHA PAGO`):

| Situación | Resultado |
|---|---|
| El BPIN **no aparece** en contratos | Conserva la fecha manual de la Matriz anterior. |
| El BPIN aparece pero la fecha está **vacía** | Conserva la fecha manual. |
| El BPIN aparece **y la fecha está registrada** | Usa la fecha de Gesproy (prioridad). |

Así, las fechas ingresadas a mano no se pierden mientras el proyecto avanza; cuando Gesproy
finalmente las registre, se tomarán automáticamente. Al terminar, la app muestra una tabla
con todas las fechas que se conservaron desde la versión anterior.

> **Contrato principal vs. primer contrato.** `FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL` es la
> fecha del contrato representativo elegido por BPIN (ver priorización arriba). Como ese contrato
> puede no ser el primero que se firmó, se agrega `FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO`, que en
> la hoja principal se calcula como la **fecha de suscripción más temprana** entre todos los
> contratos del BPIN en CG-cttos (`min` por BPIN, calculado antes de elegir el principal). En
> *Descentralizadas* esta columna se conserva de la versión anterior. (La columna antes se llamaba
> `FECHA SUSCRIPCION`; si la versión anterior aún trae ese nombre, la app lo renombra
> automáticamente.)

### Fecha de corte

Se fija automáticamente al **día 15 del mes actual**. Si la versión anterior ya tenía una fecha
de corte para un proyecto, se conserva esa.

### Columnas manuales y que migran de la versión anterior

Algunas columnas no provienen de Gesproy: se conservan tal cual entre versiones de la Matriz.

- `RESPONSABLE CARGUE EN GESPROY` (hoja principal) y `HORIZONTE DEL PROYECTO` (Descentralizadas):
  manuales, no migran de Gesproy.
- `MUNICIPIOS` (en las tres hojas, al final, con encabezado verde `#9BBB59`): manual.
- `CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN`: **ya no se calcula** en la consolidación; ahora
  **migra de la versión anterior** y el equipo la edita a mano.
- `FECHA EN LA QUE PASO A ESTADO PARA CIERRE` (hoja principal y Descentralizadas, formato fecha):
  migra de la versión anterior.

---

## Indicadores y calificaciones

**Días transcurridos** (calculados en Python, solo para estados específicos):

- *Aprobación → apertura del primer proceso* — proyectos *Sin contratar*.
- *Apertura → firma del primer contrato* — *Sin contratar* con fecha de apertura.
- *Suscripción → acta de inicio* — *Contratado sin acta de inicio*.

**Calificación desempeño en la contratación**: **ya no se calcula** en la consolidación.
Migra de la versión anterior de la Matriz (ver *Columnas manuales y que migran*).

**Fórmulas Excel** (se recalculan al abrir el archivo, no se guardan valores fijos):

- *Desempeño en el cronograma (SPI)* y *en el costo (CPI)*: escala con zona óptima entre 0.84 y 1.2.
- *Brecha físico-financiera*: compara avance físico vs. financiero con tolerancias por nivel de avance.
- *Calificación ejecución del proyecto*: depende del estado del proyecto:
  - *Contratado en ejecución*: pondera cronograma (40%), costo (20%) y brecha (40%), ajustada por
    externalidades (×`COLUMNA APOYO 2`/100 cuando hay externalidades ≥ 1). Si el **horizonte ya
    venció** respecto a la fecha de corte, la calificación es **0**.
  - *Contratado sin acta de inicio*: **100** si pasaron entre 0 y 30 días entre la suscripción
    (del primer contrato) y la fecha de corte; en otro caso **0**.
  - Cualquier otro estado: vacío.
- *Calificación información a tiempo* (solo hoja principal): según el día de recibo
  (día 10 = 100, 11 = 80, 12 = 50, otro = 0).

Las columnas auxiliares *Columna Apoyo* y *Columna Apoyo 2* son intermedias y quedan **ocultas** en el Excel.

---

## Arquitectura del código

La lógica está modularizada en el paquete `consolidacion/`. La app (`app_matriz_seguimiento.py`)
solo orquesta la UI y llama a los módulos.

```
.
├── app_matriz_seguimiento.py   # App Streamlit: UI, validación y orquestación
├── consolidacion/
│   ├── esquemas.py             # Columnas y tipos esperados por archivo/tabla
│   ├── columnas.py             # Orden de columnas y mapeo columna → color por hoja
│   ├── diseno.py               # Diseño visual: paleta, formatos y layout del Excel
│   ├── validacion.py           # validar_columnas + render de errores en Streamlit
│   ├── lectura.py              # Lectura de Excel/tablas + normalización de fechas
│   ├── calculos.py             # Días y calificación de contratación (en Python)
│   ├── formulas.py             # Generadores de fórmulas Excel (Hoja 1 y Descentralizadas)
│   ├── escritura.py            # escribir_hoja: mecánica de escritura (xlsxwriter)
│   └── procesamiento.py        # consolidar(): joins por BPIN, prioriza contratos y fechas
├── Contexto/
│   └── ConsolidacionRegalias.py  # Versión monolítica original (referencia histórica)
├── requirements.txt            # Dependencias para Railway
├── pyproject.toml / uv.lock    # Gestión local con uv
└── README.md
```

> `Contexto/ConsolidacionRegalias.py` es la **versión original en un solo archivo**, conservada
> como referencia de la lógica completa. La app en producción usa el paquete `consolidacion/`.

### Flujo de `consolidar()`

1. Filtra proyectos activos (excluye `CERRADO` y `DESAPROBADO`).
2. Selecciona el contrato representativo por BPIN y calcula la fecha de suscripción del primer
   contrato (`min` por BPIN) antes de la selección.
3. Construye la tabla de fechas con la regla *Gesproy > manual*.
4. Hace los joins por BPIN (proyectos + contexto + fechas + contratos + cargue).
5. Aplica los cálculos de **días** (la calificación de contratación migra de la versión anterior).
6. Devuelve los tres DataFrames + la lista de fechas conservadas.

---

## Instalación y ejecución

### Con uv (recomendado para desarrollo local)

```bash
uv sync                                          # crea el entorno desde uv.lock
uv run streamlit run app_matriz_seguimiento.py   # ejecuta la app
```

### Con pip / venv

```bash
python -m venv .venv
source .venv/bin/activate          # en fish: source .venv/bin/activate.fish
pip install -r requirements.txt
streamlit run app_matriz_seguimiento.py
```

La app queda disponible en `http://localhost:8501`.

### Dependencias principales

`streamlit`, `polars` (procesamiento de datos), `pandas`, `openpyxl` / `fastexcel`
(lectura de Excel), `xlsxwriter` (generación del Excel), `plotly`.

---

## Despliegue en Railway

El despliegue usa **`requirements.txt`** (Railway lo detecta automáticamente para proyectos Python).

Comando de arranque sugerido (Railway expone el puerto en la variable `$PORT`):

```bash
streamlit run app_matriz_seguimiento.py --server.port $PORT --server.address 0.0.0.0
```

> Si agregas dependencias con `uv add`, recuerda reflejarlas también en `requirements.txt`
> (o regéneralo con `uv export --no-hashes --format requirements-txt > requirements.txt`)
> para que el build de Railway no se quede atrás.

---

## Notas técnicas

- **Fórmulas en inglés:** xlsxwriter siempre escribe las fórmulas en sintaxis inglesa
  (comas como separador, punto decimal). Excel las traduce al idioma del usuario al abrir.
  No es un bug.
- **Tablas con nombre:** la versión anterior debe tener objetos *Table* de Excel, no solo hojas.
  `leer_tabla_excel` usa polars si soporta `table_name`, con fallback a openpyxl.
- **Normalización de fechas:** `normalizar_fecha` convierte a `pl.Date` sin importar si el dato
  llegó como `Date`, `Datetime`, número de serie de Excel o texto (`dd/mm/yyyy` o `yyyy-mm-dd`).
- **Compatibilidad de columnas:** la app corrige en línea el typo histórico
  `FECHA DE INCORPORACIÓN DE RECUROS` → `…RECURSOS`, renombra `FECHA SUSCRIPCION` →
  `FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL` si la versión anterior aún trae el nombre viejo, y
  agrega columnas opcionales vacías si una versión anterior no las tenía, para no romper la
  consolidación.
