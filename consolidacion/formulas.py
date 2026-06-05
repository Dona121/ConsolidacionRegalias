"""
Generadores de fórmulas Excel para las columnas calculadas dentro del libro.

Las fórmulas están escritas en inglés (sintaxis xlsxwriter), con comas como
separador de argumentos y punto como separador decimal. Esto NO es un bug:
xlsxwriter siempre exporta fórmulas en inglés y Excel las traduce al idioma
del usuario al abrir el archivo.

Restricciones conocidas:
- No usar LET / SWITCH / CAMBIAR (generan bug en metadata.xml).
- Mantener la cadena `CONTRATADO EN EJECUCI\\u00d3N` para evitar problemas de
  encoding con la Ó.

Diferencias entre H1 y Desc para `CALIFICACIÓN EJECUCIÓN DEL PROYECTO`:
- `info_sol`: H1 → INFORMACIÓN SOLICITADA (texto, ISTEXT); Desc → CALIFICACIÓN
  INFORMACIÓN A TIEMPO (número, ISNUMBER).
- Valor si la fórmula da error (IFERROR): H1 → "Revisar"; Desc → "".
- En la rama "CONTRATADO SIN ACTA DE INICIO" (días contra la fecha de corte),
  ambas hojas usan FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO. (Tomado tal cual del
  Excel de referencia en `Contexto/`.)

Lógica de la fórmula:
- "CONTRATADO EN EJECUCIÓN": 0 si el horizonte venció (< fecha de corte); si no,
  cálculo ponderado (×COLUMNA APOYO 2/100 cuando hay externalidades ≥1).
- "CONTRATADO SIN ACTA DE INICIO": 100 si pasaron entre 0 y 30 días desde la
  suscripción hasta la fecha de corte; si no, 0.
- Cualquier otro estado: "".
"""

from __future__ import annotations

from .columnas import idx_h1, idx_desc


def col_letter(i: int) -> str:
    """Convierte índice de columna (0-based) a letra(s) Excel (A, B, …, AA, …)."""
    result = ""
    while True:
        result = chr(i % 26 + 65) + result
        i = i // 26 - 1
        if i < 0:
            break
    return result


# ══════════════════════════════════════════════════════════════════════════════
# ── Generador genérico ────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════
#
# El cuerpo de las fórmulas es prácticamente idéntico entre H1 y Desc. Aquí
# se centraliza para no duplicar la lógica; cada hoja se diferencia por:
#   - el índice de columnas (idx_h1 vs idx_desc),
#   - el nombre exacto del avance físico (AVANCE FISICO en H1, AVANCE FÍSICO en Desc),
#   - la referencia para "información a tiempo" en la fórmula de ejecución,
#   - si la hoja calcula CALIFICACIÓN INFORMACIÓN A TIEMPO o no.
#
def _construir_formulas(
    r: int,
    *,
    idx: dict,
    nombre_avance_fisico: str,
    info_sol_col: str,
    info_sol_es_numero: bool,
    incluir_info_a_tiempo: bool,
    error_ejecucion: str,
    susc_sin_acta_col: str,
) -> dict:
    def ref(c: str) -> str:
        return f"${col_letter(idx[c])}{r}"

    estado     = ref("ESTADO PROYECTO")
    spi        = ref("SPI")
    cpi        = ref("CPI")
    af         = ref("AVANCE FINANCIERO")
    afis       = ref(nombre_avance_fisico)
    col_apoyo  = ref("COLUMNA APOYO")
    col_apoyo2 = ref("COLUMNA APOYO 2")
    info_sol   = ref(info_sol_col)
    cron       = ref("DESEMPEÑO EN EL CRONOGRAMA")
    cost       = ref("DESEMPEÑO EN EL COSTO")
    brecha     = ref("BRECHA FISICO - FINANCIERA")
    ext        = ref("CONTROL EXTERNALIDADES")
    horizonte  = ref("HORIZONTE DEL PROYECTO")
    fcorte     = ref("FECHA DE CORTE GESPROY")
    en_ejecucion = f'{estado}="CONTRATADO EN EJECUCIÓN"'

    def f_indicador(v: str) -> str:
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

    # Diferencia entre H1 y Desc: ISTEXT vs ISNUMBER en la referencia info_sol
    funcion_check = "ISNUMBER" if info_sol_es_numero else "ISTEXT"
    ponderado     = f"({cron}*0.4+{cost}*0.2+{brecha}*0.4)"
    calculo       = f'IFERROR(IF({funcion_check}({info_sol}),{ponderado},""),"")'

    # Rama "CONTRATADO EN EJECUCIÓN":
    #   - Si el horizonte ya venció respecto a la fecha de corte → 0.
    #   - Si hay control de externalidades (>=1) → cálculo ponderado por COLUMNA APOYO 2.
    #   - Si no → cálculo.
    horizonte_vencido = (
        f'AND(ISNUMBER({horizonte}),ISNUMBER({fcorte}),{horizonte}<{fcorte})'
    )
    rama_ejecucion = (
        f'IF({horizonte_vencido},0,'
        f'IF({ext}>=1,({calculo})*{col_apoyo2}/100,{calculo}))'
    )

    # Rama "CONTRATADO SIN ACTA DE INICIO": 100 si pasaron entre 0 y 30 días
    # entre la fecha de suscripción y la fecha de corte; si no, 0. La fecha de
    # suscripción usada difiere por hoja (ver `susc_sin_acta_col`).
    susc      = ref(susc_sin_acta_col)
    dias_susc = f'{fcorte}-{susc}'
    rama_sin_acta = f'IF(AND({dias_susc}>=0,{dias_susc}<=30),100,0)'

    f_ejecucion = (
        f'=IFERROR(IF({en_ejecucion},{rama_ejecucion},'
        f'IF({estado}="CONTRATADO SIN ACTA DE INICIO",{rama_sin_acta},"")),'
        f'{error_ejecucion})'
    )

    formulas = {
        "DESEMPEÑO EN EL CRONOGRAMA":          f_indicador(spi),
        "DESEMPEÑO EN EL COSTO":               f_indicador(cpi),
        "COLUMNA APOYO":                       f_apoyo,
        "BRECHA FISICO - FINANCIERA":          f_brecha,
        "COLUMNA APOYO 2":                     f_apoyo2,
        "CALIFICACIÓN EJECUCIÓN DEL PROYECTO": f_ejecucion,
    }

    # En H1 esta fórmula es automática; en Desc es manual.
    if incluir_info_a_tiempo:
        fecha_recibo = ref("FECHA DE RECIBO DE INFORMACIÓN")
        formulas["CALIFICACIÓN INFORMACIÓN A TIEMPO"] = (
            f'=IF(ISNUMBER({fecha_recibo}),'
            f'IF(DAY({fecha_recibo})=10,100,'
            f'IF(DAY({fecha_recibo})=11,80,'
            f'IF(DAY({fecha_recibo})=12,50,0))),"")'
        )

    return formulas


# ══════════════════════════════════════════════════════════════════════════════
# ── API pública ───────────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════════════

def formulas_para_fila_h1(r: int) -> dict:
    """Fórmulas para la fila `r` (1-based) de la hoja MatrizSeguimientoEvaluacion."""
    return _construir_formulas(
        r,
        idx=idx_h1,
        nombre_avance_fisico="AVANCE FISICO",
        info_sol_col="INFORMACIÓN SOLICITADA",
        info_sol_es_numero=False,
        incluir_info_a_tiempo=True,
        error_ejecucion='"Revisar"',
        susc_sin_acta_col="FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO",
    )


def formulas_para_fila_desc(r: int) -> dict:
    """Fórmulas para la fila `r` (1-based) de la hoja OtrosEjecutoresDescentralizadas."""
    return _construir_formulas(
        r,
        idx=idx_desc,
        nombre_avance_fisico="AVANCE FÍSICO",
        info_sol_col="CALIFICACIÓN INFORMACIÓN A TIEMPO",
        info_sol_es_numero=True,
        incluir_info_a_tiempo=False,
        error_ejecucion='""',
        susc_sin_acta_col="FECHA DE SUSCRIPCIÓN DEL PRIMER CONTRATO",
    )


__all__ = [
    "col_letter",
    "formulas_para_fila_h1",
    "formulas_para_fila_desc",
]
