"""
Esquemas esperados por archivo / tabla.

Cada entrada del diccionario asocia el nombre exacto de la columna con una
tupla `(tipo_legible, categoria)`:

    tipo_legible : str
        Texto que se muestra al usuario en la referencia de columnas
        ("Texto", "Número" o "Fecha").

    categoria    : str
        Tipo de validación que se aplica al contenido de la columna:

        - "texto"  → debe ser String; si es numérico se reporta error
                     (caso típico: BPIN que Excel convierte a número).
        - "numero" → debe poder convertirse a Float64.
        - "fecha"  → debe poder convertirse a Date.
        - "libre"  → solo se valida que la columna esté presente.

Estos esquemas se usan tanto para la validación de los archivos cargados
como para la referencia visible en la interfaz.
"""


# ── Reportes Gesproy ──────────────────────────────────────────────────────────

ESQUEMA_GESPROY_PROYECTOS = {
    "BPIN":                         ("Texto",  "texto"),
    "NOMBRE PROYECTO":              ("Texto",  "texto"),
    "SECTOR":                       ("Texto",  "texto"),
    "ESTADO PROYECTO":              ("Texto",  "texto"),
    "VALOR SGR":                    ("Número", "numero"),
    "VALOR NACIÓN":                 ("Número", "numero"),
    "VALOR OTROS":                  ("Número", "numero"),
    "VALOR OTRAS FUENTES NO SUIFP": ("Número", "numero"),
    "VALOR TOTAL PROYECTO":         ("Número", "numero"),
    "VALOR PAGOS":                  ("Número", "numero"),
}

ESQUEMA_GESPROY_CONTRATOS = {
    "BPIN":                    ("Texto",  "texto"),
    "FECHA ACT ADTIVO APERT":  ("Fecha",  "fecha"),
    "ESTADO CONTRATO":         ("Texto",  "texto"),
    "FECHA INICIAL":           ("Fecha",  "fecha"),
    "FECHA SUSCRIPCION":       ("Fecha",  "fecha"),
    "ULTIMA FECHA PAGO":       ("Fecha",  "fecha"),
    "TIPO CONTRATO":           ("Texto",  "texto"),
    "VALOR TOTAL FUENTES SGR": ("Número", "numero"),
}

ESQUEMA_GESPROY_CARGUE = {
    "BPIN":                      ("Texto",  "texto"),
    "FECHA APROBACIÓN PROYECTO": ("Fecha",  "fecha"),
    "AVANCE FISICO":             ("Número", "numero"),
    "AVANCE FINANCIERO":         ("Número", "numero"),
}


# ── Tablas de la versión anterior de la Matriz ────────────────────────────────

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
    # Fecha que migra de la versión anterior (estado para cierre)
    "FECHA EN LA QUE PASO A ESTADO PARA CIERRE": ("Fecha", "fecha"),
    "CPI":                                ("Número", "numero"),
    "SPI":                                ("Número", "numero"),
    "INFORMACIÓN SOLICITADA":             ("Texto",  "libre"),
    "INFORMACIÓN RECIBIDA":               ("Texto",  "libre"),
    "FECHA DE RECIBO DE INFORMACIÓN":     ("Fecha",  "fecha"),
    "CONTROL EXTERNALIDADES":             ("Número", "numero"),
    "FECHA DE CORTE GESPROY":             ("Fecha",  "fecha"),
    # Ahora migra de la versión anterior (ya no se calcula en consolidación)
    "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN": ("Número", "numero"),
    "CALIFICACIÓN CALIDAD INFORMACIÓN":   ("Número", "numero"),
    "COMENTARIOS CALIFICACIÓN":           ("Texto",  "libre"),
    # Columna manual nueva (no migra de los informes Gesproy)
    "RESPONSABLE CARGUE EN GESPROY":      ("Texto",  "libre"),
    # Columna manual nueva: se conserva desde la versión anterior
    "MUNICIPIOS":                         ("Texto",  "libre"),
}

ESQUEMA_MATRIZ_DESC = {
    "BPIN":                                      ("Texto",  "texto"),
    "EJECUTOR":                                  ("Texto",  "libre"),
    "NOMBRE DEL PROYECTO":                       ("Texto",  "libre"),
    "ALCANCE":                                   ("Texto",  "libre"),
    "SECTOR":                                    ("Texto",  "libre"),
    "FUENTE":                                    ("Texto",  "libre"),
    "ESTADO PROYECTO":                           ("Texto",  "texto"),
    "ESTADO CONTRATO":                           ("Texto",  "libre"),
    "VALOR SGR":                                 ("Número", "numero"),
    "VALOR OTROS":                               ("Número", "numero"),
    "VALOR TOTAL":                               ("Número", "numero"),
    "FECHA DE MIGRACIÓN A GESPROY":              ("Fecha",  "fecha"),
    "FECHA DE ASIGNACIÓN DE RECURSOS":           ("Fecha",  "fecha"),
    # Typo corregido: antes "RECUROS", ahora "RECURSOS"
    "FECHA DE INCORPORACIÓN DE RECURSOS":        ("Fecha",  "fecha"),
    "AVANCE FÍSICO":                             ("Número", "numero"),
    "AVANCE FINANCIERO":                         ("Número", "numero"),
    "CPI":                                       ("Número", "numero"),
    "SPI":                                       ("Número", "numero"),
    "FECHA APROBACIÓN PROYECTO":                 ("Fecha",  "fecha"),
    "FECHA DE APERTURA DEL PRIMER PROCESO":      ("Fecha",  "fecha"),
    # Renombrada (antes "FECHA SUSCRIPCION"). La fecha del primer contrato es
    # opcional (no se valida); migra de la versión anterior si está presente.
    "FECHA DE SUSCRIPCIÓN DEL CONTRATO PRINCIPAL": ("Fecha",  "fecha"),
    "FECHA ACTA INICIO":                         ("Fecha",  "fecha"),
    # Columna manual nueva (no migra de los informes Gesproy)
    "HORIZONTE DEL PROYECTO":                    ("Fecha",  "fecha"),
    # Fecha que migra de la versión anterior (estado para cierre)
    "FECHA EN LA QUE PASO A ESTADO PARA CIERRE": ("Fecha",  "fecha"),
    "FECHA DE CORTE GESPROY":                    ("Fecha",  "fecha"),
    "CALIFICACIÓN DESEMPEÑO EN LA CONTRATACIÓN": ("Número", "numero"),
    "CALIFICACIÓN INFORMACIÓN A TIEMPO":         ("Número", "numero"),
    "CALIFICACIÓN CALIDAD INFORMACIÓN":          ("Número", "numero"),
    "CONTROL EXTERNALIDADES":                    ("Número", "numero"),
    "COMENTARIOS CALIFICACIÓN":                  ("Texto",  "libre"),
    # Columna manual nueva: se conserva desde la versión anterior
    "MUNICIPIOS":                                ("Texto",  "libre"),
}

ESQUEMA_MATRIZ_MUN = {
    "BPIN":                               ("Texto",  "texto"),
    "EJECUTOR":                           ("Texto",  "libre"),
    "NOMBRE DEL PROYECTO":                ("Texto",  "libre"),
    "ALCANCE":                            ("Texto",  "libre"),
    "SECTOR":                             ("Texto",  "libre"),
    "FUENTE":                             ("Texto",  "libre"),
    "ESTADO PROYECTO":                    ("Texto",  "texto"),
    "ESTADO CONTRATO":                    ("Texto",  "libre"),
    "VALOR SGR":                          ("Número", "numero"),
    "VALOR OTROS":                        ("Número", "numero"),
    "VALOR TOTAL":                        ("Número", "numero"),
    "FECHA APROBACIÓN PROYECTO":          ("Fecha",  "fecha"),
    "FECHA DE ASIGNACIÓN DE RECURSOS":    ("Fecha",  "fecha"),
    # Typo corregido: antes "RECUROS", ahora "RECURSOS"
    "FECHA DE INCORPORACIÓN DE RECURSOS": ("Fecha",  "fecha"),
    "FECHA ACTA INICIO":                  ("Fecha",  "fecha"),
    "AVANCE FÍSICO":                      ("Número", "numero"),
    "AVANCE FINANCIERO":                  ("Número", "numero"),
    "COMENTARIOS":                        ("Texto",  "libre"),
    # Columna manual nueva: se conserva desde la versión anterior
    "MUNICIPIOS":                         ("Texto",  "libre"),
}


__all__ = [
    "ESQUEMA_GESPROY_PROYECTOS",
    "ESQUEMA_GESPROY_CONTRATOS",
    "ESQUEMA_GESPROY_CARGUE",
    "ESQUEMA_MATRIZ_H1",
    "ESQUEMA_MATRIZ_DESC",
    "ESQUEMA_MATRIZ_MUN",
]
