"""
Paquete `consolidacion` — módulos de la Matriz de Seguimiento y Evaluación SGR.

Cada módulo tiene una responsabilidad acotada para evitar duplicación de lógica
entre archivos. Importa lo que necesites desde el módulo correspondiente:

    from consolidacion import (
        esquemas, columnas, diseno, validacion, lectura, calculos,
        formulas, escritura, procesamiento,
    )
"""

from . import (
    esquemas,
    columnas,
    diseno,
    validacion,
    lectura,
    calculos,
    formulas,
    escritura,
    procesamiento,
)

__all__ = [
    "esquemas",
    "columnas",
    "diseno",
    "validacion",
    "lectura",
    "calculos",
    "formulas",
    "escritura",
    "procesamiento",
]
