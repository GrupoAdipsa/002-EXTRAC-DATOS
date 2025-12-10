"""Punto de entrada simplificado para extraer tablas de ETABS.

Importa y reexpone la lista de tablas predeterminadas y las funciones
:func:`tablas_etabs.extraer_tablas_etabs` y
:func:`tablas_etabs.listar_tablas_etabs`. Úsalo cuando quieras indicar
qué tablas descargar (por ejemplo pasando tu propia lista de forma
programática), el formato de exportación (CSV y/o TXT) o consultar el
catálogo completo sin preocuparte por la lógica de conexión a
`SapModel`.

Además, expone :func:`gui_tablas_etabs.lanzar_gui_etabs` para abrir una
ventana gráfica con lista de tablas, casillas de formato y selector de
carpeta.
"""
from gui_tablas_etabs import lanzar_gui_etabs
from tablas_etabs import (
    DEFAULT_TABLES,
    diagnosticar_listado_tablas,
    extraer_tablas_etabs,
    listar_tablas_etabs,
)

__all__ = [
    "DEFAULT_TABLES",
    "diagnosticar_listado_tablas",
    "extraer_tablas_etabs",
    "listar_tablas_etabs",
    "lanzar_gui_etabs",
]
