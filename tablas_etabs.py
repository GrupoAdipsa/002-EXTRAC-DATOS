"""Utilidades para descargar tablas específicas desde un modelo abierto de ETABS.

El módulo expone la lista de tablas que se extraen por defecto y una función
principal que encapsula la lectura mediante la API de COM. Está pensado para
que otros scripts simplemente importen y llamen a :func:`extraer_tablas_etabs`
sin tener que preocuparse por la interacción directa con ``SapModel``.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd


DEFAULT_TABLES: list[str] = [
    "Story Forces",
    "Diaphragm Accelerations",
    "Story Drifts",
]
"""Tablas de ETABS que se extraen por defecto."""

__all__ = ["DEFAULT_TABLES", "extraer_tablas_etabs", "listar_tablas_etabs"]


def listar_tablas_etabs(sap_model, filtro: str | None = None) -> list[str]:
    """Devuelve las tablas disponibles en el modelo abierto de ETABS.

    Parameters
    ----------
    sap_model : SapModel
        Objeto SapModel obtenido desde ETABS mediante la API de COM.
    filtro : str, optional
        Cadena para filtrar las tablas que contengan ese texto (no distingue
        mayúsculas/minúsculas). Si es ``None`` se devuelven todas.

    Returns
    -------
    list[str]
        Nombres exactos de las tablas que ETABS expone para el modelo actual.
    """

    if sap_model is None:
        raise ValueError("SapModel no puede ser None. Conecta primero con ETABS.")

    try:
        ret, table_names = _leer_tablas_disponibles(sap_model)
    except Exception as exc:  # pragma: no cover - interacción directa con COM
        raise RuntimeError(
            "No se pudieron listar las tablas disponibles desde ETABS."
        ) from exc

    if ret != 0:
        raise RuntimeError(
            f"ETABS devolvió el código {ret} al solicitar el listado de tablas."
        )

    tablas = list(table_names)
    if filtro:
        patron = filtro.lower()
        tablas = [nombre for nombre in tablas if patron in nombre.lower()]

    return tablas


def extraer_tablas_etabs(
    sap_model,
    tablas: Iterable[str] | None = None,
    carpeta_destino: str | Path | None = None,
    formatos: str | Iterable[str] = "csv",
):
    """Extrae tablas seleccionadas de un modelo abierto de ETABS.

    Parameters
    ----------
    sap_model : SapModel
        Objeto SapModel obtenido desde ETABS mediante la API de COM.
    tablas : Iterable[str], optional
        Nombres exactos de las tablas a extraer. Si no se proporcionan se
        usan las tablas de historia de nivel, aceleraciones de diafragma y
        derivas entre pisos.
    carpeta_destino : str or Path, optional
        Ruta de carpeta donde se exportarán las tablas en archivos CSV o TXT.
        Si no se indica, los :class:`pandas.DataFrame` se devuelven en un
        diccionario sin escribir a disco.
    formatos : str or Iterable[str], default "csv"
        Formatos de exportación cuando se entrega ``carpeta_destino``. Acepta
        ``"csv"``, ``"txt"`` o una colección con ambos para generar las dos
        variantes. Ignorado si no se especifica carpeta de destino.

    Returns
    -------
    dict[str, pandas.DataFrame]
        Diccionario con el nombre de cada tabla y su información convertida
        a DataFrame. Cuando ``carpeta_destino`` es ``None`` se devuelven los
        DataFrames en memoria; de lo contrario, también se escriben en
        archivos CSV.
    """

    if sap_model is None:
        raise ValueError("SapModel no puede ser None. Conecta primero con ETABS.")

    tablas_a_extraer = list(tablas or DEFAULT_TABLES)
    if not tablas_a_extraer:
        raise ValueError("No se proporcionaron tablas a extraer.")

    tablas_disponibles = listar_tablas_etabs(sap_model)
    if not tablas_disponibles:
        raise RuntimeError(
            "ETABS no devolvió ningún nombre de tabla disponible para el modelo abierto."
        )

    destino: Path | None = None
    formatos_normalizados: list[str] = []
    if carpeta_destino:
        destino = Path(carpeta_destino)
        destino.mkdir(parents=True, exist_ok=True)
        formatos_normalizados = _normalizar_formatos(formatos)

    db_tables = sap_model.DatabaseTables
    resultados: dict[str, pd.DataFrame] = {}

    for nombre_tabla in tablas_a_extraer:
        nombre_tabla_etabs = _resolver_nombre_tabla(nombre_tabla, tablas_disponibles)
        try:
            (
                ret,
                headings,
                data,
                _table_version,
                _,
                __,
                ___,
            ) = db_tables.GetTableForDisplayArray(nombre_tabla_etabs)
        except Exception as exc:  # pragma: no cover - interacción directa con COM
            raise RuntimeError(
                f"No se pudo leer la tabla '{nombre_tabla_etabs}' desde ETABS: {exc}"
            ) from exc

        if ret != 0:
            raise RuntimeError(
                f"ETABS devolvió el código {ret} al leer la tabla '{nombre_tabla_etabs}'."
            )

        headings = list(headings)
        if not headings:
            raise RuntimeError(
                f"La tabla '{nombre_tabla_etabs}' no devolvió encabezados desde ETABS."
            )

        if len(data) % len(headings) != 0:
            raise RuntimeError(
                "El tamaño de los datos no coincide con las columnas recibidas "
                f"para la tabla '{nombre_tabla_etabs}'."
            )

        filas = len(data) // len(headings)
        registros = [data[i * len(headings):(i + 1) * len(headings)] for i in range(filas)]
        df = pd.DataFrame(registros, columns=headings)
        resultados[nombre_tabla] = df

        if destino:
            nombre_archivo_base = nombre_tabla.replace(" ", "_").lower()

            for formato in formatos_normalizados:
                if formato == "csv":
                    ruta = destino / f"{nombre_archivo_base}.csv"
                    df.to_csv(ruta, index=False, encoding="utf-8")
                elif formato == "txt":
                    ruta = destino / f"{nombre_archivo_base}.txt"
                    df.to_csv(ruta, index=False, encoding="utf-8", sep="\t")
                else:  # pragma: no cover - ya validado
                    raise ValueError(f"Formato de exportación no soportado: {formato}")

    return resultados


def _normalizar_formatos(formatos: str | Iterable[str]) -> list[str]:
    """Valida y normaliza la lista de formatos a usar en disco."""

    formatos_validos = {"csv", "txt"}

    if isinstance(formatos, str):
        lista_formatos = [formatos]
    else:
        lista_formatos = list(formatos)

    if not lista_formatos:
        raise ValueError(
            "Debes especificar al menos un formato de exportación (csv y/o txt)."
        )

    normalizados: list[str] = []
    for formato in lista_formatos:
        if not isinstance(formato, str):
            raise TypeError("Los formatos deben ser cadenas (por ejemplo 'csv' o 'txt').")

        formato_limpio = formato.strip().lower()
        if formato_limpio not in formatos_validos:
            raise ValueError(
                f"Formato no soportado: {formato_limpio}. Usa 'csv', 'txt' o ambos."
            )

        if formato_limpio not in normalizados:
            normalizados.append(formato_limpio)

    return normalizados


def _resolver_nombre_tabla(nombre_solicitado: str, disponibles: list[str]) -> str:
    """Encuentra el nombre de tabla tal como lo expone ETABS.

    Si ``nombre_solicitado`` no aparece de forma exacta en ``disponibles`` se
    realiza una búsqueda parcial (case-insensitive) para localizar un nombre
    que contenga el texto indicado. Esto permite usar etiquetas abreviadas como
    "Story Drifts" aunque ETABS la exponga como "Analysis Results - Story Drifts".
    """

    if nombre_solicitado in disponibles:
        return nombre_solicitado

    patron = nombre_solicitado.lower()
    candidatos = [nombre for nombre in disponibles if patron in nombre.lower()]
    if len(candidatos) == 1:
        return candidatos[0]

    if candidatos:
        raise ValueError(
            "El nombre de tabla no es único. Especifícalo con mayor precisión. "
            f"Coincidencias encontradas: {', '.join(sorted(candidatos))}"
        )

    raise ValueError(
        f"La tabla '{nombre_solicitado}' no se encontró en ETABS. "
        "Revisa que el modelo tenga resultados disponibles y que el nombre sea correcto."
    )


def _leer_tablas_disponibles(sap_model) -> tuple[int, list[str]]:
    """Obtiene y normaliza el resultado de ``GetAvailableTables``.

    La API de ETABS puede entregar un número variable de valores al llamar a
    ``DatabaseTables.GetAvailableTables`` dependiendo de la versión y del
    wrapper COM usado. Este helper acepta las variantes más comunes y extrae
    siempre un tuple con ``(ret, table_names)``.
    """

    resultado = sap_model.DatabaseTables.GetAvailableTables()
    if not isinstance(resultado, tuple):  # pragma: no cover - defensive
        raise TypeError(
            "GetAvailableTables devolvió un tipo inesperado. Se esperaba un tuple."
        )

    if len(resultado) >= 2:
        ret = resultado[0]
        table_names = resultado[1] or []
        return ret, table_names

    raise ValueError(
        "GetAvailableTables no devolvió información de tablas. "
        "Revisa la conexión con ETABS o la versión de la API."
    )
