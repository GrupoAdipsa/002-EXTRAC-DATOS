"""Utilidades para descargar tablas específicas desde un modelo abierto de ETABS.

El módulo expone la lista de tablas que se extraen por defecto y una función
principal que encapsula la lectura mediante la API de COM. Está pensado para
que otros scripts simplemente importen y llamen a :func:`extraer_tablas_etabs`
sin tener que preocuparse por la interacción directa con ``SapModel``.
"""

from __future__ import annotations

from dataclasses import dataclass
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


@dataclass(frozen=True)
class TablaDisponible:
    """Representa una tabla expuesta por ETABS."""

    key: str
    nombre: str
    import_type: int | None = None
    esta_vacia: bool | None = None


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
        ret, tablas_disponibles = _obtener_tablas_disponibles(sap_model)
    except Exception as exc:  # pragma: no cover - interacción directa con COM
        raise RuntimeError(
            "No se pudieron listar las tablas disponibles desde ETABS."
        ) from exc

    if ret != 0:
        raise RuntimeError(
            f"ETABS devolvió el código {ret} al solicitar el listado de tablas."
        )

    tablas = [tabla.nombre for tabla in tablas_disponibles]
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

    _, tablas_disponibles = _obtener_tablas_disponibles(sap_model)
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
        tabla_destino = _resolver_tabla(nombre_tabla, tablas_disponibles)
        try:
            db_tables.SetAllTablesSelected(False)
            db_tables.SetTableSelected(tabla_destino.key)
        except Exception:
            # Algunos wrappers COM no requieren seleccionar previamente las tablas.
            # Si falla la selección, seguimos e intentamos la lectura directa.
            pass

        try:
            (
                ret,
                headings,
                data,
                _table_version,
                _,
                __,
                ___,
            ) = db_tables.GetTableForDisplayArray(tabla_destino.key)
        except Exception as exc:  # pragma: no cover - interacción directa con COM
            raise RuntimeError(
                f"No se pudo leer la tabla '{tabla_destino.nombre}' desde ETABS: {exc}"
            ) from exc

        if ret != 0:
            raise RuntimeError(
                f"ETABS devolvió el código {ret} al leer la tabla '{tabla_destino.nombre}'."
            )

        headings = list(headings)
        if not headings:
            raise RuntimeError(
                f"La tabla '{tabla_destino.nombre}' no devolvió encabezados desde ETABS."
            )

        if len(data) % len(headings) != 0:
            raise RuntimeError(
                "El tamaño de los datos no coincide con las columnas recibidas "
                f"para la tabla '{tabla_destino.nombre}'."
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


def _resolver_tabla(nombre_solicitado: str, disponibles: list[TablaDisponible]) -> TablaDisponible:
    """Encuentra la tabla solicitada comparando por nombre o key.

    Se busca coincidencia exacta o parcial (case-insensitive) primero sobre
    los nombres de pantalla y luego sobre las keys internas. Si hay múltiples
    coincidencias se solicita mayor precisión para evitar ambigüedades.
    """

    patron = nombre_solicitado.lower()

    for tabla in disponibles:
        if patron == tabla.nombre.lower() or patron == tabla.key.lower():
            return tabla

    coinciden_nombre = [t for t in disponibles if patron in t.nombre.lower()]
    if len(coinciden_nombre) == 1:
        return coinciden_nombre[0]

    coinciden_key = [t for t in disponibles if patron in t.key.lower()]
    if len(coinciden_key) == 1:
        return coinciden_key[0]

    candidatos = {t.nombre for t in coinciden_nombre + coinciden_key}
    if candidatos:
        raise ValueError(
            "El nombre de tabla no es único. Especifícalo con mayor precisión. "
            f"Coincidencias encontradas: {', '.join(sorted(candidatos))}"
        )

    raise ValueError(
        f"La tabla '{nombre_solicitado}' no se encontró en ETABS. "
        "Revisa que el modelo tenga resultados disponibles y que el nombre sea correcto."
    )


def _obtener_tablas_disponibles(sap_model) -> tuple[int, list[TablaDisponible]]:
    """Obtiene las tablas usando ``GetAllTables`` y hace fallback a ``GetAvailableTables``.

    La API de ETABS puede devolver keys y nombres (``GetAllTables``) o solo
    nombres (``GetAvailableTables``). Este helper intenta primero la opción más
    completa y normaliza los resultados para que siempre se disponga de un
    listado de :class:`TablaDisponible`.
    """

    db_tables = sap_model.DatabaseTables

    try:
        tablas_all = _normalizar_get_all_tables(db_tables.GetAllTables())
    except Exception:
        ret = None
    else:
        if tablas_all.ret == 0 and tablas_all.nombres:
            tablas = [
                TablaDisponible(
                    key=key,
                    nombre=nombre,
                    import_type=int(tipo) if tipo is not None else None,
                    esta_vacia=tablas_all.esta_vacia,
                )
                for key, nombre, tipo in zip(
                    tablas_all.keys, tablas_all.nombres, tablas_all.import_types
                )
            ]
            if tablas:
                return tablas_all.ret, tablas

        ret = tablas_all.ret

    resultado = db_tables.GetAvailableTables()
    if not isinstance(resultado, tuple):  # pragma: no cover - defensive
        raise TypeError(
            "GetAvailableTables devolvió un tipo inesperado. Se esperaba un tuple."
        )

    if len(resultado) >= 2:
        ret = resultado[0]
        table_names = list(resultado[1] or [])
        keys = list(resultado[2] or []) if len(resultado) > 2 else []
        import_types = list(resultado[3] or []) if len(resultado) > 3 else []

        if len(keys) < len(table_names):
            keys.extend(table_names[len(keys) :])

        if len(import_types) < len(table_names):
            import_types.extend([None] * (len(table_names) - len(import_types)))

        tablas = [
            TablaDisponible(
                key=str(keys[i]) if i < len(keys) else str(nombre),
                nombre=str(nombre),
                import_type=(
                    int(import_types[i])
                    if i < len(import_types) and import_types[i] is not None
                    else None
                ),
            )
            for i, nombre in enumerate(table_names)
        ]
        return ret, tablas

    raise ValueError(
        "GetAvailableTables no devolvió información de tablas. "
        "Revisa la conexión con ETABS o la versión de la API."
    )


@dataclass
class _ResultadoGetAllTables:
    ret: int
    keys: list[str]
    nombres: list[str]
    import_types: list[int | None]
    esta_vacia: bool | None


def _normalizar_get_all_tables(resultado) -> _ResultadoGetAllTables:
    """Convierte la respuesta de ``GetAllTables`` en datos homogéneos.

    La API puede devolver las tuplas con distintos números de elementos
    dependiendo de la versión (por ejemplo, con o sin ``IsEmpty``). Este
    helper valida la estructura y retorna valores listos para consumo.
    """

    if not isinstance(resultado, tuple):  # pragma: no cover - defensivo
        raise TypeError(
            "GetAllTables devolvió un tipo inesperado. Se esperaba un tuple."
        )

    if len(resultado) < 4:
        raise ValueError(
            "GetAllTables devolvió menos de 4 elementos. Revisa la versión de la API."
        )

    ret = int(resultado[0]) if resultado[0] is not None else -1
    keys = [str(k) for k in (resultado[2] or [])]
    nombres = [str(n) for n in (resultado[3] or [])]
    import_types = [int(t) for t in (resultado[4] or [])] if len(resultado) > 4 else []

    if len(keys) < len(nombres):
        keys.extend(nombres[len(keys) :])

    if len(import_types) < len(nombres):
        import_types.extend([None] * (len(nombres) - len(import_types)))

    esta_vacia = bool(resultado[5]) if len(resultado) > 5 else None

    return _ResultadoGetAllTables(
        ret=ret,
        keys=keys,
        nombres=nombres,
        import_types=import_types,
        esta_vacia=esta_vacia,
    )
