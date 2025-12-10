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

__all__ = [
    "DEFAULT_TABLES",
    "extraer_tablas_etabs",
    "listar_tablas_etabs",
    "diagnosticar_listado_tablas",
]


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


@dataclass(frozen=True)
class PasoDiagnostico:
    """Describe el intento de obtención de tablas y su resultado."""

    metodo: str
    exito: bool
    detalle: str


def diagnosticar_listado_tablas(sap_model) -> tuple[list[TablaDisponible], list[PasoDiagnostico]]:
    """Prueba varias rutas para listar tablas y detalla cuál funcionó.

    Devuelve la lista de tablas disponibles y un log con cada intento, para
    imprimir en consola y saber con precisión qué método de la API respondió.
    """

    if sap_model is None:
        raise ValueError("SapModel no puede ser None. Conecta primero con ETABS.")

    pasos: list[PasoDiagnostico] = []
    db_tables = sap_model.DatabaseTables

    tablas, paso = _intentar_get_all_tables(db_tables)
    pasos.append(paso)
    if paso.exito:
        return tablas, pasos

    for intento in range(1, 1 + 4):
        tablas, paso = _intentar_get_available_tables(db_tables, intento=intento)
        pasos.append(paso)
        if paso.exito:
            return tablas, pasos

    detalle_error = "; ".join(p.detalle for p in pasos if not p.exito)
    error = RuntimeError(
        "No se pudieron listar las tablas disponibles desde ETABS. "
        f"Intentos: {detalle_error or 'sin detalle'}."
    )
    setattr(error, "pasos", pasos)
    raise error


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


def _obtener_tablas_disponibles(
    sap_model, reintentos_available: int = 4
) -> tuple[int, list[TablaDisponible]]:
    """Obtiene las tablas usando ``GetAllTables`` y hace fallback a ``GetAvailableTables``.

    La API de ETABS puede devolver keys y nombres (``GetAllTables``) o solo
    nombres (``GetAvailableTables``). Este helper intenta primero la opción más
    completa y normaliza los resultados para que siempre se disponga de un
    listado de :class:`TablaDisponible`. Si el listado disponible llega vacío o
    la llamada arroja excepciones intermitentes, se vuelve a intentar hasta
    ``reintentos_available`` veces para cubrir respuestas inestables.
    """

    db_tables = sap_model.DatabaseTables

    try:
        ret_all, tablas_all = _normalizar_get_all_tables(db_tables.GetAllTables())
    except Exception:
        ret_all = None
    else:
        if ret_all == 0 and tablas_all:
            return ret_all, tablas_all

    ultimo_ret = -1
    tablas_disp: list[TablaDisponible] = []

    for _ in range(max(1, int(reintentos_available))):
        try:
            ret_disp, tablas_disp = _normalizar_get_available_tables(
                db_tables.GetAvailableTables()
            )
        except Exception:
            continue

        ultimo_ret = ret_disp
        if ret_disp == 0 and tablas_disp:
            return ret_disp, tablas_disp

    return ultimo_ret, tablas_disp


def _intentar_get_all_tables(db_tables) -> tuple[list[TablaDisponible], PasoDiagnostico]:
    """Ejecuta GetAllTables y devuelve tablas o el detalle del fallo."""

    try:
        ret, tablas = _normalizar_get_all_tables(db_tables.GetAllTables())
    except Exception as exc:  # pragma: no cover - interacción COM
        return [], PasoDiagnostico(
            metodo="GetAllTables",
            exito=False,
            detalle=f"excepción: {exc}",
        )

    detalle = f"ret={ret}, tablas={len(tablas)}"
    return tablas, PasoDiagnostico(
        metodo="GetAllTables",
        exito=ret == 0 and bool(tablas),
        detalle=detalle,
    )


def _intentar_get_available_tables(
    db_tables, intento: int | None = None
) -> tuple[list[TablaDisponible], PasoDiagnostico]:
    """Ejecuta GetAvailableTables con normalización defensiva."""

    try:
        ret, tablas = _normalizar_get_available_tables(db_tables.GetAvailableTables())
    except Exception as exc:  # pragma: no cover - interacción COM
        detalle = f"estructura inesperada o excepción: {exc}"
        if intento is not None:
            detalle = f"intento {intento}: {detalle}"

        return [], PasoDiagnostico(
            metodo="GetAvailableTables",
            exito=False,
            detalle=detalle,
        )

    detalle = f"ret={ret}, tablas={len(tablas)}"
    if intento is not None:
        detalle = f"intento {intento}: {detalle}"

    return tablas, PasoDiagnostico(
        metodo="GetAvailableTables",
        exito=ret == 0 and bool(tablas),
        detalle=detalle,
    )


def _normalizar_get_all_tables(resultado) -> tuple[int, list[TablaDisponible]]:
    """Convierte la respuesta de ``GetAllTables`` en datos homogéneos."""

    if not isinstance(resultado, tuple):  # pragma: no cover - defensivo
        raise TypeError(
            "GetAllTables devolvió un tipo inesperado. Se esperaba un tuple."
        )

    if len(resultado) < 3:
        raise ValueError(
            "GetAllTables devolvió menos de 3 elementos. Revisa la versión de la API."
        )

    ret = int(resultado[0]) if resultado[0] is not None else -1
    table_keys = list(resultado[1] or [])
    table_names = list(resultado[2] or [])
    tiene_import_types = len(resultado) > 3 and resultado[3] is not None
    tiene_vacias = len(resultado) > 4 and resultado[4] is not None

    import_types = list(resultado[3] or []) if len(resultado) > 3 else []
    vacias = list(resultado[4] or []) if len(resultado) > 4 else []

    total = max(len(table_keys), len(table_names))

    if not table_names and table_keys:
        table_names = list(table_keys)

    if len(table_keys) < total:
        table_keys.extend([None] * (total - len(table_keys)))
    if len(table_names) < total:
        table_names.extend([None] * (total - len(table_names)))

    if len(import_types) < total:
        import_types.extend([None] * (total - len(import_types)))
    if len(vacias) < total:
        vacias.extend([None] * (total - len(vacias)))

    tablas = []
    for key, nombre, tipo, estado in zip(table_keys, table_names, import_types, vacias):
        key = key if key not in (None, "") else nombre
        nombre = nombre if nombre not in (None, "") else key or ""

        tablas.append(
            TablaDisponible(
                key=str(key),
                nombre=str(nombre),
                import_type=tipo if tiene_import_types else None,
                esta_vacia=estado if tiene_vacias else None,
            )
        )

    return ret, tablas


def _normalizar_get_available_tables(resultado) -> tuple[int, list[TablaDisponible]]:
    """Convierte la respuesta de ``GetAvailableTables`` en una lista homogénea."""

    if not isinstance(resultado, tuple):  # pragma: no cover - defensive
        raise TypeError(
            "GetAvailableTables devolvió un tipo inesperado. Se esperaba un tuple."
        )

    if len(resultado) >= 5:
        ret, table_keys, table_names, import_types, vacias = resultado[:5]
    elif len(resultado) == 4:
        ret, table_keys, table_names, import_types = resultado[:4]
        vacias = None
    elif len(resultado) >= 3:
        ret, table_keys, table_names = resultado[:3]
        import_types = None
        vacias = None
    elif len(resultado) >= 2:
        ret = resultado[0]
        table_keys = resultado[1]
        table_names = resultado[1]
        import_types = None
        vacias = None
    else:
        raise ValueError(
            "GetAvailableTables no devolvió información de tablas. "
            "Revisa la conexión con ETABS o la versión de la API."
        )

    ret = int(ret) if ret is not None else -1

    tiene_import_types = import_types is not None
    tiene_vacias = vacias is not None

    claves = list(table_keys or [])
    nombres = list(table_names or [])

    if not nombres and claves:
        nombres = list(claves)
    if not claves and nombres:
        claves = list(nombres)

    total = max(len(claves), len(nombres))

    if len(claves) < total:
        claves.extend([None] * (total - len(claves)))
    if len(nombres) < total:
        nombres.extend([None] * (total - len(nombres)))

    tipos = list(import_types or [])
    estados_vacios = list(vacias or [])

    if len(tipos) < total:
        tipos.extend([None] * (total - len(tipos)))
    if len(estados_vacios) < total:
        estados_vacios.extend([None] * (total - len(estados_vacios)))

    tablas = []
    for key, nombre, tipo, estado in zip(claves, nombres, tipos, estados_vacios):
        key = key if key not in (None, "") else nombre
        nombre = nombre if nombre not in (None, "") else key or ""

        tablas.append(
            TablaDisponible(
                key=str(key),
                nombre=str(nombre),
                import_type=tipo if tiene_import_types else None,
                esta_vacia=estado if tiene_vacias else None,
            )
        )

    return ret, tablas
