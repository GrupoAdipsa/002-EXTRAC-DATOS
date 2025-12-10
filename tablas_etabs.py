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

    if ret != 0 and not tablas_disponibles:
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
    casos: Iterable[str] | None = None,
    combinaciones: Iterable[str] | None = None,
    debug_log: bool = False,
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
    debug_dir = destino or Path.cwd()

    db_tables = sap_model.DatabaseTables
    resultados: dict[str, pd.DataFrame] = {}

    # Selección de casos/combinaciones si se solicita
    casos_lista = list(casos or [])
    combos_lista = list(combinaciones or [])

    if casos_lista or combos_lista:
        try:
            setup = sap_model.Results.Setup
            try:
                setup.DeselectAllCasesAndCombosForOutput()
            except Exception:
                # Si falla, seguimos e intentamos seleccionar igualmente.
                pass

            errores_sel: list[str] = []
            for nombre in casos_lista:
                try:
                    ret_sel = setup.SetCaseSelectedForOutput(nombre, True)
                except Exception as exc_case:
                    errores_sel.append(f"case '{nombre}': {exc_case}")
                else:
                    if ret_sel != 0:
                        errores_sel.append(f"case '{nombre}' ret={ret_sel}")

            for nombre in combos_lista:
                try:
                    ret_sel = setup.SetComboSelectedForOutput(nombre, True)
                except Exception as exc_combo:
                    errores_sel.append(f"combo '{nombre}': {exc_combo}")
                else:
                    if ret_sel != 0:
                        errores_sel.append(f"combo '{nombre}' ret={ret_sel}")

            casos_sel = _leer_seleccion_output(setup, "GetSelectedCasesForOutput")
            combos_sel = _leer_seleccion_output(setup, "GetSelectedCombosForOutput")

            if debug_log:
                _escribir_debug_tabla(
                    debug_dir,
                    "seleccion_resultados",
                    resultado_completo=("seleccion",),
                    headings_raw="seleccion",
                    data_raw=[],
                    motivo=(
                        f"Solicitados casos={casos_lista}, combos={combos_lista}; "
                        f"Seleccionados casos={casos_sel}, combos={combos_sel}; "
                        f"Errores: {', '.join(errores_sel) if errores_sel else 'ninguno'}"
                    ),
                )

        except Exception as exc:  # pragma: no cover - depende de COM
            # No detenemos la extracción; log y seguimos con la selección actual de ETABS
            if debug_log:
                _escribir_debug_tabla(
                    debug_dir,
                    "seleccion_resultados_error",
                    resultado_completo=("seleccion_error",),
                    headings_raw="seleccion_error",
                    data_raw=[],
                    motivo=f"Fallo seleccionando casos/combos: {exc}",
                )

    for nombre_tabla in tablas_a_extraer:
        tabla_destino = _resolver_tabla(nombre_tabla, tablas_disponibles)
        try:
            db_tables.SetAllTablesSelected(False)
            db_tables.SetTableSelected(tabla_destino.key)
        except Exception:
            # Algunos wrappers COM no requieren seleccionar previamente las tablas.
            # Si falla la selección, seguimos e intentamos la lectura directa.
            pass

        ultimo_type_error: Exception | None = None
        leido = False
        ret = None
        resultado_completo: tuple | list | None = None
        ultimo_type_error = None
        leido = False
        for args in [
            (tabla_destino.key,),
            (tabla_destino.key, ""),
            (tabla_destino.key, "All"),
            (tabla_destino.key, "", ""),
            (tabla_destino.key, "All", ""),
        ]:
            try:
                resultado = db_tables.GetTableForDisplayArray(*args)
                if not isinstance(resultado, (tuple, list)) or len(resultado) < 3:
                    raise RuntimeError(
                        f"ETABS devolvió una estructura inesperada para '{tabla_destino.nombre}': {resultado}"
                    )
                resultado_completo = tuple(resultado)
                ret = resultado_completo[0]
                leido = True
                break
            except TypeError as exc:
                ultimo_type_error = exc
                continue
            except Exception as exc:  # pragma: no cover - interacción directa con COM
                raise RuntimeError(
                    f"No se pudo leer la tabla '{tabla_destino.nombre}' desde ETABS: {exc}"
                ) from exc

        if not leido or resultado_completo is None:
            raise RuntimeError(
                f"No se pudo leer la tabla '{tabla_destino.nombre}' desde ETABS: {ultimo_type_error}"
            )

        def _ret_ok(valor):
            if valor in (0, "0", None, "", False):
                return True
            if isinstance(valor, (tuple, list)) and len(valor) == 0:
                return True
            try:
                return int(valor) == 0
            except Exception:
                return False

        if not _ret_ok(ret):
            raise RuntimeError(
                f"ETABS devolvió el código {ret} al leer la tabla '{tabla_destino.nombre}'."
            )

        # Seleccionamos encabezados y datos, tratando estructuras alternativas.
        headings_raw = resultado_completo[1] if len(resultado_completo) > 1 else ()
        data_raw = resultado_completo[2] if len(resultado_completo) > 2 else ()

        def _es_iterable(x):
            return hasattr(x, "__iter__") and not isinstance(x, (str, bytes))

        # Heurística: algunas versiones entregan (ret, ncols, headings, nrows, data, version)
        if (not _es_iterable(headings_raw)) or isinstance(headings_raw, (int, float)):
            if len(resultado_completo) > 2 and _es_iterable(resultado_completo[2]):
                headings_raw = resultado_completo[2]
                if len(resultado_completo) > 4 and _es_iterable(resultado_completo[4]):
                    data_raw = resultado_completo[4]
                elif len(resultado_completo) > 3 and _es_iterable(resultado_completo[3]):
                    data_raw = resultado_completo[3]
        else:
            if len(resultado_completo) > 4 and not _es_iterable(data_raw) and _es_iterable(resultado_completo[4]):
                data_raw = resultado_completo[4]

        # Normalizamos encabezados y datos, incluso si vienen como escalares.
        try:
            headings = list(headings_raw)
        except Exception:
            headings = [str(headings_raw)]

        if not headings:
            raise RuntimeError(
                f"La tabla '{tabla_destino.nombre}' no devolvió encabezados desde ETABS."
            )

        try:
            data_lista = list(data_raw)
        except Exception:
            data_lista = [data_raw]

        # Si llega un único valor pero hay varias columnas, replicamos vacío; si llega una sola columna y un valor, construimos fila única.
        if len(headings) > 1 and len(data_lista) == 1:
            # Un solo valor para varias columnas: rellenar el resto con cadena vacía.
            data_lista = data_lista + [""] * (len(headings) - 1)
        elif len(headings) == 1 and len(data_lista) > 1 and len(data_lista) % len(headings) != 0:
            # Una columna, datos sueltos: mantenerlos como filas de una columna.
            pass

        if len(data_lista) % len(headings) != 0:
            raise RuntimeError(
                "El tamaño de los datos no coincide con las columnas recibidas "
                f"para la tabla '{tabla_destino.nombre}'."
            )

        filas = len(data_lista) // len(headings)
        registros = [data_lista[i * len(headings):(i + 1) * len(headings)] for i in range(filas)]
        df = pd.DataFrame(registros, columns=headings)

        # Filtrado adicional por casos/combos después de leer los datos (por si ETABS ignoró la selección)
        if (casos_lista or combos_lista) and "OutputCase" in df.columns:
            permitidos = set(casos_lista + combos_lista)
            permitidos_norm = {str(p).strip().lower() for p in permitidos}
            antes = len(df)
            df = df[df["OutputCase"].astype(str).str.strip().str.lower().isin(permitidos_norm)]
            if debug_log:
                _escribir_debug_tabla(
                    debug_dir,
                    f"filtro_casos_{tabla_destino.nombre}",
                    resultado_completo=("filtro",),
                    headings_raw=list(df.columns),
                    data_raw=[],
                    motivo=(
                        f"Filtrado OutputCase por permitidos={sorted(permitidos)} "
                        f"filas antes={antes} despues={len(df)}"
                    ),
                )

        if df.empty:
            _escribir_debug_tabla(
                debug_dir,
                tabla_destino.nombre,
                resultado_completo,
                headings_raw,
                data_raw,
                motivo="DataFrame vacío",
            )
            raise RuntimeError(
                f"La tabla '{tabla_destino.nombre}' no devolvió filas (0 registros). "
                "Revisa que haya resultados en ETABS para esta tabla."
            )

        if debug_log:
            _escribir_debug_tabla(
                debug_dir,
                tabla_destino.nombre,
                resultado_completo,
                headings_raw,
                data_raw,
                motivo="Debug solicitado (primeras filas)",
                df_preview=df.head(5),
            )
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
        ret_all, tablas_all = None, []
    else:
        if tablas_all:
            return ret_all, tablas_all

    ret_available, tablas_available = _normalizar_get_available_tables(
        db_tables.GetAvailableTables()
    )
    return ret_available, tablas_available


def _escribir_debug_tabla(
    carpeta: Path,
    nombre_tabla: str,
    resultado_completo,
    headings_raw,
    data_raw,
    motivo: str,
    df_preview=None,
):
    """Guarda en disco la estructura cruda devuelta por ETABS para diagnóstico."""

    try:
        carpeta.mkdir(parents=True, exist_ok=True)
        ruta = carpeta / "debug_tablas_etabs.log"
        with ruta.open("a", encoding="utf-8") as f:
            f.write(f"=== {nombre_tabla} ===\n")
            f.write(f"Motivo: {motivo}\n")
            f.write(f"Resultado bruto: {repr(resultado_completo)}\n")
            f.write(f"Headings raw: {repr(headings_raw)}\n")
            f.write(f"Data raw (primeros 50): {repr(list(data_raw)[:50])}\n")
            if df_preview is not None:
                try:
                    f.write("Preview DataFrame (hasta 5 filas):\n")
                    f.write(df_preview.to_string(index=False) + "\n")
                except Exception:
                    pass
            f.write("\n")
    except Exception:
        # Evitamos romper el flujo si el log falla.
        pass


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
        exito=bool(tablas),
        detalle=detalle,
    )


def _intentar_get_available_tables(
    db_tables, intento: int | None = None
) -> tuple[list[TablaDisponible], PasoDiagnostico]:
    """Ejecuta GetAvailableTables con normalización defensiva."""

    try:
        ret, tablas = _normalizar_get_available_tables(db_tables.GetAvailableTables())
    except Exception as exc:  # pragma: no cover - interacción COM
        return [], PasoDiagnostico(
            metodo="GetAvailableTables",
            exito=False,
            detalle=f"estructura inesperada o excepción: {exc} (intento={intento})",
        )

    detalle = f"ret={ret}, tablas={len(tablas)}, intento={intento}"
    return tablas, PasoDiagnostico(
        metodo="GetAvailableTables",
        exito=bool(tablas),
        detalle=detalle,
    )


def _normalizar_get_all_tables(resultado) -> tuple[int, list[TablaDisponible]]:
    """Convierte la respuesta de ``GetAllTables`` en datos homogéneos."""

    if not isinstance(resultado, (tuple, list)):  # pragma: no cover - defensivo
        try:
            resultado = tuple(resultado)
        except Exception:
            raise TypeError(
                "GetAllTables devolvió un tipo inesperado. Se esperaba un tuple o secuencia."
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

    if not isinstance(resultado, (tuple, list)):  # pragma: no cover - defensive
        try:
            resultado = tuple(resultado)
        except Exception:
            raise TypeError(
                "GetAvailableTables devolvió un tipo inesperado. Se esperaba un tuple o secuencia."
            )

    if len(resultado) >= 5:
        ret, table_keys, table_names, import_types, vacias = resultado[:5]
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

    if ret is None:
        ret = -1

    if (table_names is None or len(table_names) == 0) and table_keys:
        table_names = table_keys

    claves = list(table_keys or [])
    nombres = list(table_names or [])

    total = max(len(claves), len(nombres))

    if len(claves) < total:
        claves.extend([None] * (total - len(claves)))
    if len(nombres) < total:
        nombres.extend(claves[len(nombres) :])

    tipos = list(import_types or [])
    estados_vacios = list(vacias or [])

    if len(tipos) < total:
        tipos.extend([None] * (total - len(tipos)))
    if len(estados_vacios) < total:
        estados_vacios.extend([None] * (total - len(estados_vacios)))

    tablas = [
        TablaDisponible(
            key=str(key),
            nombre=str(nombre),
            import_type=tipo if import_types is not None else None,
            esta_vacia=estado if vacias is not None else None,
        )
        for key, nombre, tipo, estado in zip(claves, nombres, tipos, estados_vacios)
    ]

    return ret, tablas
