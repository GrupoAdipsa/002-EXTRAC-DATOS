"""Script de diagnóstico rápido para probar la extracción de tablas ETABS.

Ejecuta conexión, muestra el catálogo disponible y trata de leer las

(tablas predeterminadas o las que le indiques por línea de comandos).
"""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable

from conectar_etabs import obtener_sapmodel_etabs
from extraer_tablas import DEFAULT_TABLES, extraer_tablas_etabs, listar_tablas_etabs


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Prueba rápida para listar y extraer tablas de ETABS, "
            "con mensajes claros en consola."
        )
    )
    parser.add_argument(
        "--tablas",
        nargs="*",
        help="Nombres exactos de las tablas a extraer. Si no se envían se usan las predeterminadas.",
    )
    parser.add_argument(
        "--salida",
        type=Path,
        help=(
            "Carpeta donde se guardarán los archivos CSV/TXT. Si se omite, solo se muestran los primeros registros en consola."
        ),
    )
    parser.add_argument(
        "--formatos",
        nargs="*",
        default=["csv"],
        help="Formatos de exportación cuando se entrega --salida (csv, txt o ambos).",
    )
    parser.add_argument(
        "--max-filas",
        type=int,
        default=5,
        help="Número máximo de filas a imprimir por tabla al mostrar en consola.",
    )
    return parser.parse_args()


def probar_extraccion(
    tablas: Iterable[str] | None = None,
    carpeta_destino: Path | None = None,
    formatos: Iterable[str] | None = None,
    max_filas: int = 5,
) -> None:
    """Ejecuta una prueba básica de extracción y reporta el resultado."""

    print("Intentando conectar con ETABS...")
    sap_model = obtener_sapmodel_etabs()
    if not sap_model:
        print("❌ No se pudo obtener SapModel. Abre un modelo en ETABS y reintenta.")
        return

    print("✅ Conexión establecida. Obteniendo listado de tablas disponibles...")
    try:
        disponibles = listar_tablas_etabs(sap_model)
    except Exception as exc:  # pragma: no cover - interacción COM
        causa = exc.__cause__
        detalle = f" Detalle: {causa}" if causa else ""
        print(f"❌ Error al listar tablas: {exc}.{detalle}")

        # Información extendida para depuración en consola
        import traceback

        print("Traceback completo de la falla al listar tablas:\n")
        print(traceback.format_exc())
        return

    print(f"Se encontraron {len(disponibles)} tablas disponibles.")
    if disponibles:
        print("Primeras 10:")
        for nombre in disponibles[:10]:
            print(f"  - {nombre}")

    tablas_solicitadas = list(tablas or DEFAULT_TABLES)
    print("\nExtrayendo tablas:")
    for nombre in tablas_solicitadas:
        print(f"  - {nombre}")

    try:
        resultado = extraer_tablas_etabs(
            sap_model,
            tablas=tablas_solicitadas,
            carpeta_destino=carpeta_destino,
            formatos=formatos or ["csv"],
        )
    except Exception as exc:  # pragma: no cover - interacción COM
        print(f"❌ Error al extraer: {exc}")
        return

    print(f"\n✅ Se extrajeron {len(resultado)} tablas.")
    if carpeta_destino:
        print(f"Archivos guardados en: {carpeta_destino.resolve()}")
    else:
        for nombre, df in resultado.items():
            print(f"\nTabla: {nombre} (primeras {min(max_filas, len(df))} filas)")
            print(df.head(max_filas))


if __name__ == "__main__":  # pragma: no cover - script de apoyo
    args = _parse_args()
    probar_extraccion(
        tablas=args.tablas,
        carpeta_destino=args.salida,
        formatos=args.formatos,
        max_filas=args.max_filas,
    )
