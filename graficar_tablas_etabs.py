"""Funciones de graficado para tablas extraidas de ETABS.

Incluye un helper para graficar derivas maximas por piso a partir de la
tabla "Story Drifts" o similar. Esta enfocado en un uso simple: recibe
un DataFrame (o el diccionario devuelto por `extraer_tablas_etabs`),
filtra casos/direcciones, agrupa por piso y traza una serie por caso.
"""

from __future__ import annotations

from pathlib import Path
from typing import Iterable, Mapping, Sequence

import matplotlib.pyplot as plt
import pandas as pd

__all__ = ["plot_max_story_drift", "plot_story_columns", "plot_joint_drifts"]


# Utilidades internas -----------------------------------------------------

def _resolver_df(tabla, table_name: str | None) -> pd.DataFrame:
    if isinstance(tabla, pd.DataFrame):
        return tabla

    if isinstance(tabla, Mapping):
        if table_name is None and len(tabla) == 1:
            return next(iter(tabla.values()))

        nombre = table_name or "Story Drifts"
        for k, v in tabla.items():
            if str(k).lower() == nombre.lower():
                return v
        raise KeyError(
            f"No se encontro la tabla '{nombre}' dentro del diccionario recibido."
        )

    raise TypeError(
        "`tabla` debe ser un DataFrame o un mapeo nombre->DataFrame como el que "
        "devuelve `extraer_tablas_etabs`."
    )


def _clave_col(nombre: str) -> str:
    # Normaliza claves eliminando separadores comunes para coincidencias flexibles
    return (
        nombre.strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace(",", "")
        .replace(":", "")
        .replace(";", "")
    )


def _normalizar_columnas(df: pd.DataFrame) -> dict[str, str]:
    claves = {
        "story": {"story", "nivel", "piso", "storyname", "storylabel"},
        "joint": {"joint", "point", "label", "name", "node", "uniquename"},
        "case": {"outputcase", "case", "loadcase", "combo", "combination"},
        "direction": {"direction", "dir", "eje", "orientacion"},
        "drift": {"drift", "maxdrift", "maximumdrift", "storydrift", "deriva"},
    }

    res: dict[str, str] = {}
    for col in df.columns:
        key = _clave_col(col)
        for canon, posibles in claves.items():
            if key in posibles:
                res[canon] = col
    return res


def _ordenar_stories(nombres: Sequence[str], orden_preferido: Sequence[str] | None) -> list[str]:
    if orden_preferido:
        preferidos = [str(o) for o in orden_preferido]
        resto = [n for n in nombres if n not in preferidos]
        return preferidos + _ordenar_stories(resto, None)

    def clave(st: str):
        base_keys = {"base", "basement", "foundation"}
        if st.strip().lower() in base_keys:
            return (-1, -1)
        import re
        m = re.search(r"(\d+)", st)
        if m:
            return (0, int(m.group(1)))
        return (1, st.lower())

    return sorted(nombres, key=clave)


def _serie_label(row: pd.Series, cols: dict[str, str]) -> str:
    case_col = cols.get("case")
    dir_col = cols.get("direction")
    case_val = str(row[case_col]) if case_col else "Serie"
    if dir_col and dir_col in row:
        dir_val = str(row[dir_col]).strip()
        if dir_val:
            return f"{case_val} {dir_val}"
    return case_val


# Funcion principal -------------------------------------------------------

def plot_max_story_drift(
    tabla,
    table_name: str | None = "Story Drifts",
    cases: Iterable[str] | None = None,
    directions: Iterable[str] | None = None,
    prefer_direction: str | None = None,
    story_order: Sequence[str] | None = None,
    title: str = "Maximum Story Drifts",
    xlabel: str = "Drift, Unitless",
    drift_limits: Iterable[float] | None = None,
    colors: Mapping[str, str] | None = None,
    markers: Sequence[str] | None = None,
    grayscale: bool = False,
    show: bool = False,
    block: bool = True,
    save_path: str | Path | None = None,
    interactive_controls: bool = True,
    ax=None,
):
    """Genera una grafica de derivas maximas por piso.

    Parameters
    ----------
    tabla : pandas.DataFrame o Mapping[str, DataFrame]
        DataFrame de derivas o el diccionario devuelto por `extraer_tablas_etabs`.
    table_name : str, default "Story Drifts"
        Nombre de la tabla a buscar dentro del diccionario (si aplica).
    cases : Iterable[str], optional
        Filtra solo estos casos/combinaciones (coincidencia exacta por texto).
    directions : Iterable[str], optional
        Filtra solo estas direcciones si la columna existe.
    prefer_direction : str, optional
        Si no se pasan `directions`, intenta filtrar a esta direcciГіn principal (coincidencia por texto).
    story_order : Sequence[str], optional
        Orden explicito de pisos. Si no se indica se usa Base y luego Story1, Story2, ...
    title : str
        Titulo de la grafica.
    xlabel : str
        Texto para el eje X.
    colors : Mapping[str, str], optional
        Diccionario serie->color para forzar colores.
    markers : Sequence[str], optional
        Secuencia de marcadores para las series; se ciclan si hay mas series.
    grayscale : bool
        Si True usa una paleta en escala de grises y marcador negro; por defecto va a color.
    show : bool
        Si es True llama a plt.show() al final.
    block : bool
        Se pasa a plt.show(block=block) cuando show=True. Ajusta a False si llamas desde una GUI y no quieres bloquear.
    drift_limits : Iterable[float], optional
        Valores de deriva de referencia para dibujar lineas verticales (se trazan ±valor).
    save_path : str or Path, optional
        Ruta para guardar la figura (png, svg, etc). No se guarda si es None.
    ax : matplotlib Axes, optional
        Eje existente donde dibujar. Si no se proporciona se crea uno nuevo.
    """

    df = _resolver_df(tabla, table_name)
    cols = _normalizar_columnas(df)

    if "story" not in cols or "drift" not in cols:
        raise ValueError(
            "No se encontraron columnas de Story y/o Drift en la tabla proporcionada."
        )

    df_trab = df.copy()

    drift_col = cols["drift"]
    df_trab[drift_col] = pd.to_numeric(df_trab[drift_col], errors="coerce")
    df_trab = df_trab.dropna(subset=[drift_col])

    if cases and "case" in cols:
        permitidos = {str(c).strip().lower() for c in cases}
        df_trab = df_trab[df_trab[cols["case"]].astype(str).str.strip().str.lower().isin(permitidos)]

    dir_col = cols.get("direction")
    if dir_col and directions:
        permitidos_dir = {str(d).strip().lower() for d in directions}
        df_trab = df_trab[df_trab[dir_col].astype(str).str.strip().str.lower().isin(permitidos_dir)]
    elif dir_col and prefer_direction:
        prefer = str(prefer_direction).strip().lower()
        df_trab = df_trab[df_trab[dir_col].astype(str).str.strip().str.lower() == prefer]

    if df_trab.empty:
        raise ValueError("La tabla filtrada quedo vacia; revisa casos/direcciones o datos de entrada.")

    # Serie = caso/direccion (segun columnas disponibles)
    df_trab["__serie__"] = df_trab.apply(lambda r: _serie_label(r, cols), axis=1)

    historias = df_trab[cols["story"]].astype(str).unique().tolist()
    orden = _ordenar_stories(historias, story_order)
    y_ticks = list(range(len(orden)))

    if ax is None:
        _, ax = plt.subplots(figsize=(5, 8))

    if grayscale:
        palette_default = ["#000000", "#555555", "#888888", "#AAAAAA", "#CCCCCC"]
        marker_default = markers or ["o", "s", "^", "D", "v", "<", ">"]
    else:
        palette_default = ["#e53935", "#1e88e5", "#43a047", "#fdd835", "#8e24aa", "#00897b"]
        marker_default = markers or ["o", "s", "^", "D", "v", "<", ">"]

    for idx, (serie, bloque) in enumerate(df_trab.groupby("__serie__")):
        serie_colors = colors or {}
        color = serie_colors.get(serie, palette_default[idx % len(palette_default)])
        marker = marker_default[idx % len(marker_default)]

        serie_por_story = (
            bloque
            .groupby(cols["story"])[drift_col]
            .apply(lambda s: s.abs().max())
            .reindex(orden)
        )

        ax.plot(
            serie_por_story.values,
            y_ticks,
            marker=marker,
            color=color,
            label=serie,
        )

    ax.set_yticks(y_ticks)
    ax.set_yticklabels(orden)
    ax.set_xlabel(xlabel)
    ax.set_title(title)
    ax.grid(True, which="both", linestyle="--", alpha=0.5)
    ax.legend(loc="best")
    ax.set_ylim(-0.5, len(orden) - 0.5)

    if drift_limits:
        etiqueta_usada = False
        for val in drift_limits:
            try:
                v = abs(float(val))
            except Exception:
                continue
            label = f"Limite {v}" if not etiqueta_usada else None
            etiqueta_usada = etiqueta_usada or label is not None
            ax.axvline(x=v, color="#666666", linestyle=":", linewidth=1.4, label=label)
            ax.axvline(x=-v, color="#666666", linestyle=":", linewidth=1.0)

    plt.tight_layout()

    if interactive_controls and ax.figure is not None:
        _abrir_panel_estilo(ax.get_lines(), ax.figure)

    if save_path:
        Path(save_path).parent.mkdir(parents=True, exist_ok=True)
        ax.figure.savefig(save_path, dpi=200)

    if show:
        plt.show(block=block)

    return ax


def _abrir_panel_estilo(lines, fig=None):
    """Abre un panel Tk para ajustar grosor de linea y tamaño de marcador."""
    try:
        import tkinter as tk
    except Exception:
        return

    if not lines:
        return

    lw0 = lines[0].get_linewidth()
    ms0 = lines[0].get_markersize()

    root = tk.Toplevel()
    root.title("Estilo de graficas")
    root.geometry("260x140")

    tk.Label(root, text="Grosor de linea").pack(pady=(10, 2))
    s_lw = tk.Scale(root, from_=0.5, to=6.0, resolution=0.1, orient="horizontal")
    s_lw.set(lw0)
    s_lw.pack(fill="x", padx=12)

    tk.Label(root, text="Tamaño de marcador").pack(pady=(6, 2))
    s_ms = tk.Scale(root, from_=2, to=18, resolution=0.5, orient="horizontal")
    s_ms.set(ms0)
    s_ms.pack(fill="x", padx=12)

    def _apply(_evt=None):
        for ln in lines:
            ln.set_linewidth(s_lw.get())
            ln.set_markersize(s_ms.get())
        if fig:
            fig.canvas.draw_idle()

    s_lw.configure(command=_apply)
    s_ms.configure(command=_apply)

    try:
        root.attributes("-topmost", True)
        root.after(400, lambda: root.attributes("-topmost", False))
    except Exception:
        pass


def plot_story_columns(
    tabla,
    table_name: str,
    value_candidates: Sequence[str],
    cases: Iterable[str] | None = None,
    directions: Iterable[str] | None = None,
    prefer_direction: str | None = None,
    story_order: Sequence[str] | None = None,
    title: str | None = None,
    xlabel: str | None = None,
    colors: Mapping[str, str] | None = None,
    markers: Sequence[str] | None = None,
    grayscale: bool = False,
    show: bool = False,
    block: bool = True,
    save_path: str | Path | None = None,
    interactive_controls: bool = True,
    ax=None,
):
    """Grafica una o varias columnas numericas contra Story."""

    df = _resolver_df(tabla, table_name)
    cols = _normalizar_columnas(df)

    if "story" not in cols:
        raise ValueError("No se encontro la columna de Story en la tabla proporcionada.")

    df_trab = df.copy()

    dir_col = cols.get("direction")
    case_col = cols.get("case")

    if cases and case_col:
        permitidos = {str(c).strip().lower() for c in cases}
        df_trab = df_trab[df_trab[case_col].astype(str).str.strip().str.lower().isin(permitidos)]

    if directions and dir_col:
        permitidos_dir = {str(d).strip().lower() for d in directions}
        df_trab = df_trab[df_trab[dir_col].astype(str).str.strip().str.lower().isin(permitidos_dir)]
    elif dir_col and prefer_direction:
        prefer = str(prefer_direction).strip().lower()
        df_trab = df_trab[df_trab[dir_col].astype(str).str.strip().str.lower() == prefer]

    if df_trab.empty:
        raise ValueError("La tabla filtrada quedo vacia; revisa casos/direcciones o datos de entrada.")

    # elegir columnas a graficar segun candidatos presentes
    cand_norm = {_clave_col(c): c for c in value_candidates}
    cols_presentes: list[str] = []
    for col in df_trab.columns:
        key = _clave_col(col)
        if key in cand_norm:
            cols_presentes.append(col)

    if not cols_presentes:
        raise ValueError(
            f"No se encontraron columnas graficables entre los candidatos: {', '.join(value_candidates)}"
        )

    df_trab["__serie_base__"] = df_trab.apply(lambda r: _serie_label(r, cols), axis=1)

    historias = df_trab[cols["story"]].astype(str).unique().tolist()
    orden = _ordenar_stories(historias, story_order)
    y_ticks = list(range(len(orden)))

    if ax is None:
        _, ax = plt.subplots(figsize=(5, 8))

    if grayscale:
        palette_default = ["#000000", "#555555", "#888888", "#AAAAAA", "#CCCCCC"]
        marker_default = markers or ["o", "s", "^", "D", "v", "<", ">"]
    else:
        palette_default = ["#e53935", "#1e88e5", "#43a047", "#fdd835", "#8e24aa", "#00897b"]
        marker_default = markers or ["o", "s", "^", "D", "v", "<", ">"]

    series_labels: list[tuple[str, str]] = []  # (columna, serie_base)
    for col in cols_presentes:
        for serie_base in df_trab["__serie_base__"].unique():
            series_labels.append((col, serie_base))

    for idx, (col, serie_base) in enumerate(series_labels):
        parte = df_trab[df_trab["__serie_base__"] == serie_base]
        if parte.empty:
            continue

        serie_por_story = (
            parte.groupby(cols["story"])[col]
            .apply(lambda s: pd.to_numeric(s, errors="coerce").abs().max())
            .reindex(orden)
        )

        color = (colors or {}).get(f"{serie_base}:{col}", palette_default[idx % len(palette_default)])
        marker = marker_default[idx % len(marker_default)]

        label = serie_base if len(cols_presentes) == 1 else f"{serie_base} - {col}"
        ax.plot(
            serie_por_story.values,
            y_ticks,
            marker=marker,
            color=color,
            label=label,
        )

    ax.set_yticks(y_ticks)
    ax.set_yticklabels(orden)
    ax.set_xlabel(xlabel or "Valor")
    ax.set_title(title or table_name)
    ax.grid(True, which="both", linestyle="--", alpha=0.5)
    ax.legend(loc="best")
    ax.set_ylim(-0.5, len(orden) - 0.5)

    plt.tight_layout()

    if interactive_controls and ax.figure is not None:
        _abrir_panel_estilo(ax.get_lines(), ax.figure)

    if save_path:
        Path(save_path).parent.mkdir(parents=True, exist_ok=True)
        ax.figure.savefig(save_path, dpi=200)

    if show:
        plt.show(block=block)

    return ax


def plot_joint_drifts(
    tabla,
    table_name: str = "Joint Drifts",
    joints: Iterable[str] | None = None,
    directions: Iterable[str] | None = None,
    prefer_direction: str | None = None,
    value_candidates: Iterable[str] | None = None,
    cases: Iterable[str] | None = None,
    title: str = "Joint Drifts",
    xlabel: str = "Drift, Unitless",
    drift_limits: Iterable[float] | None = None,
    grayscale: bool = False,
    show: bool = False,
    block: bool = True,
    save_path: str | Path | None = None,
    interactive_controls: bool = True,
    ax=None,
):
    """Grafica derivas por nodo (Joint Drifts) filtrando por joints/direcciones/casos."""

    df = _resolver_df(tabla, table_name)
    cols = _normalizar_columnas(df)

    joint_col = None
    story_col = None

    # Preferir Label para joint, luego UniqueName
    for col in df.columns:
        key = _clave_col(col)
        if key in {"label", "storylabel"}:
            joint_col = col
            break
    if joint_col is None:
        for col in df.columns:
            key = _clave_col(col)
            if key in {"joint", "point", "uniquename", "name", "node"}:
                joint_col = col
                break

    # Detectar Story
    for col in df.columns:
        if "story" in _clave_col(col):
            story_col = col
            break

    if not joint_col or not story_col:
        raise ValueError("No se encontraron columnas de Story y/o Joint en la tabla proporcionada.")

    # Detectar columnas de drift candidatas (DriftX, DriftY, Drift, etc.)
    drift_cols = [c for c in df.columns if "drift" in _clave_col(c)]
    if value_candidates:
        cand_norm = {_clave_col(c): c for c in value_candidates}
        drift_cols = [c for c in drift_cols if _clave_col(c) in cand_norm]

    if directions:
        dirs_norm = {str(d).strip().lower() for d in directions}
        drift_cols = [
            c
            for c in drift_cols
            if any(d in _clave_col(c) for d in dirs_norm)
        ]
    elif prefer_direction:
        pref = str(prefer_direction).strip().lower()
        drift_cols = [c for c in drift_cols if pref in _clave_col(c)] or drift_cols

    if not drift_cols:
        raise ValueError("No se encontraron columnas de drift (Drift, DriftX, DriftY, etc.) en la tabla proporcionada.")

    df_trab = df.copy()
    for col in drift_cols:
        df_trab[col] = pd.to_numeric(df_trab[col], errors="coerce")
    df_trab = df_trab.dropna(subset=drift_cols, how="all")

    if joints:
        permitidos = {str(j).strip().lower() for j in joints}
        df_trab = df_trab[df_trab[joint_col].astype(str).str.strip().str.lower().isin(permitidos)]

    if cases and cols.get("case"):
        permitidos = {str(c).strip().lower() for c in cases}
        df_trab = df_trab[df_trab[cols["case"]].astype(str).str.strip().str.lower().isin(permitidos)]

    if df_trab.empty:
        raise ValueError("La tabla filtrada quedo vacia; revisa joints/direcciones/casos o datos de entrada.")

    # Expandir a formato largo por drift para etiquetar serie como Caso-Direccion-Joint
    registros = []
    for _, fila in df_trab.iterrows():
        for col in drift_cols:
            val = fila[col]
            if pd.isna(val):
                continue
            dir_tag = "X" if "driftx" in _clave_col(col) else "Y" if "drifty" in _clave_col(col) else col
            case_val = fila[cols["case"]] if cols.get("case") in fila else ""
            serie = f"{case_val}-{dir_tag}-{fila[joint_col]}"
            reg = fila.to_dict()
            reg["__serie__"] = serie
            reg["__drift_val__"] = val
            reg["__drift_dir__"] = dir_tag
            registros.append(reg)

    df_trab = pd.DataFrame(registros)

    # Orden por Story en eje vertical
    stories = df_trab[story_col].astype(str).unique().tolist()
    stories_orden = _ordenar_stories(stories, None)
    y_ticks = list(range(len(stories_orden)))

    fig = None
    if ax is None:
        fig, ax = plt.subplots(figsize=(7, max(6, len(stories_orden) * 0.35 + 2)))
    else:
        fig = ax.figure

    if grayscale:
        palette_default = ["#000000", "#555555", "#888888", "#AAAAAA", "#CCCCCC"]
        marker_default = ["o", "s", "^", "D", "v", "<", ">"]
    else:
        palette_default = ["#e53935", "#1e88e5", "#43a047", "#fdd835", "#8e24aa", "#00897b"]
        marker_default = ["o", "s", "^", "D", "v", "<", ">"]

    for idx, (serie, grupo) in enumerate(df_trab.groupby("__serie__")):
        serie_por_story = (
            grupo.groupby(story_col)["__drift_val__"]
            .apply(lambda s: s.abs().max())
            .reindex(stories_orden)
        )

        color = palette_default[idx % len(palette_default)]
        marker = marker_default[idx % len(marker_default)]

        ax.plot(
            serie_por_story.values,
            y_ticks,
            marker=marker,
            color=color,
            label=serie,
        )

    # Promedio por direccion (X / Y)
    if "__drift_dir__" in df_trab.columns:
        for dir_tag, color in [("X", "#555555"), ("Y", "#999999")]:
            sub = df_trab[df_trab["__drift_dir__"] == dir_tag]
            if sub.empty:
                continue
            prom = (
                sub.groupby(story_col)["__drift_val__"]
                .apply(lambda s: s.abs().mean())
                .reindex(stories_orden)
            )
            ax.plot(
                prom.values,
                y_ticks,
                linestyle="--",
                marker="",
                color=color,
                label=f"Promedio {dir_tag}",
                linewidth=2,
            )
            p84 = (
                sub.groupby(story_col)["__drift_val__"]
                .apply(lambda s: s.abs().quantile(0.84))
                .reindex(stories_orden)
            )
            ax.plot(
                p84.values,
                y_ticks,
                linestyle=":",
                marker="",
                color=color,
                label=f"P84 {dir_tag}",
                linewidth=2,
            )

    ax.set_yticks(y_ticks)
    ax.set_yticklabels(stories_orden)
    ax.set_xlabel(xlabel)
    ax.set_title(title)
    ax.grid(True, which="both", linestyle="--", alpha=0.5)
    ax.legend(loc="best")
    ax.set_ylim(-0.5, len(stories_orden) - 0.5)

    plt.tight_layout()

    if interactive_controls and fig is not None:
        _abrir_panel_estilo(ax.get_lines(), fig)

    if save_path:
        Path(save_path).parent.mkdir(parents=True, exist_ok=True)
        ax.figure.savefig(save_path, dpi=200)

    if show:
        plt.show(block=block)

    return ax
