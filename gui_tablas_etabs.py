"""Interfaz gráfica sencilla para elegir y exportar tablas de ETABS.

La ventana permite:
- Cargar el catálogo de tablas disponibles desde el `SapModel` activo.
- Seleccionar con casillas de verificación las tablas a extraer.
- Elegir el formato de exportación (CSV, TXT o ambos).
- Definir la carpeta de destino donde se guardarán los archivos.

Está pensada como un prototipo para trabajar en paralelo a los scripts
existentes y luego integrar la lógica según se necesite. Requiere que
ya exista un `SapModel` conectado (por ejemplo, con
:func:`conectar_etabs.obtener_sapmodel_etabs`).
"""

from __future__ import annotations

from pathlib import Path
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    TOP,
    Button,
    Checkbutton,
    Entry,
    Frame,
    IntVar,
    Label,
    Listbox,
    Radiobutton,
    StringVar,
    Tk,
    Toplevel,
    ttk,
    filedialog,
    messagebox,
)
from tkinter.scrolledtext import ScrolledText

from conectar_etabs import obtener_sapmodel_etabs
from tablas_etabs import (
    DEFAULT_TABLES,
    diagnosticar_listado_tablas,
    extraer_tablas_etabs,
)

__all__ = ["lanzar_gui_etabs"]


class _ExtractorGUI:
    """Ventana principal para seleccionar tablas y exportar."""

    def __init__(self, sap_model):
        self.sap_model = sap_model
        self.root = Tk()
        self.root.title("Exportar tablas ETABS")
        self.root.geometry("520x520")

        self._tabla_vars: list[tuple[str, IntVar]] = []
        self._tabla_nodes: dict[str, str] = {}
        self._formato_csv = IntVar(value=1)
        self._formato_txt = IntVar(value=0)
        self._ruta_destino = StringVar(value=str(Path.cwd() / "salidas_etabs"))
        self._casos: list[str] = []
        self._combos: list[str] = []
        self._estado_cc = StringVar(value="")

        self._construir_layout()
        self._cargar_tablas_disponibles()
        self._cargar_casos_combos()

    def _construir_layout(self) -> None:
        info = Label(
            self.root,
            text=(
                "Selecciona las tablas que quieres extraer. "
                "Los formatos elegidos se guardarán en la carpeta indicada."
            ),
            wraplength=480,
            justify=LEFT,
        )
        info.pack(pady=10, padx=10, anchor="w")

        frame_tablas = Frame(self.root)
        frame_tablas.pack(fill=BOTH, expand=True, padx=10)

        self.lista_tablas = ttk.Treeview(
            frame_tablas,
            columns=("tabla",),
            show="tree",
            selectmode="none",
            height=16,
        )
        self.lista_tablas.column("#0", width=240, anchor="w")
        self.lista_tablas.heading("#0", text="Categoría")
        self.lista_tablas.column("tabla", anchor="w", width=260)
        self.lista_tablas.heading("tabla", text="Tabla")
        self.lista_tablas.pack(side=LEFT, fill=BOTH, expand=True)

        scroll = ttk.Scrollbar(frame_tablas, orient="vertical", command=self.lista_tablas.yview)
        scroll.pack(side=RIGHT, fill="y")
        self.lista_tablas.configure(yscrollcommand=scroll.set)

        botones_tablas = Frame(self.root)
        botones_tablas.pack(pady=5, padx=10, anchor="e")
        Button(botones_tablas, text="Conectar a ETABS", command=self._reconectar).pack(side=LEFT, padx=5)
        Button(botones_tablas, text="Actualizar", command=self._cargar_tablas_disponibles).pack(side=LEFT, padx=5)
        Button(botones_tablas, text="Seleccionar predeterminadas", command=self._seleccionar_default).pack(side=LEFT, padx=5)

        frame_cc = Frame(self.root)
        frame_cc.pack(fill="x", padx=10, pady=5)
        Label(frame_cc, text="Casos de carga:").pack(side=LEFT, anchor="n")
        self.lista_casos = Listbox(frame_cc, selectmode="extended", height=6, exportselection=False, width=30)
        self.lista_casos.pack(side=LEFT, padx=5, fill="both", expand=True)
        Label(frame_cc, text="Combinaciones:").pack(side=LEFT, anchor="n")
        self.lista_combos = Listbox(frame_cc, selectmode="extended", height=6, exportselection=False, width=30)
        self.lista_combos.pack(side=LEFT, padx=5, fill="both", expand=True)
        Button(frame_cc, text="Refrescar casos/combos", command=self._cargar_casos_combos).pack(side=LEFT, padx=5)
        Label(self.root, textvariable=self._estado_cc, anchor="w", fg="gray").pack(fill="x", padx=10)
        frame_cc_opts = Frame(self.root)
        frame_cc_opts.pack(fill="x", padx=10, pady=2)
        self._usar_casos = IntVar(value=1)
        self._usar_combos = IntVar(value=1)
        Checkbutton(frame_cc_opts, text="Aplicar casos seleccionados", variable=self._usar_casos).pack(side=LEFT, padx=5)
        Checkbutton(frame_cc_opts, text="Aplicar combinaciones seleccionadas", variable=self._usar_combos).pack(side=LEFT, padx=5)

        frame_formatos = Frame(self.root)
        frame_formatos.pack(fill="x", padx=10, pady=10)
        Label(frame_formatos, text="Formatos de exportación:").pack(side=LEFT)
        Checkbutton(frame_formatos, text="CSV", variable=self._formato_csv).pack(side=LEFT, padx=5)
        Checkbutton(frame_formatos, text="TXT", variable=self._formato_txt).pack(side=LEFT, padx=5)

        frame_ruta = Frame(self.root)
        frame_ruta.pack(fill="x", padx=10)
        Label(frame_ruta, text="Carpeta destino:").pack(side=LEFT)
        self.entry_ruta = Entry(frame_ruta, textvariable=self._ruta_destino, width=40)
        self.entry_ruta.pack(side=LEFT, padx=5, fill="x", expand=True)
        Button(frame_ruta, text="Examinar", command=self._seleccionar_carpeta).pack(side=LEFT)

        frame_acciones = Frame(self.root)
        frame_acciones.pack(side=TOP, pady=15)
        Button(frame_acciones, text="Grヴficas", command=self._abrir_graficas).pack(side=LEFT, padx=10)
        Button(frame_acciones, text="Previsualizar", command=self._previsualizar).pack(side=LEFT, padx=10)
        Button(frame_acciones, text="Extraer", command=self._extraer).pack(side=LEFT, padx=10)
        Button(frame_acciones, text="Cerrar", command=self.root.destroy).pack(side=LEFT, padx=10)

    def _asegurar_sapmodel(self) -> bool:
        if self.sap_model:
            return True

        self.sap_model = obtener_sapmodel_etabs()
        if not self.sap_model:
            messagebox.showwarning(
                "SapModel no disponible",
                "No se pudo conectar a ETABS. Abre un modelo y vuelve a intentarlo.",
            )
            return False

        return True

    def _cargar_tablas_disponibles(self) -> None:
        for item in self.lista_tablas.get_children():
            self.lista_tablas.delete(item)
        self._tabla_vars.clear()
        self._tabla_nodes.clear()

        if not self._asegurar_sapmodel():
            return

        try:
            tablas_disponibles, pasos = diagnosticar_listado_tablas(self.sap_model)
            tablas = [t.nombre for t in tablas_disponibles]
        except Exception as exc:  # pragma: no cover - interacción COM
            pasos = getattr(exc, "pasos", [])
            log = "\n".join(f"- {p.metodo}: {p.detalle}" for p in pasos) or "Sin detalle"
            messagebox.showerror(
                "Error",
                f"No se pudo obtener el listado de tablas:\n{exc}\n\nDiagnóstico:\n{log}",
            )
            return

        if not tablas:
            messagebox.showwarning(
                "Sin resultados",
                "ETABS no devolvió tablas disponibles para este modelo.",
            )
            return

        grupos = self._agrupar_tablas(tablas)

        for categoria, tablas_cat in grupos.items():
            parent_id = self.lista_tablas.insert("", END, text=categoria, open=False)

            for tabla in tablas_cat:
                var = IntVar(value=0)
                self._tabla_vars.append((tabla, var))
                item_id = self.lista_tablas.insert(
                    parent_id, END, text="", values=(f"[ ] {tabla}",), tags=(tabla,)
                )
                self._tabla_nodes[tabla] = item_id
                self.lista_tablas.tag_bind(
                    tabla, "<ButtonRelease-1>", lambda e, v=var, t=tabla: self._toggle_tabla(v, t)
                )

    def _reconectar(self) -> None:
        self.sap_model = None
        if self._asegurar_sapmodel():
            self._cargar_tablas_disponibles()
            self._cargar_casos_combos()

    def _toggle_tabla(self, var: IntVar, tabla: str) -> None:
        var.set(0 if var.get() else 1)
        texto = f"[{'x' if var.get() else ' '}] {tabla}"
        item_id = self._buscar_item(tabla)
        if item_id:
            self.lista_tablas.item(item_id, values=(texto,))

    def _buscar_item(self, tabla: str):
        return self._tabla_nodes.get(tabla)

    def _seleccionar_default(self) -> None:
        tablas_default = set(nombre.lower() for nombre in DEFAULT_TABLES)
        for tabla, var in self._tabla_vars:
            marcado = tabla.lower() in tablas_default
            var.set(1 if marcado else 0)
            texto = f"[{'x' if marcado else ' '}] {tabla}"
            item_id = self._buscar_item(tabla)
            if item_id:
                self.lista_tablas.item(item_id, values=(texto,))

    def _seleccionar_carpeta(self) -> None:
        carpeta = filedialog.askdirectory()
        if carpeta:
            self._ruta_destino.set(carpeta)

    def _agrupar_tablas(self, tablas: list[str]) -> dict[str, list[str]]:
        """Agrupa las tablas por categoría derivada (heurística simple)."""
        grupos: dict[str, list[str]] = {}

        for nombre in sorted(tablas, key=str.lower):
            categoria = self._extraer_categoria(nombre)
            grupos.setdefault(categoria, []).append(nombre)

        # Ordenar tablas dentro de cada categoría
        for cat in list(grupos):
            grupos[cat] = sorted(grupos[cat], key=str.lower)

        return dict(sorted(grupos.items(), key=lambda kv: kv[0].lower()))

    @staticmethod
    def _extraer_categoria(nombre_tabla: str) -> str:
        """Deriva una categoría simple a partir del nombre de tabla."""
        if ":" in nombre_tabla:
            return nombre_tabla.split(":", 1)[0].strip()

        tokens = nombre_tabla.split()
        if len(tokens) >= 2:
            return " ".join(tokens[:2])
        if tokens:
            return tokens[0]
        return "Otras"

    def _extraer(self) -> None:
        if not self._asegurar_sapmodel():
            return

        seleccionadas = [tabla for tabla, var in self._tabla_vars if var.get()]
        if not seleccionadas:
            usar_default = messagebox.askyesno(
                "Sin tablas seleccionadas",
                "No elegiste ninguna tabla. ¿Usar las predeterminadas?",
            )
            if usar_default:
                seleccionadas = DEFAULT_TABLES
            else:
                return

        formatos = []
        if self._formato_csv.get():
            formatos.append("csv")
        if self._formato_txt.get():
            formatos.append("txt")

        if not formatos:
            messagebox.showwarning("Formato requerido", "Selecciona al menos un formato: CSV y/o TXT.")
            return

        ruta = self._ruta_destino.get().strip()
        if not ruta:
            messagebox.showwarning("Carpeta requerida", "Indica una carpeta de destino para guardar los archivos.")
            return

        casos, combos = self._obtener_casos_combos_seleccionados()
        if not self._usar_casos.get():
            casos = []
        if not self._usar_combos.get():
            combos = []
        if not self._usar_casos.get():
            casos = []
        if not self._usar_combos.get():
            combos = []

        try:
            resultado = extraer_tablas_etabs(
                self.sap_model,
                tablas=seleccionadas,
                carpeta_destino=Path(ruta),
                formatos=formatos,
                casos=casos,
                combinaciones=combos,
                debug_log=True,
            )
        except Exception as exc:  # pragma: no cover - interacción COM
            messagebox.showerror("Error", f"No se pudieron exportar las tablas:\n{exc}")
            return

        messagebox.showinfo(
            "Listo",
            f"Se exportaron {len(resultado)} tablas en formato {', '.join(formatos)}\n"
            f"Carpeta: {Path(ruta).resolve()}",
        )

    def _previsualizar(self) -> None:
        """Lee las tablas seleccionadas y muestra un preview en memoria."""
        if not self._asegurar_sapmodel():
            return

        seleccionadas = [tabla for tabla, var in self._tabla_vars if var.get()]
        if not seleccionadas:
            usar_default = messagebox.askyesno(
                "Sin tablas seleccionadas",
                "No elegiste ninguna tabla. ¿Usar las predeterminadas?",
            )
            if usar_default:
                seleccionadas = DEFAULT_TABLES
            else:
                return

        casos, combos = self._obtener_casos_combos_seleccionados()

        try:
            resultado = extraer_tablas_etabs(
                self.sap_model,
                tablas=seleccionadas,
                carpeta_destino=None,
                casos=casos,
                combinaciones=combos,
                debug_log=True,
            )
        except Exception as exc:
            messagebox.showerror("Error", f"No se pudieron previsualizar las tablas:\n{exc}")
            return

        if not resultado:
            messagebox.showinfo("Sin datos", "No se devolvieron datos para las tablas seleccionadas.")
            return

        popup = Tk()
        popup.title("Preview de tablas")
        texto = ScrolledText(popup, width=100, height=35)
        texto.pack(fill=BOTH, expand=True, padx=8, pady=8)

        for nombre, df in resultado.items():
            texto.insert(END, f"=== {nombre} ===\n")
            try:
                vista = df.head(50)
                texto.insert(END, vista.to_string(index=False) + "\n")
                if len(df) > 50:
                    texto.insert(END, f"... ({len(df) - 50} filas más)\n")
                texto.insert(END, "\n")
            except Exception as exc:  # pragma: no cover - defensivo
                texto.insert(END, f"(No se pudo mostrar la tabla: {exc})\n\n")

        Button(popup, text="Cerrar", command=popup.destroy).pack(pady=4)

    def _cargar_casos_combos(self) -> None:
        """Carga los casos de carga y combinaciones disponibles."""
        if not self._asegurar_sapmodel():
            return

        self.lista_casos.delete(0, END)
        self.lista_combos.delete(0, END)
        self._estado_cc.set("")

        ret_casos, nombres_casos = self._leer_lista_etabs(self.sap_model.LoadCases.GetNameList)
        if nombres_casos:
            self.lista_casos.insert(END, *[str(n) for n in nombres_casos])

        ret_combos, nombres_combos = self._leer_lista_etabs(self.sap_model.RespCombo.GetNameList)
        if nombres_combos:
            self.lista_combos.insert(END, *[str(n) for n in nombres_combos])

        self._estado_cc.set(
            f"Casos ret={ret_casos} encontrados={len(nombres_casos)} | "
            f"Combos ret={ret_combos} encontrados={len(nombres_combos)}"
        )

    def _obtener_casos_combos_seleccionados(self) -> tuple[list[str], list[str]]:
        casos = [self.lista_casos.get(i) for i in self.lista_casos.curselection()]
        combos = [self.lista_combos.get(i) for i in self.lista_combos.curselection()]
        return casos, combos

    @staticmethod
    def _leer_lista_etabs(func):
        """Intenta leer listas desde métodos GetNameList de ETABS manejando estructuras variadas."""
        try:
            resultado = func()
        except Exception as exc:
            return f"err:{exc}", []

        # Normalizamos a lista
        try:
            items = list(resultado)
        except Exception:
            items = [resultado]

        # Primer elemento como ret si es numérico, de lo contrario ret=-1 y usamos todo como datos
        if items and isinstance(items[0], (int, float)):
            ret = items[0]
            datos = items[1:]
        else:
            ret = -1
            datos = items

        nombres: list[str] = []

        def _flatten_strings(obj):
            if isinstance(obj, str):
                nombres.append(obj)
            elif isinstance(obj, (list, tuple)):
                for el in obj:
                    _flatten_strings(el)

        for elem in datos:
            _flatten_strings(elem)

        return ret, nombres

    # Seccion de grГЎficas -------------------------------------------------

    def _abrir_graficas(self) -> None:
        """Abre un cuadro simple para elegir la grГЎfica a generar."""
        if not self._asegurar_sapmodel():
            return

        popup = Toplevel(self.root)
        popup.title("GrГЎficas")

        Label(popup, text="Elige la tabla a graficar").pack(anchor="w", padx=10, pady=(10, 5))

        opciones = [
            ("Story Drifts (deriva maxima por piso)", "story_drifts"),
            ("Story Forces", "story_forces"),
            ("Diaphragm Accelerations", "diaphragm_acc"),
            ("Story Accelerations", "story_accel"),
            ("Story Max Over Avg Displacement", "story_max_over_avg_disp"),
            ("Story Max Over Avg Drift", "story_max_over_avg_drift"),
        ]
        self._grafica_opcion = StringVar(value=opciones[0][1])
        for texto, valor in opciones:
            Radiobutton(popup, text=texto, variable=self._grafica_opcion, value=valor).pack(
                anchor="w", padx=10
            )

        Label(popup, text="Direcciones a incluir (coma separada, ej. X,Y) - deja vacio para todas:").pack(anchor="w", padx=10, pady=(6, 0))
        self._grafica_direcciones_lista = StringVar(value="")
        Entry(popup, textvariable=self._grafica_direcciones_lista, width=20).pack(anchor="w", padx=10, pady=(0, 4))

        Label(popup, text="Direccion preferida (opcional, ej. X o Y):").pack(anchor="w", padx=10, pady=(2, 0))
        self._grafica_direccion = StringVar(value="")
        Entry(popup, textvariable=self._grafica_direccion, width=12).pack(anchor="w", padx=10, pady=(0, 6))

        Label(
            popup,
            text=(
                "Usa las selecciones actuales de casos/combos para filtrar.\n"
                "La grГЎfica se abre en una ventana de Matplotlib."
            ),
            fg="gray",
            justify=LEFT,
        ).pack(anchor="w", padx=10, pady=5)

        botones = Frame(popup)
        botones.pack(pady=10)
        Button(botones, text="Graficar", command=lambda: self._graficar_seleccion(popup)).pack(
            side=LEFT, padx=5
        )
        Button(botones, text="Cerrar", command=popup.destroy).pack(side=LEFT, padx=5)

    def _graficar_seleccion(self, popup: Toplevel) -> None:
        opcion = getattr(self, "_grafica_opcion", StringVar(value="")).get()
        if opcion == "story_drifts":
            self._graficar_story_drifts()
        elif opcion in {"story_forces", "diaphragm_acc", "story_accel", "story_max_over_avg_disp", "story_max_over_avg_drift"}:
            self._graficar_generico(opcion)
        else:
            messagebox.showinfo("Pendiente", "Todavia no hay graficas para esa opcion.")
        popup.lift()

    def _graficar_story_drifts(self) -> None:
        """Genera la grГЎfica de derivas mГЎximas por piso usando Matplotlib."""
        try:
            from graficar_tablas_etabs import plot_max_story_drift
        except Exception as exc:
            messagebox.showerror(
                "Matplotlib requerido",
                f"No se pudo importar la funciГіn de grГЎfico: {exc}\n"
                "Instala matplotlib (pip install matplotlib) e intenta de nuevo.",
            )
            return

        if not self._asegurar_sapmodel():
            return

        casos, combos = self._obtener_casos_combos_seleccionados()
        if not self._usar_casos.get():
            casos = []
        if not self._usar_combos.get():
            combos = []

        direcciones = []
        if hasattr(self, "_grafica_direcciones_lista"):
            raw_dirs = (self._grafica_direcciones_lista.get() or "")
            direcciones = [d.strip() for d in raw_dirs.split(",") if d.strip()]

        prefer_dir = (self._grafica_direccion.get() or "").strip() if hasattr(self, "_grafica_direccion") else ""

        try:
            tablas = extraer_tablas_etabs(
                self.sap_model,
                tablas=["Story Drifts"],
                carpeta_destino=None,
                casos=casos,
                combinaciones=combos,
                debug_log=False,
            )
        except Exception as exc:  # pragma: no cover - interacciГіn COM
            messagebox.showerror("Error", f"No se pudo leer la tabla para graficar:\n{exc}")
            return

        try:
            plot_max_story_drift(
                tablas,
                table_name="Story Drifts",
                cases=list(casos) + list(combos),
                directions=direcciones or None,
                prefer_direction=prefer_dir or None,
                title="Maximum Story Drifts",
                xlabel="Drift, Unitless",
                show=True,
                block=False,
            )
        except Exception as exc:
            messagebox.showerror("Error al graficar", f"No se pudo generar la grГЎfica:\n{exc}")
            return

        messagebox.showinfo("Listo", "La grГЎfica se abriГі en la ventana de Matplotlib.")

    def _graficar_generico(self, opcion: str) -> None:
        """Graficador genГ©rico para tablas tipo Story."""
        try:
            from graficar_tablas_etabs import plot_story_columns
        except Exception as exc:
            messagebox.showerror(
                "Matplotlib requerido",
                f"No se pudo importar la funciГіn de grГЎfico: {exc}\n"
                "Instala matplotlib (pip install matplotlib) e intenta de nuevo.",
            )
            return

        if not self._asegurar_sapmodel():
            return

        casos, combos = self._obtener_casos_combos_seleccionados()
        if not self._usar_casos.get():
            casos = []
        if not self._usar_combos.get():
            combos = []

        direcciones = []
        if hasattr(self, "_grafica_direcciones_lista"):
            raw_dirs = (self._grafica_direcciones_lista.get() or "")
            direcciones = [d.strip() for d in raw_dirs.split(",") if d.strip()]

        prefer_dir = (self._grafica_direccion.get() or "").strip() if hasattr(self, "_grafica_direccion") else ""

        config = {
            "story_forces": {
                "tabla": "Story Forces",
                "candidatos": ["P", "V2", "V3", "T", "M2", "M3", "MX", "MY", "VX", "VY"],
                "xlabel": "Fuerza / Momento",
                "titulo": "Story Forces",
            },
            "diaphragm_acc": {
                "tabla": "Diaphragm Accelerations",
                "candidatos": ["Accel", "Acceleration", "U1", "U2", "U3", "UX", "UY", "UZ"],
                "xlabel": "Acceleration",
                "titulo": "Diaphragm Accelerations",
            },
            "story_accel": {
                "tabla": "Story Accelerations",
                "candidatos": ["Accel", "Acceleration", "U1", "U2", "U3", "UX", "UY", "UZ"],
                "xlabel": "Acceleration",
                "titulo": "Story Accelerations",
            },
            "story_max_over_avg_disp": {
                "tabla": "Story Max Over Avg Displacement",
                "candidatos": ["Max Displacement", "Avg Displacement", "Ratio"],
                "xlabel": "Displacement",
                "titulo": "Story Max Over Avg Displacement",
            },
            "story_max_over_avg_drift": {
                "tabla": "Story Max Over Avg Drift",
                "candidatos": ["Max Drift", "Avg Drift", "Ratio"],
                "xlabel": "Drift",
                "titulo": "Story Max Over Avg Drift",
            },
        }.get(opcion)

        if not config:
            messagebox.showinfo("Pendiente", "No hay grГЎfica configurada para esta opciГіn.")
            return

        try:
            tablas = extraer_tablas_etabs(
                self.sap_model,
                tablas=[config["tabla"]],
                carpeta_destino=None,
                casos=casos,
                combinaciones=combos,
                debug_log=False,
            )
        except Exception as exc:  # pragma: no cover - interacciГіn COM
            messagebox.showerror("Error", f"No se pudo leer la tabla para graficar:\n{exc}")
            return

        try:
            plot_story_columns(
                tablas,
                table_name=config["tabla"],
                value_candidates=config["candidatos"],
                cases=list(casos) + list(combos),
                directions=direcciones or None,
                prefer_direction=prefer_dir or None,
                title=config["titulo"],
                xlabel=config["xlabel"],
                show=True,
                block=False,
            )
        except Exception as exc:
            messagebox.showerror("Error al graficar", f"No se pudo generar la grГЎfica:\n{exc}")
            return

        messagebox.showinfo("Listo", "La grГЎfica se abriГі en la ventana de Matplotlib.")

    def run(self) -> None:
        self.root.mainloop()


def lanzar_gui_etabs(sap_model) -> None:
    """Inicializa y muestra la interfaz gráfica para exportar tablas.

    Parameters
    ----------
    sap_model : SapModel
        Objeto ya conectado a ETABS. Se utilizará para consultar el listado de
        tablas disponibles y ejecutar la exportación.
    """

    app = _ExtractorGUI(sap_model)
    app.run()


if __name__ == "__main__":  # pragma: no cover - uso interactivo
    sap = obtener_sapmodel_etabs()
    if sap:
        lanzar_gui_etabs(sap)
    else:
        messagebox.showwarning(
            "SapModel no disponible",
            "No se pudo obtener SapModel. Abre un modelo en ETABS y vuelve a intentarlo.",
        )
