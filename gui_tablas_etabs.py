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
from tkinter import BOTH, END, LEFT, RIGHT, TOP, Button, Checkbutton, Entry, Frame, IntVar, Label, StringVar, Tk, ttk, filedialog, messagebox

from conectar_etabs import obtener_sapmodel_etabs
from tablas_etabs import DEFAULT_TABLES, extraer_tablas_etabs, listar_tablas_etabs

__all__ = ["lanzar_gui_etabs"]


class _ExtractorGUI:
    """Ventana principal para seleccionar tablas y exportar."""

    def __init__(self, sap_model):
        self.sap_model = sap_model
        self.root = Tk()
        self.root.title("Exportar tablas ETABS")
        self.root.geometry("520x520")

        self._tabla_vars: list[tuple[str, IntVar]] = []
        self._formato_csv = IntVar(value=1)
        self._formato_txt = IntVar(value=0)
        self._ruta_destino = StringVar(value=str(Path.cwd() / "salidas_etabs"))

        self._construir_layout()
        self._cargar_tablas_disponibles()

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
            show="",
            selectmode="none",
            height=12,
        )
        self.lista_tablas.column("#0", width=0, stretch=False)
        self.lista_tablas.column("tabla", anchor="w", width=450)
        self.lista_tablas.pack(side=LEFT, fill=BOTH, expand=True)

        scroll = ttk.Scrollbar(frame_tablas, orient="vertical", command=self.lista_tablas.yview)
        scroll.pack(side=RIGHT, fill="y")
        self.lista_tablas.configure(yscrollcommand=scroll.set)

        botones_tablas = Frame(self.root)
        botones_tablas.pack(pady=5, padx=10, anchor="e")
        Button(botones_tablas, text="Conectar a ETABS", command=self._reconectar).pack(side=LEFT, padx=5)
        Button(botones_tablas, text="Actualizar", command=self._cargar_tablas_disponibles).pack(side=LEFT, padx=5)
        Button(botones_tablas, text="Seleccionar predeterminadas", command=self._seleccionar_default).pack(side=LEFT, padx=5)

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

        if not self._asegurar_sapmodel():
            return

        try:
            tablas = listar_tablas_etabs(self.sap_model)
        except Exception as exc:  # pragma: no cover - interacción COM
            messagebox.showerror("Error", f"No se pudo obtener el listado de tablas:\n{exc}")
            return

        if not tablas:
            messagebox.showwarning(
                "Sin resultados",
                "ETABS no devolvió tablas disponibles para este modelo.",
            )
            return

        for tabla in tablas:
            var = IntVar(value=0)
            self._tabla_vars.append((tabla, var))
            self.lista_tablas.insert("", END, values=(f"[ ] {tabla}",), tags=(tabla,))
            self.lista_tablas.tag_bind(tabla, "<ButtonRelease-1>", lambda e, v=var, t=tabla: self._toggle_tabla(v, t))

    def _reconectar(self) -> None:
        self.sap_model = None
        if self._asegurar_sapmodel():
            self._cargar_tablas_disponibles()

    def _toggle_tabla(self, var: IntVar, tabla: str) -> None:
        var.set(0 if var.get() else 1)
        texto = f"[{'x' if var.get() else ' '}] {tabla}"
        item_id = self._buscar_item(tabla)
        if item_id:
            self.lista_tablas.item(item_id, values=(texto,))

    def _buscar_item(self, tabla: str):
        for item in self.lista_tablas.get_children():
            if self.lista_tablas.item(item, "values")[0].endswith(tabla):
                return item
        return None

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

        try:
            resultado = extraer_tablas_etabs(
                self.sap_model,
                tablas=seleccionadas,
                carpeta_destino=Path(ruta),
                formatos=formatos,
            )
        except Exception as exc:  # pragma: no cover - interacción COM
            messagebox.showerror("Error", f"No se pudieron exportar las tablas:\n{exc}")
            return

        messagebox.showinfo(
            "Listo",
            f"Se exportaron {len(resultado)} tablas en formato {', '.join(formatos)}\n"
            f"Carpeta: {Path(ruta).resolve()}",
        )

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
