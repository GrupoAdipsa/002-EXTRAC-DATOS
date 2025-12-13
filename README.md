# Extraccion de tablas ETABS

Kit de utilidades para conectarse a un modelo abierto de ETABS via COM, listar tablas disponibles y traerlas a pandas o a disco. Incluye helpers para graficar y una GUI rapida con Tkinter.

## Componentes principales

- `conectar_etabs.py`: obtiene `SapModel` activo (`obtener_sapmodel_etabs`) y permite fijar unidades (`establecer_units_etabs`). Maneja errores de import de `comtypes` en 3.13.
- `tablas_etabs.py`: `DEFAULT_TABLES`, `extraer_tablas_etabs` (exporta a CSV/TXT, filtra por casos y combinaciones, escribe log de debug opcional), `listar_tablas_etabs` y `diagnosticar_listado_tablas` para saber que metodo de la API respondio.
- `extraer_tablas.py`: reexpone lo anterior y agrega `lanzar_gui_etabs`, `plot_max_story_drift` y `plot_story_columns`.
- `graficar_tablas_etabs.py`: helpers de Matplotlib para graficar derivas maximas por piso y otras tablas tipo Story.
- `gui_tablas_etabs.py`: interfaz Tkinter para elegir tablas, casos/combos, formatos y carpeta de salida, previsualizar, exportar y generar graficos rapidos.
- `prueba_extraccion.py`: script de consola para probar la conexion y la extraccion.

## Flujo rapido en codigo

```python
from conectar_etabs import obtener_sapmodel_etabs
from extraer_tablas import DEFAULT_TABLES, extraer_tablas_etabs, listar_tablas_etabs

sap = obtener_sapmodel_etabs()
if not sap:
    raise SystemExit("Abre un modelo en ETABS antes de ejecutar este script")

print("Tablas disponibles con 'drift':", listar_tablas_etabs(sap, filtro="drift"))

tablas = extraer_tablas_etabs(
    sap,
    tablas=DEFAULT_TABLES,
    carpeta_destino="./salidas_etabs",
    formatos=["csv", "txt"],
    casos=["EX", "EY"],           # opcional: filtrar casos de carga
    combinaciones=["COMB_Sismo"], # opcional: filtrar combinaciones
    debug_log=True,               # escribe debug_tablas_etabs.log con estructuras crudas
)

print(tablas["Story Drifts"].head())
```

## Diagnostico y listado

- `diagnosticar_listado_tablas(sap_model)` intenta `GetAllTables` y varias formas de `GetAvailableTables`, devolviendo tanto la lista como el log de pasos para revisar en consola si algo falla.
- `extraer_tablas_etabs(..., debug_log=True)` guarda `debug_tablas_etabs.log` en la carpeta de salida (o en el cwd) con la respuesta completa de ETABS cuando una tabla llega vacia o se activa el flag.

## Graficos rapidos

```python
from extraer_tablas import plot_max_story_drift, plot_story_columns

# Derivas maximas por piso
plot_max_story_drift(tablas, cases=["EX", "EY"], directions=["X", "Y"], save_path="salidas_etabs/derivas.png", show=False)

# Otras tablas tipo Story (usa candidatos para detectar columnas numericas)
plot_story_columns(
    tablas,
    table_name="Story Forces",
    value_candidates=["V2", "V3", "M2", "M3"],
    cases=["EX"],
    directions=["X"],
    title="Story Forces - ejemplo",
    save_path="salidas_etabs/story_forces.png",
    show=False,
)
```

## GUI de exportacion y graficas

- Ejecuta `python gui_tablas_etabs.py` (o importa `lanzar_gui_etabs(obtener_sapmodel_etabs())`) con ETABS abierto.
- La ventana agrupa las tablas por tipo, permite marcar las predeterminadas, elegir carpeta y formatos, seleccionar casos y combinaciones y previsualizar resultados sin escribir a disco.
- Incluye un panel para graficar rapidamente `Story Drifts`, `Story Forces`, `Diaphragm Accelerations`, `Story Accelerations`, `Story Max Over Avg Displacement` y `Story Max Over Avg Drift` usando Matplotlib.
- La GUI activa `debug_log=True` para dejar rastro en `debug_tablas_etabs.log`.

## Script de prueba en consola

Para una verificacion rapida sin GUI:

```bash
python prueba_extraccion.py ^
  --tablas "Story Forces" "Diaphragm Accelerations" "Story Drifts" ^
  --salida ./salidas_etabs ^
  --formatos csv txt
```

Si omites `--salida`, solo imprime los primeros registros en consola.

## Notas y requisitos

- Necesitas ETABS abierto con un modelo cargado; la conexion se realiza via `comtypes`.
- Si `comtypes` falla en Python 3.13, instala una version reciente o usa Python 3.12.x (el modulo lo reporta con un mensaje claro).
- Las tablas devueltas se filtran por `OutputCase` cuando envias `casos` o `combinaciones`; si ETABS no entrega filas se registra en el log y se lanza una excepcion.
