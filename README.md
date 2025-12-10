# Extracción de tablas ETABS

Este repositorio reúne utilidades simples para conectarse a un modelo abierto de ETABS mediante COM y descargar datos tabulares a `pandas` o a disco.

## Componentes principales

- `conectar_etabs.py` expone helpers mínimos para obtener un `SapModel` abierto (`obtener_sapmodel_etabs`) y configurar unidades (`establecer_units_etabs`).
- `tablas_etabs.py` define la lista de tablas predeterminadas (`DEFAULT_TABLES`) y la función que realiza la lectura real (`extraer_tablas_etabs`).
- `extraer_tablas.py` reexporta lo anterior para que otros scripts puedan importar únicamente este módulo.

## ¿Cómo funciona `extraer_tablas_etabs`?

1. Requiere un `SapModel` activo; normalmente lo obtienes con `obtener_sapmodel_etabs()`.
2. Por defecto lee las tres tablas más comunes para sismo:
   - `Story Forces`
   - `Diaphragm Accelerations`
   - `Story Drifts`
3. Si necesitas otras tablas, pasa los nombres exactos en español o inglés tal como aparecen en ETABS mediante el argumento `tablas`.
4. Si proporcionas `carpeta_destino`, cada tabla se puede guardar como CSV o TXT (tabulado) en esa carpeta; si lo omites, la función devuelve un diccionario con `pandas.DataFrame` en memoria.

Para saber qué tablas tienes disponibles en tu modelo, usa `listar_tablas_etabs` y, si quieres, aplícale un filtro.

### ¿Cómo sé que funcionó?

- Si `prueba_extraccion.py` o tu script no muestran errores y aparece el mensaje de ✅ con el número de tablas, la extracción se ejecutó.
- Al inicio `prueba_extraccion.py` imprime qué método de la API respondió (`GetAllTables` o `GetAvailableTables`); si alguno falla verás un ⚠️/❌ con el detalle devuelto por ETABS.
- Cuando no envías `--salida`, el script imprime las primeras filas de cada tabla; si una tabla aparece como "(sin filas devueltas)", significa que ETABS no tenía datos para esa tabla en ese momento (por ejemplo, porque faltan resultados calculados).
- Si envías `--salida`, revisa los archivos generados en la carpeta indicada para confirmar que se escribieron correctamente.

## Ejemplo rápido

```python
from conectar_etabs import obtener_sapmodel_etabs
from extraer_tablas import (
    DEFAULT_TABLES,
    extraer_tablas_etabs,
    listar_tablas_etabs,
)

sap_model = obtener_sapmodel_etabs()
if sap_model:
    # 1) Averiguar las tablas disponibles (opcional)
    todas = listar_tablas_etabs(sap_model)
    print("Tablas disponibles:", todas)

    # También puedes aplicar un filtro parcial, por ejemplo "drift"
    drifts = listar_tablas_etabs(sap_model, filtro="drift")
    print("Solo las que contienen 'drift':", drifts)

    # 2) Usar las tablas por defecto y obtener los DataFrames en memoria
    tablas = extraer_tablas_etabs(sap_model)

    # 3) Especificar tablas personalizadas y exportarlas a CSV y/o TXT
    tablas_personalizadas = ["Joint Reactions", "Modal Participating Mass Ratios"]
    extraer_tablas_etabs(
        sap_model,
        tablas=tablas_personalizadas,
        carpeta_destino="./salidas_etabs",
        formatos=["csv", "txt"],
    )
```

## Flujo típico

1. Abre el modelo en ETABS y deja la aplicación ejecutándose.
2. Ejecuta tu script de Python y llama a `obtener_sapmodel_etabs()` para conectarte.
3. Opcionalmente, ajusta las unidades con `establecer_units_etabs()`.
4. Llama a `extraer_tablas_etabs()` (directamente o importado desde `extraer_tablas.py`).
5. Usa los `DataFrame` en memoria o revisa los archivos generados en disco en el formato que elijas.

## Notas

- Si ETABS devuelve un código de error, la función elevará una excepción para que sepas exactamente qué tabla falló.
- Los nombres de tabla son sensibles y deben coincidir con los que ETABS expone en su base de datos.
- La selección de tablas es **programática**: no existe un menú desplegable ni interfaz gráfica. Debes decidir en tu script qué nombres usar, por ejemplo filtrando con `listar_tablas_etabs` y construyendo la lista que le pasarás a `extraer_tablas_etabs`.
- Si envías `carpeta_destino`, puedes elegir `formatos="csv"`, `formatos="txt"` o una lista con ambos para guardar las dos versiones. Si omites la carpeta, las tablas se devuelven solo como `DataFrame`.

### Vista gráfica opcional

Si prefieres marcar casillas y elegir la carpeta de salida con un selector, puedes abrir una ventana básica con `lanzar_gui_etabs()`. La interfaz intentará conectarse a ETABS automáticamente; si ya tienes un `SapModel` lo puedes pasar como argumento, pero no es obligatorio. Desde la ventana puedes consultar las tablas disponibles, marcarlas, escoger CSV/TXT y la carpeta de exportación. Es un prototipo que puedes ejecutar en paralelo y luego integrar al flujo principal.

### Prueba rápida en consola

Para confirmar que la conexión y la extracción funcionan en tu entorno, ejecuta el script `prueba_extraccion.py`. El programa se conectará a ETABS, listará las primeras tablas disponibles y tratará de leerlas, mostrando en consola las primeras filas o la carpeta donde se guardaron los archivos.

```bash
python prueba_extraccion.py \
  --tablas "Story Forces" "Diaphragm Accelerations" "Story Drifts" \
  --salida ./salidas_etabs \
  --formatos csv txt
```

Si omites `--salida`, la prueba imprimirá un resumen en pantalla sin generar archivos.

### ¿Cómo elegir desde un listado sin menú?

A continuación un ejemplo sencillo para que el usuario seleccione por consola a partir del catálogo disponible. No crea un menú gráfico, pero sirve para escoger interactívamente y luego pasar los nombres exactos al extractor:

```python
sap_model = obtener_sapmodel_etabs()
catalogo = listar_tablas_etabs(sap_model)

for indice, nombre in enumerate(catalogo, start=1):
    print(f"{indice}. {nombre}")

opcion = int(input("Ingresa el número de la tabla que quieres extraer: "))
tabla_seleccionada = [catalogo[opcion - 1]]

extraer_tablas_etabs(
    sap_model,
    tablas=tabla_seleccionada,
    carpeta_destino="./salidas",
    formatos=["csv", "txt"],
)
```
