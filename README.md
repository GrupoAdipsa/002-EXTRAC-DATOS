# Extraccion de tablas ETABS

Utilidades para conectarse a un modelo abierto de ETABS via COM, listar tablas y exportarlas a pandas o disco. Incluye GUI con Tkinter, graficos en Matplotlib (derivas, joint drifts, promedios/p84) y empaquetado a exe/instalador.

## Componentes principales
- `conectar_etabs.py`: obtiene `SapModel` activo y fija unidades; maneja errores de import de `comtypes`.
- `tablas_etabs.py`: `extraer_tablas_etabs` (CSV/TXT, filtros por casos/combos, `debug_log`, `permitir_tablas_vacias`), `listar_tablas_etabs`, `diagnosticar_listado_tablas`, `DEFAULT_TABLES`.
- `graficar_tablas_etabs.py`: `plot_max_story_drift`, `plot_story_columns`, `plot_joint_drifts` (series Caso-Direccion-Label, promedio y percentil 84 por direccion, panel de estilo externo para linea/marker).
- `gui_tablas_etabs.py`: GUI para elegir tablas, casos/combos, formatos, carpeta, previsualizar, exportar y graficar rapido (Story/Joint). Abre panel de estilo aparte al graficar Joint Drifts.
- `extraer_tablas.py`: reexporta helpers y GUI.
- `prueba_extraccion.py`: prueba de consola para listar y extraer.

## Uso basico en codigo
```python
from conectar_etabs import obtener_sapmodel_etabs
from extraer_tablas import extraer_tablas_etabs, listar_tablas_etabs

sap = obtener_sapmodel_etabs()
tablas = extraer_tablas_etabs(
    sap,
    tablas=["Story Drifts", "Joint Drifts"],
    carpeta_destino="./salidas_etabs",
    formatos=["csv", "txt"],
    casos=["Qx", "Qy"],
    combinaciones=[],
    debug_log=True,
    permitir_tablas_vacias=True,
)
print(listar_tablas_etabs(sap, filtro="drift"))
```

## Graficos
- Derivas por piso:
```python
from graficar_tablas_etabs import plot_max_story_drift
plot_max_story_drift(tablas, table_name="Story Drifts", cases=["Qx","Qy"], show=True)
```
- Joint Drifts (series Caso-Direccion-Label, promedio y P84 por direccion):
```python
from graficar_tablas_etabs import plot_joint_drifts
plot_joint_drifts(tablas, joints=None, directions=["X","Y"], show=True)
```
- Story Forces u otras tablas tipo Story:
```python
from graficar_tablas_etabs import plot_story_columns
plot_story_columns(tablas, table_name="Story Forces", value_candidates=["V2","V3","M2","M3"], cases=["Qx"], show=True)
```

## GUI
```bash
python gui_tablas_etabs.py
```
- Selecciona tablas, casos/combos, formatos y carpeta; previsualiza/exporta.
- Graficos integrados: Story Drifts, Joint Drifts, Story Forces, Diaphragm/Story Accelerations, Max Over Avg Displacement/Drift.
- Joint Drifts abre un panel externo para ajustar grosor y tamano de marcadores.

## Script de prueba en consola
```bash
python prueba_extraccion.py ^
  --tablas "Story Forces" "Diaphragm Accelerations" "Story Drifts" ^
  --salida ./salidas_etabs ^
  --formatos csv txt
```
Omitiendo `--salida` solo imprime en consola.

## Empaquetado a exe e instalador
1. Crear exe (desde la carpeta del proyecto):
```bash
python -m venv .venv
.venv\Scripts\activate
pip install pandas matplotlib comtypes pyinstaller
pyinstaller --onefile --noconsole --hidden-import comtypes --hidden-import comtypes.client gui_tablas_etabs.py
```
2. Instalador con Inno Setup (`exportador_etabs.iss` en la raiz):
```
[Setup]
AppName=Exportador ETABS
AppVersion=1.0
DefaultDirName={commonpf}\ExportadorETABS
OutputBaseFilename=ExportadorETABS-Setup
Compression=lzma
SolidCompression=yes
ArchitecturesInstallIn64BitMode=x64compatible

[Files]
Source: "dist\gui_tablas_etabs.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Exportador ETABS"; Filename: "{app}\gui_tablas_etabs.exe"
Name: "{commondesktop}\Exportador ETABS"; Filename: "{app}\gui_tablas_etabs.exe"
```
Compila el `.iss` en Inno Setup; el instalador queda en `Output/`.

## Requisitos
- Windows con ETABS instalado y un modelo abierto (COM).
- Si usas el exe/instalador no necesitas Python; si usas el codigo, instala pandas, matplotlib, comtypes.
