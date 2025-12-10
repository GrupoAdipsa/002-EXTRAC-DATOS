import sys
import types
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

try:  # pragma: no cover - depende del entorno del ejecutor
    import pandas as pd  # type: ignore
except ImportError:  # pragma: no cover - fallback ligero para entornos sin pandas
    class _MiniDataFrame:
        def __init__(self, data, columns):
            self._columns = list(columns)
            self._data = [list(row) for row in data]

        @property
        def columns(self):
            return self._columns

        @property
        def empty(self):
            return len(self._data) == 0

        def to_csv(self, ruta, index=False, encoding="utf-8", sep=","):
            ruta = Path(ruta)
            with ruta.open("w", encoding=encoding) as f:
                f.write(sep.join(self._columns) + "\n")
                for fila in self._data:
                    f.write(sep.join(map(str, fila)) + "\n")

        def head(self, n):
            return _MiniDataFrame(self._data[:n], self._columns)

        def __len__(self):
            return len(self._data)

    _pandas_stub = types.ModuleType("pandas")
    _pandas_stub.DataFrame = _MiniDataFrame
    sys.modules["pandas"] = _pandas_stub
    pd = _pandas_stub

from tablas_etabs import (
    TablaDisponible,
    _normalizar_get_all_tables,
    _normalizar_get_available_tables,
    _obtener_tablas_disponibles,
    _resolver_tabla,
    diagnosticar_listado_tablas,
    extraer_tablas_etabs,
)


class FakeDatabaseTables:
    def __init__(self):
        self.selected = []
        self.calls = []

    def GetAllTables(self):
        return (
            0,
            ["TABLE_A", "TABLE_B"],
            ["Table A", "Table B"],
            [1, 2],
            [False, True],
        )

    def GetAvailableTables(self):
        return (
            0,
            ["TABLE_A", "TABLE_B"],
            ["Table A", "Table B"],
            [1, 2],
            [False, True],
        )

    def SetAllTablesSelected(self, value):
        self.calls.append(("SetAllTablesSelected", value))

    def SetTableSelected(self, key):
        self.selected.append(key)
        self.calls.append(("SetTableSelected", key))

    def GetTableForDisplayArray(self, key):
        headings = ["Col1", "Col2"]
        data = [f"{key}-r1c1", f"{key}-r1c2", f"{key}-r2c1", f"{key}-r2c2"]
        return 0, headings, data, 1, None, None, None


class FakeSapModel:
    def __init__(self):
        self.DatabaseTables = FakeDatabaseTables()


class FakeDatabaseTablesWithRetries:
    def __init__(self, respuestas_available):
        self._respuestas = list(respuestas_available)
        self._indice = 0

    def GetAllTables(self):
        raise RuntimeError("GetAllTables no está disponible")

    def GetAvailableTables(self):
        respuesta = self._respuestas[min(self._indice, len(self._respuestas) - 1)]
        self._indice += 1
        return respuesta


class NormalizacionTests(unittest.TestCase):
    def test_normalizar_get_all_tables_preserva_metadatos(self):
        ret, tablas = _normalizar_get_all_tables(
            (0, ["K1"], ["Nombre 1"], [3], [False])
        )
        self.assertEqual(ret, 0)
        self.assertEqual(tablas[0].key, "K1")
        self.assertEqual(tablas[0].nombre, "Nombre 1")
        self.assertEqual(tablas[0].import_type, 3)
        self.assertFalse(tablas[0].esta_vacia)

    def test_normalizar_get_available_tables_cuando_faltan_nombres(self):
        ret, tablas = _normalizar_get_available_tables((0, ["K1"], [], None, None))
        self.assertEqual(ret, 0)
        self.assertEqual(tablas[0].key, "K1")
        self.assertEqual(tablas[0].nombre, "K1")

    def test_resolver_tabla_coinicide_por_nombre_o_key(self):
        disponibles = [
            TablaDisponible("K1", "Nombre Uno"),
            TablaDisponible("K2", "Otro"),
        ]
        self.assertEqual(_resolver_tabla("Nombre Uno", disponibles).key, "K1")
        self.assertEqual(_resolver_tabla("k2", disponibles).key, "K2")

    def test_obtener_tablas_reintenta_available(self):
        db_tables = FakeDatabaseTablesWithRetries(
            [
                (0, [], [], None, None),  # primer intento vacío
                (0, None, None, None, None),  # segundo intento sin datos
                (0, ["KX"], ["Nombre X"], None, None),  # éxito en el tercero
            ]
        )
        ret, tablas = _obtener_tablas_disponibles(types.SimpleNamespace(DatabaseTables=db_tables))

        self.assertEqual(ret, 0)
        self.assertEqual(len(tablas), 1)
        self.assertEqual(tablas[0].key, "KX")


class ExtraccionTests(unittest.TestCase):
    def test_extraer_tablas_devuelve_dataframes(self):
        sap_model = FakeSapModel()
        resultado = extraer_tablas_etabs(
            sap_model,
            tablas=["Table A"],
            carpeta_destino=None,
        )
        self.assertIn("Table A", resultado)
        df = resultado["Table A"]
        self.assertIsInstance(df, pd.DataFrame)
        self.assertEqual(list(df.columns), ["Col1", "Col2"])
        self.assertEqual(len(df), 2)

    def test_extraer_tablas_exporta_a_csv_y_txt(self):
        sap_model = FakeSapModel()
        with TemporaryDirectory() as tmpdir:
            destino = Path(tmpdir)
            resultado = extraer_tablas_etabs(
                sap_model,
                tablas=["Table B"],
                carpeta_destino=destino,
                formatos=["csv", "txt"],
            )
            self.assertEqual(len(resultado["Table B"]), 2)
            csv_path = destino / "table_b.csv"
            txt_path = destino / "table_b.txt"
            self.assertTrue(csv_path.exists())
            self.assertTrue(txt_path.exists())

    def test_diagnosticar_listado_tablas_reintenta_y_loguea(self):
        db_tables = FakeDatabaseTablesWithRetries(
            [
                (0, [], [], None, None),
                (0, ["K3"], ["Nombre 3"], None, None),
            ]
        )
        sap_model = types.SimpleNamespace(DatabaseTables=db_tables)

        tablas, pasos = diagnosticar_listado_tablas(sap_model)

        self.assertEqual([t.key for t in tablas], ["K3"])
        # 1 intento de GetAllTables fallido + 2 intentos de GetAvailableTables
        self.assertEqual(len(pasos), 3)
        self.assertTrue(any("intento 1" in p.detalle for p in pasos))
        self.assertTrue(any("intento 2" in p.detalle for p in pasos))


if __name__ == "__main__":
    unittest.main()
