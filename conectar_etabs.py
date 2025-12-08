import comtypes.client
import pandas as pd

# -- Constantes para tipos de material
# MAT_TYPE_STEEL = 1
MAT_TYPE_CONCRETE = 2
# MAT_TYPE_NODESIGN = 3
# MAT_TYPE_ALUMINIUM = 4
# MAT_TYPE_COLDFORMED = 5
MAT_TYPE_REBAR = 6
# MAT_TYPE_TENDON = 7
# MAT_TYPE_MASONRY = 8

# Units Length
UNITS_LENGTH_IN = 1
UNITS_LENGTH_FT = 2
UNITS_LENGTH_MM = 4
UNITS_LENGTH_CM = 5
UNITS_LENGTH_M = 6

# Units Force
UNITS_FORCE_LB = 1
UNITS_FORCE_KIP = 2
UNITS_FORCE_N = 3
UNITS_FORCE_KN = 4
UNITS_FORCE_KGF = 5
UNITS_FORCE_TONF = 6

# Unidades de Temperatura (eTemp)
UNITS_TEMP_F = 1  # Fahrenheit
UNITS_TEMP_C = 2  # Celsius


def establecer_units_etabs(
    sap_model, unidad_fuerza, unidad_longitud, unidad_temperatura
):
    if sap_model is None:
        print("Error: El objeto SapModel proporcionado no es válido.")
        return False

    print(
        f"\nIntentando establecer unidades: Fuerza={unidad_fuerza}, Longitud={unidad_longitud}, Temperatura={unidad_temperatura}"
    )

    try:
        # La función SetPresentUnits devuelve 0 si tiene éxito.
        # Firma: SetPresentUnits(eForce Force, eLength Length, eTemp Temp)
        ret = sap_model.SetPresentUnits_2(
            unidad_fuerza, unidad_longitud, unidad_temperatura
        )

        if ret == 0:
            print("Unidades establecidas exitosamente en ETABS.")
            # Opcional: Verificar las unidades actuales después de establecerlas
            current_units = (
                sap_model.GetPresentUnits_2()
            )  # Devuelve una tupla (fuerza, longitud, temperatura)
            print(current_units)
            print(
                f"Unidades actuales verificadas: Fuerza={current_units[0]}, Longitud={current_units[1]}, Temperatura={current_units[2]}"
            )
            return True
        else:
            print(
                f"Error al establecer las unidades en ETABS. Código de retorno: {ret}"
            )
            print(
                "Verifica que los códigos de unidades sean válidos para tu versión de ETABS."
            )
            return False

    except comtypes.COMError as e:
        print(f"Error de COM interactuando con ETABS al establecer unidades: {e}")
        return False
    except Exception as e:
        print(f"Ocurrió un error inesperado al establecer unidades: {e}")
        import traceback

        traceback.print_exc()
        return False


def obtener_sapmodel_etabs():
    sap_model = None
    ETABSObject = None

    try:
        print("Intentando conectar con una instancia activa de ETABS...")
        ETABSObject = comtypes.client.GetActiveObject("CSI.ETABS.API.ETABSObject")
        print("Conexión exitosa con ETABSObject.")

        sap_model = ETABSObject.SapModel
        if sap_model is None:
            print("No se pudo obtener el SapModel.")
            return None

        print("SapModel obtenido exitosamente.")
        return sap_model

    except (OSError, comtypes.COMError) as e:
        print("No se pudo encontrar una instancia activa de ETABS o error COM.")
        print(f"Error: {e}")
        print("Asegúrate de que ETABS esté abierto con un modelo.")
        return None
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")
        return None


def close_connection():
    comtypes.CoUninitialize()
    
