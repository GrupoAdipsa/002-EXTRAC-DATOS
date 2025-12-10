import comtypes.client as cc


def get_etabs_model():
    try:
        etabs = cc.GetActiveObject("CSI.ETABS.API.ETABSObject")
    except Exception as e:
        raise RuntimeError("No se encontró una instancia abierta de ETABS.") from e
    return etabs.SapModel


def listar_tablas(sap_model):
    """
    Usa la firma documentada:
    int GetAvailableTables(ref int NumberTables,
                           ref string[] TableKey,
                           ref string[] TableName,
                           ref int[] ImportType)
    """

    NumberTables = 0
    TableKey = []
    TableName = []
    ImportType = []

    ret, NumberTables, TableKey, TableName, ImportType = (
        sap_model.DatabaseTables.GetAvailableTables(
            NumberTables,
            TableKey,
            TableName,
            ImportType,
        )
    )

    if ret != 0:
        raise RuntimeError(f"Error ETABS al obtener tablas. Código: {ret}")

    print(f"\nSe encontraron {NumberTables} tablas:\n")
    for key, name in zip(TableKey, TableName):
        print(f"- {key} -> {name}")

    return TableKey, TableName


def main():
    sap_model = get_etabs_model()
    print("Conectado a ETABS.\n")

    table_keys, table_names = listar_tablas(sap_model)


if __name__ == "__main__":
    main()
