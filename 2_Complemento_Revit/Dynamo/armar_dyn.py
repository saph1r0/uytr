"""Arma Cargar_Parametros_NEXO_NO.dyn (Dynamo 4, PythonNet3) con el script cargar_parametros_nexo_no.py."""
import json
import sys
import uuid
from pathlib import Path

AQUI = Path(__file__).parent
codigo = (AQUI / "cargar_parametros_nexo_no.py").read_text(encoding="utf-8")
datos = sys.argv[1] if len(sys.argv) > 1 else r"D:\NEXO OBRA (NO)\3_Modelo_y_parametros\parametros_revit_REP0-MOD2-R02.json"
salida = Path(sys.argv[2]) if len(sys.argv) > 2 else AQUI / "Cargar_Parametros_NEXO_NO.dyn"
u = lambda: uuid.uuid4().hex


def entrada(nid, oid, nombre, valor):
    nodo = {"ConcreteType": "CoreNodeModels.Input.StringInput, CoreNodeModels", "SerializedWidth": 0.0, "SerializedHeight": 0.0,
            "Id": nid, "NodeType": "StringInputNode", "Inputs": [],
            "Outputs": [{"Id": oid, "Name": "", "Description": "String", "UsingDefaultValue": False, "Level": 2, "UseLevels": False,
                         "KeepListStructure": False}],
            "Replication": "Disabled", "Description": "Crea una cadena.", "InputValue": valor}
    io = {"Id": nid, "Name": nombre, "Type": "string", "Type2": "string", "Value": valor, "Description": "Crea una cadena."}
    return nodo, io


n1, o1, n2, o2, npy, ipy0, ipy1, opy = (u() for _ in range(8))
nodo1, io1 = entrada(n1, o1, "Archivo de datos (parametros_revit_<modelo>.json)", datos)
nodo2, io2 = entrada(n2, o2, "Fecha de corte (dd/mm/aaaa, vacío = hoy)", "")
python = {"ConcreteType": "PythonNodeModels.PythonNode, PythonNodeModels", "Code": codigo, "Engine": "PythonNet3",
          "VariableInputPorts": True, "Id": npy, "NodeType": "PythonScriptNode",
          "Inputs": [{"Id": ipy0, "Name": "IN[0]", "Description": "Archivo de datos", "UsingDefaultValue": False, "Level": 2,
                      "UseLevels": False, "KeepListStructure": False},
                     {"Id": ipy1, "Name": "IN[1]", "Description": "Fecha de corte", "UsingDefaultValue": False, "Level": 2,
                      "UseLevels": False, "KeepListStructure": False}],
          "Outputs": [{"Id": opy, "Name": "OUT", "Description": "Resultado del script de Python", "UsingDefaultValue": False,
                       "Level": 2, "UseLevels": False, "KeepListStructure": False}],
          "Replication": "Disabled", "Description": "Ejecuta un script de Python incrustado."}
dyn = {
    "Uuid": str(uuid.uuid4()), "IsCustomNode": False,
    "Description": ("NEXO OBRA (NO): crea los 15 parámetros VZ_ de valorización y control en las categorías estructurales, "
                    "los llena con el cronograma y el presupuesto (archivo JSON), respeta las fechas de ejecución y crea "
                    "las vistas VZ - SECTORES, VZ - ESTADO y la tabla VZ - VALORIZACION. Se puede correr varias veces."),
    "Name": "Cargar_Parametros_NEXO_NO", "ElementResolver": {"ResolutionMap": {}},
    "Inputs": [io1, io2],
    "Outputs": [{"Id": npy, "Name": "Reporte", "Type": "unknown", "InitialValue": "", "Description": "Resultado del script."}],
    "Nodes": [nodo1, nodo2, python],
    "Connectors": [{"Start": o1, "End": ipy0, "Id": u(), "IsHidden": "False"},
                   {"Start": o2, "End": ipy1, "Id": u(), "IsHidden": "False"}],
    "Dependencies": [], "NodeLibraryDependencies": [], "EnableLegacyPolyCurveBehavior": None, "Thumbnail": None,
    "GraphDocumentationURL": None, "ExtensionWorkspaceData": [], "Author": "NEXO OBRA (NO)",
    "Linting": {"activeLinter": "Ninguno", "activeLinterId": "7b75fb44-43fd-4631-a878-29f4d5d8399a", "warningCount": 0, "errorCount": 0},
    "Bindings": [],
    "View": {"Dynamo": {"ScaleFactor": 1.0, "HasRunWithoutCrash": False, "IsVisibleInDynamoLibrary": True, "Version": "4.0.0.3277",
                        "RunType": "Manual", "RunPeriod": "1000"},
             "Camera": {"Name": "_Vista preliminar de fondo", "EyeX": -17.0, "EyeY": 24.0, "EyeZ": 50.0, "LookX": 12.0, "LookY": -13.0,
                        "LookZ": -58.0, "UpX": 0.0, "UpY": 1.0, "UpZ": 0.0},
             "ConnectorPins": [],
             "NodeViews": [
                 {"Id": n1, "Name": "Archivo de datos (parametros_revit_<modelo>.json)", "IsSetAsInput": True, "IsSetAsOutput": False,
                  "Excluded": False, "ShowGeometry": True, "X": 0.0, "Y": 0.0},
                 {"Id": n2, "Name": "Fecha de corte (dd/mm/aaaa, vacío = hoy)", "IsSetAsInput": True, "IsSetAsOutput": False,
                  "Excluded": False, "ShowGeometry": True, "X": 0.0, "Y": 160.0},
                 {"Id": npy, "Name": "NEXO OBRA (NO) · Cargar parámetros VZ_", "IsSetAsInput": False, "IsSetAsOutput": True,
                  "Excluded": False, "ShowGeometry": True, "X": 520.0, "Y": 60.0}],
             "Annotations": [], "X": 0.0, "Y": 0.0, "Zoom": 1.0}}
salida.write_text(json.dumps(dyn, ensure_ascii=False, indent=2), encoding="utf-8")
print("Guardado:", salida)
