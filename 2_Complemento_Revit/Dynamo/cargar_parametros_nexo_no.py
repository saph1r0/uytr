# -*- coding: utf-8 -*-
# NEXO OBRA (NO) · Cargar parámetros VZ_ de valorización y control (Revit 2027 · Dynamo 4 · PythonNet3 / CPython 3)
#
# MODIFICA EL MODELO (no lo guarda):
#   1. Crea 15 parámetros compartidos VZ_ (por ejemplar, grupo «Datos de identidad») en las categorías estructurales.
#   2. Escribe en cada elemento su paquete, sector, actividad, partida, fecha programada, metrados y montos,
#      leídos de parametros_revit_<modelo>.json (lo genera cronograma_nexo.py a partir del cronograma y el presupuesto).
#   3. NUNCA borra las fechas de ejecución (VZ_Fecha_Acero / Encofrado / Concreto): las respeta y con ellas calcula
#      VZ_Estado, VZ_Monto_Valorizado y VZ_Valorizacion_N a la fecha de corte.
#   4. Crea las vistas 3D «VZ - SECTORES» y «VZ - ESTADO» con colores y la tabla «VZ - VALORIZACION».
# Se puede correr las veces que quieras: no duplica nada, solo actualiza valores.
#
# Entradas: IN[0] = ruta del JSON de datos · IN[1] = fecha de corte dd/mm/aaaa (vacío = hoy)

import clr, os, json, datetime
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
import System
from System.Collections.Generic import List

doc = DocumentManager.Instance.CurrentDBDocument
app = DocumentManager.Instance.CurrentUIApplication.Application

GRUPO = "NEXO OBRA (NO)"
CARPETA = os.path.dirname(doc.PathName) or os.environ["TEMP"]
ARCHIVO_SP = os.path.join(CARPETA, "Parametros_NEXO_NO.txt")
CATEGORIAS = [BuiltInCategory.OST_StructuralFoundation, BuiltInCategory.OST_StructuralColumns,
              BuiltInCategory.OST_StructuralFraming, BuiltInCategory.OST_Floors, BuiltInCategory.OST_Walls,
              BuiltInCategory.OST_Stairs, BuiltInCategory.OST_Ramps]

TXT, ENT, NUM, VOL, ARE, MON = "txt", "ent", "num", "vol", "are", "mon"
# (nombre, tipo, descripción, quién lo escribe)
PARAMETROS = [
    # --- planificación (los escribe este Dynamo desde el cronograma)
    ("VZ_Paquete", TXT, "Paquete de vaciado: T01..T10 (verticales de un nivel + techo superior), CIM, AN1, AN2", "Dynamo"),
    ("VZ_Sector", TXT, "Sector del paquete: A o B", "Dynamo"),
    ("VZ_Actividad", TXT, "Actividad del cronograma que vacía el elemento (ej. T05-CV-A)", "Dynamo"),
    ("VZ_Partida", TXT, "Ítem del presupuesto S10 del concreto del elemento (ej. 04.06.01)", "Dynamo"),
    ("VZ_Fecha_Prog", TXT, "Fecha programada de vaciado (dd/mm/aaaa)", "Dynamo"),
    # --- metrado y costo contractual (los escribe este Dynamo desde el presupuesto)
    ("VZ_Concreto_m3", VOL, "Metrado de concreto del presupuesto (aligerado: 0.0875 m3/m2)", "Dynamo"),
    ("VZ_Encofrado_m2", ARE, "Metrado de encofrado del presupuesto", "Dynamo"),
    ("VZ_Acero_kg", NUM, "Acero de refuerzo del presupuesto (kg, por cuantía)", "Dynamo"),
    ("VZ_Monto_Contractual", MON, "Costo directo del elemento: concreto + encofrado + acero (S/ sin IGV)", "Dynamo"),
    # --- ejecución (las escribe la app NEXO OBRA o el residente; este Dynamo no las borra)
    ("VZ_Fecha_Acero", TXT, "Fecha real de colocación del acero (dd/mm/aaaa)", "App"),
    ("VZ_Fecha_Encofrado", TXT, "Fecha real del encofrado (dd/mm/aaaa)", "App"),
    ("VZ_Fecha_Concreto", TXT, "Fecha real de vaciado (dd/mm/aaaa)", "App"),
    # --- control (se calculan a la fecha de corte)
    ("VZ_Estado", TXT, "Programado · Atrasado · En ejecución · Ejecutado · Revisar fecha", "Calculado"),
    ("VZ_Monto_Valorizado", MON, "Monto valorizado a la fecha de corte (S/ sin IGV)", "Calculado"),
    ("VZ_Valorizacion_N", ENT, "N° de valorización mensual de la última etapa ejecutada (1 = primer mes de obra)", "Calculado"),
]
SPEC = {TXT: SpecTypeId.String.Text, ENT: SpecTypeId.Int.Integer, NUM: SpecTypeId.Number,
        VOL: SpecTypeId.Volume, ARE: SpecTypeId.Area, MON: SpecTypeId.Currency}
COLOR_SECTOR = {"A": (37, 99, 235), "B": (229, 184, 105)}
COLOR_ESTADO = {"Programado": (200, 200, 200), "Atrasado": (220, 38, 38), "En ejecución": (245, 158, 11),
                "Ejecutado": (22, 163, 74), "Revisar fecha": (147, 51, 234)}


def leer_fecha(texto):
    """dd/mm/aaaa → date · vacío → None · ilegible → 'error'."""
    if texto is None or not str(texto).strip():
        return None
    t = str(texto).strip().replace("-", "/").replace(".", "/")
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.datetime.strptime(t, fmt).date()
        except ValueError:
            pass
    return "error"


def a_id(n):
    try:
        return ElementId(System.Int64(n))
    except Exception:
        return ElementId(n)


def poner(el, nombre, valor):
    """Escribe un valor según el tipo de almacenamiento del parámetro."""
    p = el.LookupParameter(nombre)
    if p is None or p.IsReadOnly:
        return False
    st = p.StorageType
    if st == StorageType.String:
        return p.Set(str(valor))
    if st == StorageType.Integer:
        try:
            return p.Set.Overloads[System.Int32](int(valor))
        except Exception:
            return p.Set(int(valor))
    if st == StorageType.Double:
        try:
            return p.Set.Overloads[System.Double](float(valor))
        except Exception:
            return p.Set(float(valor))
    return False


def texto(el, nombre):
    p = el.LookupParameter(nombre)
    return (p.AsString() or "") if p is not None else ""


def color(rgb):
    return Color(System.Byte(rgb[0]), System.Byte(rgb[1]), System.Byte(rgb[2]))


m3 = lambda v: UnitUtils.ConvertToInternalUnits(v, UnitTypeId.CubicMeters)
m2 = lambda v: UnitUtils.ConvertToInternalUnits(v, UnitTypeId.SquareMeters)
reporte = []

try:
    # ------------------------------------------------------------------ entradas
    ruta = IN[0] if len(IN) > 0 and IN[0] else ""
    if not ruta or not os.path.exists(ruta):
        raise Exception("No encuentro el archivo de datos: %s" % ruta)
    with open(ruta, encoding="utf-8") as fh:
        datos = json.load(fh)
    modelo = os.path.splitext(os.path.basename(doc.PathName or doc.Title))[0]
    if datos.get("modelo") and datos["modelo"].lower() != modelo.lower():
        reporte.append("AVISO: el archivo es del modelo %s y el abierto es %s." % (datos["modelo"], modelo))
    corte = leer_fecha(IN[1] if len(IN) > 1 else "")
    if corte in (None, "error"):
        if corte == "error":
            reporte.append("Fecha de corte ilegible: se usa la de hoy.")
        corte = datetime.date.today()
    inicio = datetime.date.fromisoformat(datos["inicio_obra"])
    reporte.append("Fecha de corte: %s · inicio de obra %s" % (corte.strftime("%d/%m/%Y"), inicio.strftime("%d/%m/%Y")))

    TransactionManager.Instance.EnsureInTransaction(doc)

    # ------------------------------------------------------------------ 1. parámetros compartidos
    vinculados = {}
    it = doc.ParameterBindings.ForwardIterator(); it.Reset()
    while it.MoveNext():
        vinculados[it.Key.Name] = it.Key
    sp_original = app.SharedParametersFilename
    creados, existian = [], []
    try:
        if not os.path.exists(ARCHIVO_SP):
            open(ARCHIVO_SP, "w").close()
        app.SharedParametersFilename = ARCHIVO_SP
        archivo = app.OpenSharedParameterFile()
        grupo = archivo.Groups.get_Item(GRUPO) or archivo.Groups.Create(GRUPO)
        cats = app.Create.NewCategorySet()
        for bic in CATEGORIAS:
            cat = Category.GetCategory(doc, bic)
            if cat is not None and cat.AllowsBoundParameters:
                cats.Insert(cat)
        vinculo = app.Create.NewInstanceBinding(cats)
        for nombre, tipo, desc, _ in PARAMETROS:
            definicion = grupo.Definitions.get_Item(nombre)
            if definicion is None:
                opc = ExternalDefinitionCreationOptions(nombre, SPEC[tipo])
                opc.Description = desc
                definicion = grupo.Definitions.Create(opc)
            if nombre in vinculados:
                doc.ParameterBindings.ReInsert(vinculados[nombre], vinculo, GroupTypeId.IdentityData)
                existian.append(nombre)
            elif doc.ParameterBindings.Insert(definicion, vinculo, GroupTypeId.IdentityData):
                creados.append(nombre)
    finally:
        app.SharedParametersFilename = sp_original
    doc.Regenerate()
    reporte.append("Parámetros VZ_: %d creados · %d ya existían · archivo %s" % (len(creados), len(existian), ARCHIVO_SP))

    id_param = {}
    it = doc.ParameterBindings.ForwardIterator(); it.Reset()
    while it.MoveNext():
        if it.Key.Name.startswith("VZ_"):
            id_param[it.Key.Name] = it.Key.Id

    # ------------------------------------------------------------------ 2-3. valores por elemento
    escritos, faltan, fechas_malas, conteo = 0, [], [], {}
    contractual, valorizado_total = 0.0, 0.0
    for e in datos["elementos"]:
        el = doc.GetElement(a_id(e["id"]))
        if el is None:
            faltan.append(e["id"]); continue
        monto = round(e["monto_concreto"] + e["monto_encofrado"] + e["monto_acero"], 2)
        ok = all([
            poner(el, "VZ_Paquete", e["paquete"]), poner(el, "VZ_Sector", e["sector"]),
            poner(el, "VZ_Actividad", e["actividad"]), poner(el, "VZ_Partida", e["partida"]),
            poner(el, "VZ_Fecha_Prog", e["fecha_prog"]),
            poner(el, "VZ_Concreto_m3", m3(e["concreto_m3"])), poner(el, "VZ_Encofrado_m2", m2(e["encofrado_m2"])),
            poner(el, "VZ_Acero_kg", e["acero_kg"]), poner(el, "VZ_Monto_Contractual", monto)])
        # ejecución: se leen, no se tocan
        fa, fe, fc = (leer_fecha(texto(el, n)) for n in ("VZ_Fecha_Acero", "VZ_Fecha_Encofrado", "VZ_Fecha_Concreto"))
        prog = leer_fecha(e["fecha_prog"])
        if "error" in (fa, fe, fc):
            estado, val, ult = "Revisar fecha", 0.0, None
            fechas_malas.append(e["id"])
        else:
            hechas = [(d, m) for d, m in ((fa, e["monto_acero"]), (fe, e["monto_encofrado"]), (fc, e["monto_concreto"]))
                      if d is not None and d <= corte]
            val = round(sum(m for _, m in hechas), 2)
            ult = max((d for d, _ in hechas), default=None)
            if fc is not None and fc <= corte:
                estado = "Ejecutado"
            elif hechas:
                estado = "En ejecución"
            elif prog and corte > prog:
                estado = "Atrasado"
            else:
                estado = "Programado"
        n_val = (ult.year - inicio.year) * 12 + ult.month - inicio.month + 1 if ult else 0
        ok = all([ok, poner(el, "VZ_Estado", estado), poner(el, "VZ_Monto_Valorizado", val),
                  poner(el, "VZ_Valorizacion_N", n_val)])
        escritos += 1 if ok else 0
        conteo[estado] = conteo.get(estado, 0) + 1
        contractual += monto; valorizado_total += val
    reporte.append("Elementos escritos: %d de %d" % (escritos, len(datos["elementos"])))
    if faltan:
        reporte.append("No encontrados en el modelo (%d): %s" % (len(faltan), faltan[:20]))
    if fechas_malas:
        reporte.append("Fechas ilegibles (usa dd/mm/aaaa) en: %s" % fechas_malas[:20])
    reporte.append("Estados: " + " · ".join("%s %d" % kv for kv in sorted(conteo.items())))
    reporte.append("Contractual S/ {:,.2f} · valorizado S/ {:,.2f}".format(contractual, valorizado_total))

    # ------------------------------------------------------------------ 4. filtros, vistas y tabla
    ids_cat = List[ElementId]([Category.GetCategory(doc, b).Id for b in CATEGORIAS
                               if Category.GetCategory(doc, b) is not None and Category.GetCategory(doc, b).AllowsBoundParameters])
    filtros = {f.Name: f for f in FilteredElementCollector(doc).OfClass(ParameterFilterElement)}
    relleno = next((fp for fp in FilteredElementCollector(doc).OfClass(FillPatternElement) if fp.GetFillPattern().IsSolidFill), None)

    def filtro(nombre, parametro, valor):
        try:
            regla = ParameterFilterRuleFactory.CreateEqualsRule(id_param[parametro], valor)
        except Exception:
            regla = ParameterFilterRuleFactory.CreateEqualsRule(id_param[parametro], valor, False)
        epf = ElementParameterFilter(regla)
        if nombre in filtros:
            f = filtros[nombre]; f.SetCategories(ids_cat); f.SetElementFilter(epf)
        else:
            f = ParameterFilterElement.Create(doc, nombre, ids_cat, epf); filtros[nombre] = f
        return f

    def vista_3d(nombre):
        for v in FilteredElementCollector(doc).OfClass(View3D):
            if not v.IsTemplate and v.Name == nombre:
                return v
        vft = next(t for t in FilteredElementCollector(doc).OfClass(ViewFamilyType) if t.ViewFamily == ViewFamily.ThreeDimensional)
        v = View3D.CreateIsometric(doc, vft.Id); v.Name = nombre
        return v

    def pintar(v, f, rgb):
        ogs = OverrideGraphicSettings()
        ogs.SetSurfaceForegroundPatternId(relleno.Id); ogs.SetSurfaceForegroundPatternColor(color(rgb))
        ogs.SetCutForegroundPatternId(relleno.Id); ogs.SetCutForegroundPatternColor(color(rgb))
        if not v.IsFilterApplied(f.Id):
            v.AddFilter(f.Id)
        v.SetFilterOverrides(f.Id, ogs); v.SetFilterVisibility(f.Id, True)

    for nombre_vista, parametro, colores, prefijo in (("VZ - SECTORES", "VZ_Sector", COLOR_SECTOR, "VZ Sector "),
                                                       ("VZ - ESTADO", "VZ_Estado", COLOR_ESTADO, "VZ Estado ")):
        try:   # si una vista falla, los parámetros y valores igual quedan
            v = vista_3d(nombre_vista)
            v.ViewTemplateId = ElementId.InvalidElementId
            v.DisplayStyle = DisplayStyle.Shading
            for valor, rgb in colores.items():
                pintar(v, filtro(prefijo + valor, parametro, valor), rgb)
            reporte.append("Vista 3D «%s»: %d filtros de color" % (nombre_vista, len(colores)))
        except Exception as ex:
            reporte.append("La vista «%s» no se pudo crear: %s" % (nombre_vista, ex))

    try:
        NOMBRE_TABLA = "VZ - VALORIZACION"
        tabla = next((t for t in FilteredElementCollector(doc).OfClass(ViewSchedule)
                      if not t.IsTemplate and t.Name == NOMBRE_TABLA), None)
        if tabla is None:
            tabla = ViewSchedule.CreateSchedule(doc, ElementId.InvalidElementId)   # tabla multicategoría
            tabla.Name = NOMBRE_TABLA
        defn = tabla.Definition
        campos = ["VZ_Paquete", "VZ_Sector", "VZ_Actividad", "VZ_Partida", "VZ_Fecha_Prog", "VZ_Concreto_m3",
                  "VZ_Encofrado_m2", "VZ_Acero_kg", "VZ_Monto_Contractual", "VZ_Fecha_Acero", "VZ_Fecha_Encofrado",
                  "VZ_Fecha_Concreto", "VZ_Estado", "VZ_Valorizacion_N", "VZ_Monto_Valorizado"]
        presentes = {defn.GetField(i).GetName(): defn.GetField(i) for i in range(defn.GetFieldCount())}
        disponibles = {sf.GetName(doc): sf for sf in defn.GetSchedulableFields()}
        for c in campos:
            if c not in presentes and c in disponibles:
                presentes[c] = defn.AddField(disponibles[c])
        defn.ClearSortGroupFields()
        for c in ("VZ_Paquete", "VZ_Sector", "VZ_Actividad"):
            defn.AddSortGroupField(ScheduleSortGroupField(presentes[c].FieldId, ScheduleSortOrder.Ascending))
        defn.ClearFilters()
        defn.AddFilter(ScheduleFilter(presentes["VZ_Actividad"].FieldId, ScheduleFilterType.HasValue))
        defn.IsItemized = True
        defn.ShowGrandTotal = True
        for c in ("VZ_Concreto_m3", "VZ_Encofrado_m2", "VZ_Acero_kg", "VZ_Monto_Contractual", "VZ_Monto_Valorizado"):
            try:
                presentes[c].DisplayType = ScheduleFieldDisplayType.Totals
            except Exception:
                pass
        reporte.append("Tabla «%s» lista (%d campos)" % (NOMBRE_TABLA, len(campos)))
    except Exception as ex:
        reporte.append("La tabla no se pudo crear (los parámetros sí quedaron): %s" % ex)

    TransactionManager.Instance.TransactionTaskDone()
    reporte.append("Listo. Revisa las vistas y guarda el modelo con Ctrl+S.")
except Exception as ex:
    TransactionManager.Instance.ForceCloseTransaction()
    reporte.append("ERROR: %s" % ex)

OUT = reporte
