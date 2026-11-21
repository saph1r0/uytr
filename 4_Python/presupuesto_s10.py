"""NEXO OBRA (NO) · Presupuesto de estructuras con formato tipo S10 a partir del modelo de Revit.

Uso:  python presupuesto_s10.py <metrados_modelo.json> <salida.xlsx>
Hojas: Presupuesto · Análisis de Precios Unitarios · Precios y Cantidades de Recursos · Hoja de Metrados · Supuestos y Fuentes
Todo se enlaza con fórmulas: si cambias un precio, un rendimiento o una cuantía, el presupuesto se recalcula.
"""
import re
import sys
from collections import defaultdict
from datetime import date

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

import metrados as M

F = "Arial"
f9, f9b, f8, f8b = Font(name=F, size=9), Font(name=F, size=9, bold=True), Font(name=F, size=8), Font(name=F, size=8, bold=True)
f12b = Font(name=F, size=12, bold=True)
fin9 = Font(name=F, size=9, color="0000FF")          # entradas editables
fgr8 = Font(name=F, size=8, color="808080")
fin = Side(style="thin")
caja = Border(top=fin, bottom=fin, left=fin, right=fin)
linea_ab = Border(bottom=fin)
linea_ar = Border(top=fin)
N2, N4 = '#,##0.00;-#,##0.00;"-"', '#,##0.0000;-#,##0.0000;"-"'
DER, CEN, IZQ = Alignment(horizontal="right"), Alignment(horizontal="center"), Alignment(horizontal="left")
HOY = date.today().strftime("%d/%m/%Y")
OBRA = "EDIFICIO MULTIFAMILIAR CON SÓTANOS - MODELO REP0-MOD2-R02"
LUGAR = "AREQUIPA - AREQUIPA - AREQUIPA"
CLIENTE = "(POR DEFINIR)"

# ---------------------------------------------------------------- INSUMOS (precios sin IGV, S/.)
# código, descripción, unidad, tipo (MO/MAT/EQ), precio (None = se calcula), fuente
INS = [
    ("0147010002", "OPERARIO", "hh", "MO", None, "Jornal 2026 (RM 197-2025-TR) + BUC + leyes sociales + movilidad; ver Supuestos"),
    ("0147010003", "OFICIAL", "hh", "MO", None, "Ídem"),
    ("0147010004", "PEON", "hh", "MO", None, "Ídem"),
    ("0202010022", "CLAVOS PARA MADERA C/C 3\"", "kg", "MAT", 5.08, "Supuesto: S/ 6.00 con IGV (ferretería)"),
    ("0202040009", "ALAMBRE NEGRO RECOCIDO N° 16", "kg", "MAT", 5.51, "Supuesto: S/ 6.50 con IGV"),
    ("0202040010", "ALAMBRE NEGRO RECOCIDO N° 8", "kg", "MAT", 5.08, "Supuesto: S/ 6.00 con IGV"),
    ("0203020003", "ACERO CORRUGADO fy=4200 kg/cm2 GRADO 60", "kg", "MAT", 3.43, "S/ 4,050/t con IGV (varilla 1/2\", OneEstimate ago-2026) / 1.18"),
    ("0204000000", "ARENA GRUESA", "m3", "MAT", 52.54, "S/ 62.00 con IGV (rango 55-70, OneEstimate ago-2026) / 1.18"),
    ("0205000003", "PIEDRA CHANCADA 1/2\"", "m3", "MAT", 72.03, "S/ 85.00 con IGV (rango 75-95, OneEstimate ago-2026) / 1.18"),
    ("0207030001", "HORMIGON", "m3", "MAT", 48.31, "S/ 57.00 con IGV (rango 50-65, OneEstimate ago-2026) / 1.18"),
    ("0217010008", "LADRILLO DE TECHO HUECO 15x30x30 cm", "und", "MAT", 2.54, "S/ 3.00 con IGV (OneEstimate ago-2026) / 1.18"),
    ("0217010010", "LADRILLO KK 18 HUECOS 9x12.5x23 cm", "und", "MAT", 1.23, "S/ 1.45 con IGV (OneEstimate ago-2026) / 1.18"),
    ("0221000000", "CEMENTO PORTLAND TIPO I (42.5 kg)", "bol", "MAT", 25.42, "S/ 30.00 con IGV (rango 28.50-31.00, OneEstimate ago-2026) / 1.18"),
    ("0229060005", "YESO EN BOLSAS DE 28 kg", "bol", "MAT", 11.86, "Supuesto a partir de yeso 20 kg S/ 9.50-12.00 con IGV"),
    ("0239050000", "AGUA", "m3", "MAT", 5.00, "Supuesto"),
    ("0243010003", "MADERA TORNILLO INCLUYE CORTE PARA ENCOFRADO", "p2", "MAT", 5.51, "Supuesto: S/ 6.50 con IGV"),
    ("0337010001", "HERRAMIENTAS MANUALES", "%MO", "EQ", None, "3 % de la mano de obra"),
    ("0348040027", "CAMION VOLQUETE 15 m3", "hm", "EQ", 170.00, "Supuesto: tarifa de alquiler con operador, Arequipa"),
    ("0348960005", "MEZCLADORA DE CONCRETO TAMBOR 18 HP 11 p3", "hm", "EQ", 25.00, "Supuesto"),
    ("0348960002", "WINCHE ELECTRICO 3.6 HP DE DOS BALDES", "hm", "EQ", 15.00, "Supuesto"),
    ("0349040010", "EXCAVADORA SOBRE ORUGAS 170-250 HP 1.1-2.75 yd3", "hm", "EQ", 250.00, "Supuesto: tarifa de alquiler con operador"),
    ("0349040033", "CARGADOR SOBRE LLANTAS 125-155 HP 3 yd3", "hm", "EQ", 190.00, "Supuesto: tarifa de alquiler con operador"),
    ("0349070004", "VIBRADOR DE CONCRETO 4 HP 1.35\"", "hm", "EQ", 12.00, "Supuesto"),
    ("0349100011", "COMPACTADORA VIBRATORIA TIPO PLANCHA 4 HP", "hm", "EQ", 15.00, "Supuesto"),
    ("0349880020", "NIVEL TOPOGRAFICO CON TRIPODE", "hm", "EQ", 6.00, "Supuesto"),
    ("0205010004", "AFIRMADO", "m3", "MAT", 38.14, "Supuesto: S/ 45.00 con IGV puesto en obra"),
    ("0221010350", "CONCRETO PREMEZCLADO f'c=350 kg/cm2", "m3", "MAT", 440.68, "Supuesto: S/ 520.00 con IGV — cotizar con concretera de Arequipa"),
    ("0349190010", "SERVICIO DE BOMBA DE CONCRETO", "m3", "EQ", 38.14, "Supuesto: S/ 45.00 con IGV por m3 bombeado"),
    ("0410010001", "SC CERCO PROVISIONAL h=2.40 m (PANELES Y POSTES)", "m", "SC", 85.00, "Supuesto: subcontrato, a cotizar"),
    ("0410010002", "SC OFICINA, ALMACEN, VESTUARIOS Y SS.HH. PROVISIONALES", "glb", "SC", 18000.00, "Supuesto: subcontrato, a cotizar"),
    ("0410010003", "SC INSTALACIONES PROVISIONALES DE AGUA, DESAGUE Y ENERGIA", "glb", "SC", 8000.00, "Supuesto: subcontrato, a cotizar"),
    ("0410010004", "SC CARTEL DE OBRA 3.60x2.40 m", "und", "SC", 2500.00, "Supuesto: subcontrato, a cotizar"),
    ("0410010005", "SC MOVILIZACION Y DESMOVILIZACION DE EQUIPOS", "glb", "SC", 6000.00, "Supuesto: subcontrato, a cotizar"),
    ("0410010006", "SC SEGURIDAD Y SALUD EN EL TRABAJO (EPP, SEÑALIZACION, CAPACITACION)", "mes", "SC", 3500.00, "Supuesto: costo mensual"),
    ("0410010007", "SC ANCLAJE POSTENSADO TEMPORAL (PERFORACION, CABLE, INYECCION, TENSADO Y DESTENSADO)", "und", "SC", 5500.00,
     "Supuesto: subcontrato especializado — cotizar (es la partida más incierta)"),
    ("0410010008", "SC ENSAYOS DE CONTROL DE CALIDAD (PROBETAS, ACERO, COMPACTACION)", "glb", "SC", 12000.00, "Supuesto: laboratorio, a cotizar"),
    ("0410010009", "SC IMPERMEABILIZACION Y WATERSTOP EN CISTERNA", "glb", "SC", 6500.00, "Supuesto: subcontrato, a cotizar"),
]
INS_D = {c: (d, u, t) for c, d, u, t, _, _ in INS}
OP, OF, PE = "0147010002", "0147010003", "0147010004"
HER = "0337010001"

# Dosificación por m3 (cemento bol, arena m3, piedra m3, agua m3) – referencial
DOS = {140: (7.01, 0.51, 0.64, 0.184), 210: (9.73, 0.52, 0.53, 0.186), 280: (11.50, 0.45, 0.51, 0.190), 350: (13.60, 0.42, 0.48, 0.200)}


def apu_sc(und, cod, q=1):
    """Partida resuelta con un subcontrato (como en S10: sección «Subcontratos»)."""
    return dict(und=und, rend=1, mo=[], mat=[], eq=[], sc=[(cod, q)])


def apu_concreto(fc, rend, winche):
    if fc == 350:   # premezclado con bomba: menos cuadrilla, sin mezcladora ni winche
        return dict(und="m3", rend=rend * 2, mo=[(OP, 2), (OF, 1), (PE, 6)],
                    mat=[("0221010350", 1.03)], eq=[("0349070004", 1), ("0349190010", None)])
    c, a, p, w = DOS[fc]
    eq = [("0348960005", 1), ("0349070004", 1)] + ([("0348960002", 1)] if winche else [])
    return dict(und="m3", rend=rend, mo=[(OP, 2), (OF, 2), (PE, 8)],
                mat=[("0221000000", c), ("0204000000", a), ("0205000003", p), ("0239050000", w)], eq=eq)


def apu_encofrado(rend, madera, peon=0):
    mo = [(OP, 1), (OF, 1)] + ([(PE, peon)] if peon else [])
    return dict(und="m2", rend=rend, mo=mo, mat=[("0202040010", 0.20), ("0202010022", 0.15), ("0243010003", madera)], eq=[])


APU_ACERO = dict(und="kg", rend=250, mo=[(OP, 1), (OF, 1)], mat=[("0202040009", 0.03), ("0203020003", 1.07)], eq=[])

# Grupos estructurales: nombre, rendimiento concreto, ¿winche?, rendimiento encofrado, madera p2/m2, peón encofrado, cuantía kg/m3
GRUPOS = {
    "ZAP": ("ZAPATAS", 25, False, 12, 2.50, 0, 60),
    "CIMM": ("CIMIENTOS DE MUROS DE ANCLAJE", 25, False, 12, 2.50, 0, 60),
    "LOSC": ("LOSAS DE CIMENTACION Y CISTERNA", 25, False, 12, 2.50, 0, 90),
    "MUR": ("MUROS DE CONCRETO ARMADO", 12, False, 12, 3.20, 0, 90),
    "COL": ("COLUMNAS", 10, True, 10, 3.20, 0, 160),
    "PLA": ("PLACAS", 12, True, 12, 3.20, 0, 120),
    "VIG": ("VIGAS", 14, True, 8.5, 4.50, 0, 130),
    "VCH": ("VIGAS CHATAS Y VIGUETAS", 14, True, 10, 3.50, 0, 90),
    "LMA": ("LOSAS MACIZAS", 20, True, 12, 3.80, 1, 90),
    "ALI": ("LOSAS ALIGERADAS", 20, True, 14, 3.40, 1, 45),
    "ESC": ("ESCALERAS", 10, True, 8, 4.80, 0, 80),
    "RAM": ("RAMPAS VEHICULARES", 18, True, 10, 3.80, 1, 80),
}


def construir(ruta_json, salida):
    D, filas, X = M.calcular(ruta_json)
    # ---------------- agregación para la hoja de metrados: (grupo, f'c, tipo, nivel) -> n, concreto, encofrado, área
    ag = defaultdict(lambda: [0, 0.0, 0.0, 0.0])
    for f in filas:
        k = (f["g"], f["fc"], f["tipo"], f["nivel"])
        r = ag[k]; r[0] += 1; r[1] += f["conc"]; r[2] += f["enc"]; r[3] += f["area"]

    def lineas(grupo, campo, fc=None):
        idx = {"conc": 1, "enc": 2, "area": 3}[campo]
        out = [(f"{k[2]} — {k[3]}", v[0], v[idx]) for k, v in ag.items()
               if k[0] == grupo and (fc is None or k[1] == fc) and v[idx] > 1e-6]
        return sorted(out)

    wb = Workbook()
    wsP = wb.active; wsP.title = "Presupuesto"
    wsA = wb.create_sheet("Análisis de Precios Unitarios")
    wsI = wb.create_sheet("Precios y Cantidades Recursos")
    wsM = wb.create_sheet("Hoja de Metrados")
    wsS = wb.create_sheet("Supuestos y Fuentes")
    SA, SI, SM, SS = "'Análisis de Precios Unitarios'", "'Precios y Cantidades Recursos'", "'Hoja de Metrados'", "'Supuestos y Fuentes'"

    # ================================================================ SUPUESTOS
    wsS["A1"] = "SUPUESTOS Y FUENTES"; wsS["A1"].font = f12b
    wsS["A2"] = "Celdas en azul = datos editables. Todo lo demás se calcula. Precios sin IGV, en soles."; wsS["A2"].font = f8
    r = 4
    wsS.cell(r, 1, "COSTO DE MANO DE OBRA (hora-hombre)").font = f9b; r += 1
    for c, t in enumerate(["Categoría", "Jornal básico S/./día", "BUC %", "Leyes sociales % (s/ básico)", "Movilidad S/./día", "Costo día", "Costo hh"], 1):
        x = wsS.cell(r, c, t); x.font = f8b; x.border = caja; x.alignment = CEN
    r += 1
    MO_HH = {}
    for cod, nom, jb, buc in ((OP, "OPERARIO", 89.30, 0.32), (OF, "OFICIAL", 69.75, 0.30), (PE, "PEON", 62.80, 0.30)):
        wsS.cell(r, 1, nom).font = f9
        for c, v, fmt in ((2, jb, N2), (3, buc, "0.00%"), (4, 1.1182, "0.00%"), (5, 8.60, N2)):
            x = wsS.cell(r, c, v); x.font = fin9; x.number_format = fmt
        wsS.cell(r, 6, f"=B{r}*(1+C{r}+D{r})+E{r}").number_format = N2
        wsS.cell(r, 7, f"=ROUND(F{r}/8,2)").number_format = N2
        for c in range(1, 8):
            wsS.cell(r, c).border = caja
            if c not in (2, 3, 4, 5):
                wsS.cell(r, c).font = f9
        MO_HH[cod] = f"{SS}!$G${r}"
        r += 1
    wsS.cell(r, 1, "Fuente jornal: RM 197-2025-TR / El Peruano 2026 (operario 89.30, oficial 69.75, peón 62.80). BUC 32/30 % y leyes sociales 111.82 % "
                   "son valores usuales de la tabla CAPECO-FTCCP: verificar con la tabla vigente de CAPECO.").font = f8
    wsS.cell(r, 1).comment = Comment("Leyes sociales 111.82 %: valor típico usado en tablas de costo de mano de obra CAPECO; confirmar.", "NEXO")
    r += 2

    wsS.cell(r, 1, "PARÁMETROS DEL PRESUPUESTO").font = f9b; r += 1
    PAR = {}
    for clave, nom, v, fmt in (("gg", "Gastos generales (% del costo directo)", 0.12, "0.00%"), ("ut", "Utilidad (% del costo directo)", 0.08, "0.00%"),
                               ("igv", "IGV", 0.18, "0.00%"), ("her", "Herramientas manuales (% de la mano de obra)", 0.03, "0.00%"),
                               ("huella", "Huella de sótanos (m2) — caja del modelo bajo cota -1.00", round(X["huella_m2"], 2), N2),
                               ("terreno", "Cota de terreno natural (m)", M.COTA_TERRENO, N2),
                               ("fondo", "Fondo de cimentación más bajo (m) — del modelo", round(X["fondo_cim"], 3), N2),
                               ("esp", "Factor de esponjamiento", M.ESPONJAMIENTO, N2)):
        wsS.cell(r, 1, nom).font = f9
        x = wsS.cell(r, 4, v); x.font = fin9; x.number_format = fmt
        PAR[clave] = f"{SS}!$D${r}"; r += 1
    r += 1
    wsS.cell(r, 1, "CUANTÍAS DE ACERO (kg por m3 de concreto) — el modelo solo trae 197 barras, no todo el refuerzo").font = f9b; r += 1
    CUA = {}
    for g, v in list(GRUPOS.items()) + [("ALB", ("ALBAÑILERIA ARMADA",) + (None,) * 5 + (30,))]:
        wsS.cell(r, 1, v[0]).font = f9
        x = wsS.cell(r, 4, v[6]); x.font = fin9; x.number_format = N2
        CUA[g] = f"{SS}!$D${r}"; r += 1
    r += 1
    wsS.cell(r, 1, "CRITERIOS DE METRADO").font = f9b; r += 1
    for a, b in M.SUPUESTOS + [("Dosificación de concreto (por m3)", "f'c 140: 7.01 bol; 210: 9.73; 280: 11.50; 350: 13.60 bol de cemento (referencial; para 350 se recomienda premezclado)"),
                               ("Rendimientos", "Referenciales tipo CAPECO (no es la publicación oficial de CAPECO, que es de pago); ver cada análisis"),
                               ("Fuentes de precios", "oneestimate.ai (lista de precios Perú, ago-2026; rendimientos, abr-2026); El Peruano (jornales 2026). Precios sin IGV = precio de ferretería / 1.18"),
                               ("Lugar y fecha", f"{LUGAR}, costo al {HOY}. Cotizar en Arequipa antes de usar el presupuesto")]:
        wsS.cell(r, 1, a).font = f9b; wsS.cell(r, 3, b).font = f9; r += 1
    for c, w in zip("ABCDEFG", (44, 16, 12, 18, 14, 12, 12)): wsS.column_dimensions[c].width = w

    # ================================================================ PARTIDAS (estructura del presupuesto)
    # Cada partida: (item, descripción, apu dict, fuente de metrado)
    # fuente: ("hm", lista de líneas) | ("f", fórmula que usa otras partidas)
    P = []   # lista de nodos: ("T", item, desc) título | ("P", item, desc, apu, fuente, clave)
    P.append(("T", "01", "OBRAS PROVISIONALES Y TRABAJOS PRELIMINARES"))
    P.append(("P", "01.01", "TRAZO, NIVELES Y REPLANTEO DURANTE LA OBRA",
              dict(und="m2", rend=250, mo=[(OP, 1), (PE, 2)], mat=[("0229060005", 0.02), ("0243010003", 0.03)], eq=[("0349880020", 1)]),
              ("f", f"={PAR['huella']}"), "trazo"))
    # Obras provisionales (subcontratos con metrado editable en Supuestos)
    wsS["I4"] = "OBRAS PROVISIONALES (metrados editables)"; wsS["I4"].font = f9b
    perim = 2 * (X["dim_huella"][0] + X["dim_huella"][1])
    prov = [("cerco", "01.02", "CERCO PROVISIONAL h=2.40 m", "m", "0410010001", round(perim, 2), "Perímetro de la huella del modelo"),
            ("ofic", "01.03", "OFICINA, ALMACEN, VESTUARIOS Y SS.HH. PROVISIONALES", "glb", "0410010002", 1, ""),
            ("inst", "01.04", "INSTALACIONES PROVISIONALES DE AGUA, DESAGUE Y ENERGIA", "glb", "0410010003", 1, ""),
            ("cart", "01.05", "CARTEL DE OBRA 3.60x2.40 m", "und", "0410010004", 1, ""),
            ("movi", "01.06", "MOVILIZACION Y DESMOVILIZACION DE EQUIPOS", "glb", "0410010005", 1, ""),
            ("sst", "01.07", "SEGURIDAD Y SALUD EN EL TRABAJO", "mes", "0410010006", 12, "Plazo supuesto del casco: 12 meses")]
    for i, (k, it, d, u, cod, q, nota) in enumerate(prov):
        rr = 5 + i
        wsS.cell(rr, 9, d).font = f9; wsS.cell(rr, 10, u).font = f9
        x = wsS.cell(rr, 11, q); x.font = fin9; x.number_format = N2
        wsS.cell(rr, 12, nota).font = f8
        PAR["m_" + k] = f"{SS}!$K${rr}"
        P.append(("P", it, d, apu_sc(u, cod), ("f", f"={PAR['m_' + k]}"), k))

    # Capas de excavación masiva (una por sótano), editables
    r0c = 13
    wsS.cell(r0c - 1, 9, "EXCAVACIÓN MASIVA POR CAPAS (un sótano por capa)").font = f9b
    for c, t in enumerate(["Capa / sótano", "Área m2", "Fondo m", "Espesor m", "Volumen m3"], 9):
        x = wsS.cell(r0c, c, t); x.font = f8b; x.border = caja; x.alignment = CEN
    for i, cp in enumerate(X["capas"]):
        rr = r0c + 1 + i
        wsS.cell(rr, 9, cp["nivel"]).font = f9
        x = wsS.cell(rr, 10, round(cp["area"], 2)); x.font = fin9; x.number_format = N2
        x = wsS.cell(rr, 11, cp["fondo"]); x.font = fin9; x.number_format = N2
        arriba = PAR["terreno"] if i == 0 else f"K{rr - 1}"
        x = wsS.cell(rr, 12, f"=MAX(0,{arriba}-K{rr})"); x.number_format = N2; x.font = f9
        x = wsS.cell(rr, 13, f"=J{rr}*L{rr}"); x.number_format = N2; x.font = f9
    rt = r0c + 1 + len(X["capas"])
    wsS.cell(rt, 9, "Total excavación masiva").font = f9b
    x = wsS.cell(rt, 13, f"=SUM(M{r0c + 1}:M{rt - 1})"); x.font = f9b; x.number_format = N2
    wsS.cell(rt + 1, 9, "Huella = caja en planta de los elementos de cada sótano; fondo = nivel del sótano - 0.10 m").font = f8
    for c, w in zip("IJKLM", (52, 10, 10, 10, 12)): wsS.column_dimensions[c].width = w

    P.append(("T", "02", "MOVIMIENTO DE TIERRAS"))
    P.append(("P", "02.01", "EXCAVACION MASIVA CON MAQUINARIA PARA SOTANOS",
              dict(und="m3", rend=350, mo=[(OP, 1), (PE, 2)], mat=[], eq=[("0349040010", 1)]),
              ("f", f"={SS}!$M${rt}"), "exm"))
    P.append(("P", "02.02", "EXCAVACION MANUAL DE ZANJAS PARA CIMENTACION (BAJO EL FONDO DE LA EXCAVACION MASIVA)",
              dict(und="m3", rend=3.5, mo=[(PE, 1)], mat=[], eq=[]), ("hm", X["lin_exz"]), "exz"))
    P.append(("P", "02.03", "REFINE, NIVELACION Y COMPACTACION DE FONDO",
              dict(und="m2", rend=40, mo=[(PE, 1)], mat=[], eq=[("0349100011", 0.5)]), ("f", "=SOLADO"), "ref"))
    P.append(("P", "02.04", "ELIMINACION DE MATERIAL EXCEDENTE C/CARGADOR Y VOLQUETE D=10 km",
              dict(und="m3", rend=300, mo=[(OP, 1), (PE, 1)], mat=[], eq=[("0349040033", 1), ("0348040027", 4)]),
              ("f", f"=(EXM+EXZ)*{PAR['esp']}"), "eli"))
    P.append(("P", "02.05", "ANCLAJES POSTENSADOS TEMPORALES PARA MUROS ANCLADOS", apu_sc("und", "0410010007"),
              ("hm", X["lin_anc"]), "anc"))
    P.append(("P", "02.06", "BASE DE AFIRMADO COMPACTADO e=0.15 m PARA FALSO PISO DE SOTANOS",
              dict(und="m2", rend=80, mo=[(OF, 1), (PE, 2)], mat=[("0205010004", 0.18), ("0239050000", 0.02)], eq=[("0349100011", 1)]),
              ("f", "=FPISO"), "afi"))
    P.append(("T", "03", "OBRAS DE CONCRETO SIMPLE"))
    P.append(("P", "03.01", "SOLADO PARA ZAPATAS Y CIMIENTOS e=0.10 m C:H 1:12",
              dict(und="m2", rend=120, mo=[(OP, 1), (OF, 1), (PE, 6)], mat=[("0221000000", 0.36), ("0207030001", 0.13), ("0239050000", 0.02)],
                   eq=[("0348960005", 1)]), ("hm", None), "sol"))
    P.append(("P", "03.02", "SUB-ZAPATAS CONCRETO f'c=140 kg/cm2", apu_concreto(140, 20, False), ("hm", lineas("SUBZ", "conc")), "subz"))
    P.append(("P", "03.03", "FALSO PISO DE SOTANOS e=0.10 m C:H 1:8",
              dict(und="m2", rend=120, mo=[(OP, 1), (OF, 1), (PE, 6)], mat=[("0221000000", 0.50), ("0207030001", 0.13), ("0239050000", 0.02)],
                   eq=[("0348960005", 1)]), ("hm", X["lin_fp"]), "fpi"))
    P.append(("T", "04", "OBRAS DE CONCRETO ARMADO"))
    sub = 0
    for g, (nom, rc, win, re_, mad, pe, _) in GRUPOS.items():
        fcs = sorted({k[1] for k in ag if k[0] == g})
        if not fcs:
            continue
        sub += 1
        it = f"04.{sub:02d}"
        P.append(("T", it, nom))
        n = 0
        for fc in fcs:
            n += 1
            P.append(("P", f"{it}.{n:02d}", f"CONCRETO f'c={fc} kg/cm2 EN {nom}", apu_concreto(fc, rc, win), ("hm", lineas(g, "conc", fc)), f"c_{g}_{fc}"))
        n += 1
        P.append(("P", f"{it}.{n:02d}", f"ENCOFRADO Y DESENCOFRADO NORMAL EN {nom}", apu_encofrado(re_, mad, pe),
                  ("hm", lineas(g, "area" if g == "ALI" else "enc")), f"e_{g}"))
        if g == "ALI":
            n += 1
            P.append(("P", f"{it}.{n:02d}", "LADRILLO DE TECHO HUECO 15x30x30 cm EN LOSA ALIGERADA",
                      dict(und="und", rend=1200, mo=[(OP, 0.1), (PE, 1)], mat=[("0217010008", 1.05)], eq=[]),
                      ("f", f"=ROUND(AREA_ALI*{M.LADRILLO_M2},0)"), "lad"))
        n += 1
        conc_claves = "+".join(f"c_{g}_{fc}" for fc in fcs)
        P.append(("P", f"{it}.{n:02d}", f"ACERO DE REFUERZO fy=4200 kg/cm2 EN {nom}", APU_ACERO,
                  ("f", f"=({conc_claves})*{CUA[g]}"), f"a_{g}"))
    P.append(("T", "05", "ALBAÑILERIA"))
    P.append(("P", "05.01", "MURO DE ALBAÑILERIA ARMADA e=0.15 m C/LADRILLO KK, MORTERO 1:4",
              dict(und="m2", rend=8, mo=[(OP, 1), (PE, 0.5)], mat=[("0217010010", 39), ("0221000000", 0.27), ("0204000000", 0.03), ("0239050000", 0.008)], eq=[]),
              ("hm", lineas("ALB", "area")), "alb"))
    P.append(("P", "05.02", "ACERO DE REFUERZO fy=4200 kg/cm2 EN ALBAÑILERIA ARMADA", APU_ACERO,
              ("f", f"=VOL_ALB*{CUA['ALB']}"), "a_ALB"))
    P.append(("T", "06", "VARIOS Y CONTROL DE CALIDAD"))
    P.append(("P", "06.01", "CURADO DE ELEMENTOS DE CONCRETO CON AGUA",
              dict(und="m2", rend=250, mo=[(PE, 1)], mat=[("0239050000", 0.01)], eq=[]), ("f", "=ENC_TOTAL"), "cur"))
    P.append(("P", "06.02", "ENSAYOS DE CONTROL DE CALIDAD (PROBETAS, ACERO, COMPACTACION)", apu_sc("glb", "0410010008"), ("f", "=1"), "ens"))
    P.append(("P", "06.03", "IMPERMEABILIZACION Y WATERSTOP EN CISTERNA", apu_sc("glb", "0410010009"), ("f", "=1"), "imp"))

    # líneas de metrado para excavación local y solado (por elemento de cimentación agrupado)
    ex = defaultdict(lambda: [0, 0.0, 0.0])
    for f in filas:
        if f["g"] in ("ZAP", "CIMM", "SUBZ", "LOSC"):
            k = f"{f['tipo']} — {f['nivel']}"
            ex[k][0] += 1; ex[k][2] += f["inf"]
            if f["g"] != "LOSC":
                ex[k][1] += f["inf"] * (f["caja"][5] - f["caja"][2] + 0.10)
    LIN_SOL = sorted((k, v[0], v[2]) for k, v in ex.items() if v[2] > 0)
    P = [(n[0], n[1], n[2], n[3], ("hm", LIN_SOL), n[5]) if n[0] == "P" and n[5] == "sol" else n for n in P]

    # ================================================================ HOJA DE METRADOS
    wsM["A1"] = "HOJA DE METRADOS"; wsM["A1"].font = f12b
    wsM["A2"] = f"Obra: {OBRA}"; wsM["A2"].font = f9
    wsM["A3"] = f"Fuente: modelo Revit {D['modelo']} leído con NexoMetrados ({D['fecha']}). Líneas agrupadas por tipo y nivel."; wsM["A3"].font = f8
    for c, t in enumerate(["Item", "Descripción", "Und.", "N° elem.", "Parcial", "Total"], 1):
        x = wsM.cell(5, c, t); x.font = f8b; x.border = caja; x.alignment = CEN
    rm = 6
    TOT_M = {}     # clave -> celda total en hoja de metrados
    for n in P:
        if n[0] != "P" or n[4][0] != "hm":
            continue
        _, it, desc, apu, (_, lins), clave = n
        wsM.cell(rm, 1, it).font = f9b; wsM.cell(rm, 2, desc).font = f9b; wsM.cell(rm, 3, apu["und"]).font = f9
        r0 = rm + 1
        for d, cnt, v in (lins or []):
            rm += 1
            wsM.cell(rm, 2, "   " + d).font = f8
            wsM.cell(rm, 4, cnt).font = f8
            x = wsM.cell(rm, 5, round(v, 4)); x.font = f8; x.number_format = N2
        x = wsM.cell(r0 - 1, 6, f"=SUM(E{r0}:E{rm})" if rm >= r0 else 0); x.font = f9b; x.number_format = N2
        TOT_M[clave] = f"{SM}!$F${r0 - 1}"
        rm += 2
    for c, w in zip("ABCDEF", (10, 70, 6, 9, 12, 13)): wsM.column_dimensions[c].width = w

    # ================================================================ ANÁLISIS DE PRECIOS UNITARIOS
    wsA["A1"] = "Análisis de precios unitarios"; wsA["A1"].font = f12b
    wsA["A3"] = "Presupuesto"; wsA["B3"] = f"0101001   {OBRA}"
    wsA["A4"] = "Subpresupuesto"; wsA["B4"] = "001   ESTRUCTURAS"; wsA["F4"] = "Fecha presupuesto"; wsA["G4"] = HOY
    for c in ("A3", "B3", "A4", "B4", "F4", "G4"): wsA[c].font = f9
    ra = 6
    CU = {}                   # clave -> celda costo unitario
    FILA_CANT_OBRA = []       # (fila, clave) para la columna H (cantidad en obra)
    for n in P:
        if n[0] != "P":
            continue
        _, it, desc, apu, _, clave = n
        wsA.cell(ra, 1, "Partida").font = f9b; wsA.cell(ra, 2, f"{it}      {desc}").font = f9b
        for c in range(1, 8): wsA.cell(ra, c).border = linea_ar
        ra += 1
        wsA.cell(ra, 1, "Rendimiento").font = f9
        wsA.cell(ra, 2, f"{apu['und']}/DIA").font = f9
        wsA.cell(ra, 3, "MO.").font = f9
        x = wsA.cell(ra, 4, apu["rend"]); x.font = fin9; x.number_format = N4
        wsA.cell(ra, 5, f'="EQ.  "&TEXT(D{ra},"0.0000")').font = f9
        wsA.cell(ra, 6, "Costo unitario directo por : " + apu["und"]).font = f9
        rend = f"$D${ra}"
        fila_cu = ra
        ra += 1
        for c, t in enumerate(["Código", "Descripción Recurso", "Unidad", "Cuadrilla", "Cantidad", "Precio S/.", "Parcial S/.", "Cant. obra"], 1):
            x = wsA.cell(ra, c, t); x.font = f8b; x.border = linea_ab; x.alignment = CEN
        ra += 1
        subt = {}
        for sec, nom, items in (("MO", "Mano de Obra", apu["mo"]), ("MAT", "Materiales", apu["mat"]),
                                ("EQ", "Equipos", apu["eq"] + ([(HER, None)] if apu["mo"] else [])), ("SC", "Subcontratos", apu.get("sc", []))):
            if not items:
                continue
            wsA.cell(ra, 2, nom).font = f9b
            fila_sec = ra; ra += 1; r_ini = ra
            for cod, q in items:
                d, u, _ = INS_D[cod]
                wsA.cell(ra, 1, cod).font = f8; wsA.cell(ra, 2, d).font = f8; wsA.cell(ra, 3, u).font = f8
                if cod == HER:
                    x = wsA.cell(ra, 5, f"={PAR['her']}*100"); x.number_format = N4
                    x = wsA.cell(ra, 6, f"=G{subt['MO']}"); x.number_format = N2
                    x = wsA.cell(ra, 7, f"=ROUND(E{ra}/100*F{ra},2)"); x.number_format = N2
                elif sec in ("MO", "EQ") and q is not None:
                    x = wsA.cell(ra, 4, q); x.font = fin9; x.number_format = N4
                    x = wsA.cell(ra, 5, f"=ROUND(D{ra}*8/{rend},4)"); x.number_format = N4
                else:
                    x = wsA.cell(ra, 5, 1.0 if q is None else q); x.font = fin9; x.number_format = N4
                if cod != HER:
                    x = wsA.cell(ra, 6, f"=INDEX({SI}!$E:$E,MATCH(A{ra},{SI}!$A:$A,0))"); x.number_format = N2
                    x = wsA.cell(ra, 7, f"=ROUND(E{ra}*F{ra},2)"); x.number_format = N2
                for c in (4, 5, 6, 7):
                    if wsA.cell(ra, c).font != fin9: wsA.cell(ra, c).font = f8
                FILA_CANT_OBRA.append((ra, clave, cod == HER))
                ra += 1
            x = wsA.cell(fila_sec, 7, f"=SUM(G{r_ini}:G{ra - 1})"); x.font = f9b; x.number_format = N2
            subt[sec] = fila_sec
            ra += 0
        x = wsA.cell(fila_cu, 7, "=" + "+".join(f"G{v}" for v in subt.values())); x.font = f9b; x.number_format = N2
        CU[clave] = f"{SA}!$G${fila_cu}"
        ra += 1
    for c, w in zip("ABCDEFGH", (12, 50, 7, 10, 11, 11, 12, 13)): wsA.column_dimensions[c].width = w

    # ================================================================ PRESUPUESTO
    ws = wsP
    ws["A1"] = "S10"; ws["A1"].font = f8; ws["F1"] = "Página 1"; ws["F1"].font = f8; ws["F1"].alignment = DER
    ws["A3"] = "Presupuesto"; ws["A3"].font = f12b; ws.merge_cells("A3:F3"); ws["A3"].alignment = CEN
    ws["A5"] = "Presupuesto"; ws["B5"] = f"0101001      {OBRA}"
    ws["A6"] = "Subpresupuesto"; ws["B6"] = "001      ESTRUCTURAS"
    ws["A7"] = "Cliente"; ws["B7"] = CLIENTE; ws["E7"] = "Costo al"; ws["F7"] = HOY
    ws["A8"] = "Lugar"; ws["B8"] = LUGAR
    for c in ("A5", "A6", "A7", "A8", "E7"): ws[c].font = f9b
    for c in ("B5", "B6", "B7", "B8", "F7"): ws[c].font = f9
    hdr = 10
    for c, t in enumerate(["Item", "Descripción", "Und.", "Metrado", "Precio S/.", "Parcial S/."], 1):
        x = ws.cell(hdr, c, t); x.font = f9b; x.border = caja; x.alignment = CEN
    rp = hdr + 1
    MET = {}      # clave -> celda de metrado en el presupuesto
    FILA = {}     # item -> fila
    pend = []     # (fila, fórmula con claves) para resolver luego
    titulos = []
    for n in P:
        if n[0] == "T":
            _, it, desc = n
            ws.cell(rp, 1, it).font = f9b; ws.cell(rp, 2, desc).font = f9b
            FILA[it] = rp; titulos.append((it, rp))
        else:
            _, it, desc, apu, fuente, clave = n
            ws.cell(rp, 1, it).font = f9; ws.cell(rp, 2, desc).font = f9; ws.cell(rp, 3, apu["und"]).font = f9
            ws.cell(rp, 3).alignment = CEN
            if fuente[0] == "hm":
                x = ws.cell(rp, 4, f"=ROUND({TOT_M[clave]},2)")
            else:
                pend.append((rp, fuente[1])); x = ws.cell(rp, 4)
            x.number_format = N2; x.font = f9
            x = ws.cell(rp, 5, f"=ROUND({CU[clave]},2)"); x.number_format = N2; x.font = f9
            x = ws.cell(rp, 6, f"=ROUND(D{rp}*E{rp},2)"); x.number_format = N2; x.font = f9
            MET[clave] = f"$D${rp}"; FILA[it] = rp
        rp += 1
    # fórmulas de metrado que dependen de otras partidas
    alias = {"SOLADO": MET["sol"], "EXM": MET["exm"], "EXZ": MET["exz"],
             "AREA_ALI": MET["e_ALI"], "FPISO": MET["fpi"], "ENC_TOTAL": "+".join(v for k, v in MET.items() if k.startswith("e_")), "VOL_ALB": f"{round(sum(f['vol'] for f in filas if f['g'] == 'ALB'), 4)}"}
    alias.update({k: v for k, v in MET.items()})
    for fila, fx in pend:
        for k in sorted(alias, key=len, reverse=True):
            fx = re.sub(r"\b" + re.escape(k) + r"\b", lambda _m, v=alias[k]: v, fx)
        x = ws.cell(fila, 4, fx if fx.startswith("=") else "=" + fx)
        if "ROUND" not in fx:
            x.value = "=ROUND(" + x.value[1:] + ",2)"
        x.number_format = N2; x.font = f9
    ws.cell(FILA["05.02"], 4).comment = Comment(f"Volumen de albañilería armada del modelo: {alias['VOL_ALB']} m3 x cuantía (Supuestos).", "NEXO")
    # parciales de títulos (suma de partidas cuyo ítem empieza con el del título y son hojas)
    hojas = [(n[1], FILA[n[1]]) for n in P if n[0] == "P"]
    for it, fila in titulos:
        filas_h = [f for i, f in hojas if i.startswith(it + ".")]
        x = ws.cell(fila, 6, "=" + "+".join(f"F{f}" for f in filas_h)); x.font = f9b; x.number_format = N2
    caps = [f for it, f in titulos if len(it) == 2]
    rp += 1
    pie = [("Costo Directo", "=" + "+".join(f"F{f}" for f in caps)),
           (None, None),
           ("GASTOS GENERALES ({:.0%})", f"=ROUND(F{{cd}}*{PAR['gg']},2)"),
           ("UTILIDAD ({:.0%})", f"=ROUND(F{{cd}}*{PAR['ut']},2)"),
           ("SUBTOTAL", "=F{cd}+F{gg}+F{ut}"),
           ("IGV ({:.0%})", f"=ROUND(F{{st}}*{PAR['igv']},2)"),
           ("TOTAL PRESUPUESTO", "=F{st}+F{ig}")]
    pos = {}
    nombres = ["cd", None, "gg", "ut", "st", "ig", "tot"]
    for (lbl, fx), nm in zip(pie, nombres):
        if lbl is None:
            rp += 1; continue
        pos[nm] = rp; rp += 1
    pct = {"gg": 0.10, "ut": 0.08, "ig": 0.18}
    for (lbl, fx), nm in zip(pie, nombres):
        if lbl is None:
            continue
        f_ = pos[nm]
        clave_pct = {"gg": "gg", "ut": "ut", "ig": "igv"}.get(nm)
        txt = (f'="{lbl.split(" (")[0]} ("&TEXT({PAR[clave_pct]}*100,"0")&"%)"' if clave_pct else lbl)
        x = ws.cell(f_, 2, txt); x.font = f9b
        x = ws.cell(f_, 6, fx.format(**pos)); x.font = f9b; x.number_format = N2
        if nm in ("cd", "st", "tot"):
            ws.cell(f_, 6).border = linea_ar
    ws.cell(pos["gg"], 2).comment = Comment("Los porcentajes se cambian en 'Supuestos y Fuentes'.", "NEXO")
    son_fila = pos["tot"] + 2
    ws.cell(son_fila, 1, "SON :").font = f9b
    ws.cell(son_fila, 2, "__SON__").font = f9
    for c, w in zip("ABCDEF", (10, 62, 6, 12, 11, 14)): ws.column_dimensions[c].width = w
    ws.print_title_rows = f"{hdr}:{hdr}"
    ws.page_setup.orientation = "portrait"; ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.freeze_panes = f"A{hdr + 1}"

    # cantidades en obra para la hoja de insumos (columna H del APU)
    for fila, clave, es_her in FILA_CANT_OBRA:
        met = f"Presupuesto!{MET[clave]}"
        x = wsA.cell(fila, 8, f"=G{fila}*{met}" if es_her else f"=E{fila}*{met}")
        x.number_format = N2; x.font = fgr8
    wsA.cell(8, 8).comment = Comment("Cant. obra = cantidad unitaria x metrado de la partida (para la hoja de recursos). En herramientas es el monto S/.", "NEXO")

    # ================================================================ PRECIOS Y CANTIDADES DE RECURSOS
    wsI["A1"] = "Precios y cantidades de recursos requeridos"; wsI["A1"].font = f12b
    wsI["A3"] = "Obra"; wsI["B3"] = f"0101001   {OBRA}"; wsI["A4"] = "Fecha"; wsI["B4"] = HOY; wsI["A5"] = "Lugar"; wsI["B5"] = LUGAR
    for c in ("A3", "B3", "A4", "B4", "A5", "B5"): wsI[c].font = f9
    for c, t in enumerate(["Código", "Recurso", "Unidad", "Cantidad", "Precio S/.", "Parcial S/.", "Fuente / observación"], 1):
        x = wsI.cell(7, c, t); x.font = f9b; x.border = caja; x.alignment = CEN
    ri = 8
    tot_sec = []
    for sec, nom in (("MO", "MANO DE OBRA"), ("MAT", "MATERIALES"), ("EQ", "EQUIPOS"), ("SC", "SUBCONTRATOS")):
        wsI.cell(ri, 1, nom).font = f9b; fs = ri; ri += 1; r0 = ri
        for cod, d, u, t, pr, fu in INS:
            if t != sec:
                continue
            wsI.cell(ri, 1, cod).font = f8; wsI.cell(ri, 2, d).font = f8; wsI.cell(ri, 3, u).font = f8
            if cod == HER:
                wsI.cell(ri, 4, "").font = f8
                x = wsI.cell(ri, 6, f"=ROUND(SUMIF({SA}!$A:$A,A{ri},{SA}!$H:$H),2)")
            else:
                x = wsI.cell(ri, 4, f"=SUMIF({SA}!$A:$A,A{ri},{SA}!$H:$H)"); x.number_format = N4; x.font = f8
                if pr is None:
                    x = wsI.cell(ri, 5, f"={MO_HH[cod]}"); x.font = f8
                else:
                    x = wsI.cell(ri, 5, pr); x.font = Font(name=F, size=8, color="0000FF")
                x.number_format = N2
                x = wsI.cell(ri, 6, f"=ROUND(D{ri}*E{ri},2)")
            x.number_format = N2; x.font = f8
            wsI.cell(ri, 7, fu).font = fgr8
            ri += 1
        x = wsI.cell(fs, 6, f"=SUM(F{r0}:F{ri - 1})"); x.font = f9b; x.number_format = N2
        tot_sec.append(fs); ri += 1
    wsI.cell(ri, 2, "Total (debe coincidir con el costo directo, salvo redondeos)").font = f9b
    x = wsI.cell(ri, 6, "=" + "+".join(f"F{f}" for f in tot_sec)); x.font = f9b; x.number_format = N2; x.border = linea_ar
    wsI.cell(ri + 1, 2, "Diferencia con el costo directo del presupuesto").font = f8
    x = wsI.cell(ri + 1, 6, f"=F{ri}-Presupuesto!F{pos['cd']}"); x.number_format = N2; x.font = f8
    for c, w in zip("ABCDEFG", (12, 50, 7, 13, 10, 13, 70)): wsI.column_dimensions[c].width = w

    for w in (wsP, wsA, wsI, wsM, wsS):
        w.sheet_view.showGridLines = False
    wb.calculation.fullCalcOnLoad = True
    wb.save(salida)
    return pos, son_fila


if __name__ == "__main__":
    pos, son = construir(sys.argv[1], sys.argv[2])
    print("Guardado", sys.argv[2], "| total en fila", pos["tot"], "| SON en fila", son)
