"""NEXO OBRA (NO) · Sectorización, cronograma de ejecución, cronograma valorizado y datos para la app.

Entradas:  metrados_<modelo>.json (Revit)  +  Presupuesto_..._v2.xlsx (presupuesto S10 ya recalculado)
Salidas:   Cronograma_Valorizado_<modelo>.xlsx · nexo_no_datos.json · sectorizacion_<paquete>.svg

Calendario: LUNES A SÁBADO (el domingo y los feriados no se trabajan).
Uso: python cronograma_nexo.py <metrados.json> <presupuesto_v2.xlsx> <carpeta_salida>
"""
import json
import math
import re
import sys
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName

import metrados as M

# ------------------------------------------------------------------ parámetros de programación (editables en el Excel)
INICIO_OBRA = date(2027, 1, 4)          # lunes
FERIADOS = [  # Perú 2027 + Arequipa + Día del trabajador de construcción civil (verificar con el calendario oficial)
    ("2027-01-01", "Año Nuevo"), ("2027-03-25", "Jueves Santo"), ("2027-03-26", "Viernes Santo"), ("2027-05-01", "Día del Trabajo"),
    ("2027-06-07", "Batalla de Arica y Día de la Bandera"), ("2027-06-29", "San Pedro y San Pablo"), ("2027-07-23", "Día de la Fuerza Aérea"),
    ("2027-07-28", "Fiestas Patrias"), ("2027-07-29", "Fiestas Patrias"), ("2027-08-06", "Batalla de Junín"),
    ("2027-08-15", "Aniversario de Arequipa (local)"), ("2027-08-30", "Santa Rosa de Lima"), ("2027-10-08", "Combate de Angamos"),
    ("2027-10-25", "Día del Trabajador de Construcción Civil"), ("2027-11-01", "Todos los Santos"), ("2027-12-08", "Inmaculada Concepción"),
    ("2027-12-09", "Batalla de Ayacucho"), ("2027-12-25", "Navidad"),
    ("2028-01-01", "Año Nuevo"), ("2028-04-13", "Jueves Santo"), ("2028-04-14", "Viernes Santo"), ("2028-05-01", "Día del Trabajo"),
]
# Tipos de actividad: (descripción, N° de cuadrillas/equipos por defecto)
TIPOS = {
    "PRO": ("Obras provisionales y movilización (duración fija)", None),
    "EXC": ("Excavación masiva (excavadoras)", 1),
    "ANC": ("Anclajes: perforación e inyección (equipos)", 2),
    "TEN": ("Tensado de anclajes (duración fija)", None),
    "MUA": ("Muro anclado: acero, encofrado y concreto (cuadrillas)", 4),
    "MAN": ("Excavación manual, refine, solado y subzapatas (cuadrillas)", 8),
    "CIM": ("Cimentaciones: acero, encofrado y concreto (cuadrillas)", 4),
    "AV": ("Acero en elementos verticales (cuadrillas)", 6),
    "EV": ("Encofrado de elementos verticales (cuadrillas)", 8),
    "CV": ("Concreto en elementos verticales (cuadrillas de vaciado)", 1),
    "EH": ("Encofrado de vigas, losas y escaleras (cuadrillas)", 8),
    "AH": ("Acero horizontal y ladrillo de techo (cuadrillas)", 6),
    "CH": ("Concreto de vigas, losas y escaleras (cuadrillas de vaciado)", 1),
    "ALB": ("Albañilería armada (cuadrillas)", 3),
    "FPI": ("Afirmado y falso piso de sótanos (cuadrillas)", 2),
    "IMP": ("Impermeabilización de cisterna (duración fija)", None),
    "GLO": ("Actividad continua (toda la obra)", None),
}
DUR_FIJA = {"PRO": 12, "TEN": 3, "IMP": 6}
REND_ANCLAJE = 2.0       # anclajes por día por equipo
LAG_TENSADO = 7          # días de curado del muro antes de tensar
LAG_TECHO = 1            # día entre el vaciado de un techo y el inicio del siguiente nivel en el mismo sector
WEEKEND = "0000001"      # Excel WORKDAY.INTL: solo el domingo es no laborable

F_ = "Arial"
f8, f8b, f9, f9b, f12b = (Font(name=F_, size=8), Font(name=F_, size=8, bold=True), Font(name=F_, size=9), Font(name=F_, size=9, bold=True),
                          Font(name=F_, size=12, bold=True))
fin9, fin8 = Font(name=F_, size=9, color="0000FF"), Font(name=F_, size=8, color="0000FF")
fgr8 = Font(name=F_, size=8, color="808080")
lin = Side(style="thin"); caja = Border(top=lin, bottom=lin, left=lin, right=lin)
N2, N4, N0 = '#,##0.00;-#,##0.00;"-"', '#,##0.0000;-#,##0.0000;"-"', '#,##0;-#,##0;"-"'
FD = "dd/mm/yyyy"
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)


# ------------------------------------------------------------------ calendario (igual que WORKDAY.INTL "0000001")
FER_SET = {date.fromisoformat(d) for d, _ in FERIADOS}


def laborable(d):
    return d.weekday() != 6 and d not in FER_SET


def workday(d, n):
    """WORKDAY.INTL(d, n): n días laborables después de d (n=0 devuelve d)."""
    while n > 0:
        d += timedelta(days=1)
        if laborable(d):
            n -= 1
    return d


def networkdays(a, b):
    if a > b:
        return 0
    n, d = 0, a
    while d <= b:
        n += laborable(d); d += timedelta(days=1)
    return n


# ------------------------------------------------------------------ lectura del presupuesto v2
def leer_presupuesto(ruta):
    wb = load_workbook(ruta, data_only=True)
    p = wb["Presupuesto"]
    part, cap = [], {}
    for r in range(11, p.max_row + 1):
        it, de, un, me, pu = (p.cell(r, c).value for c in range(1, 6))
        if not it or not isinstance(it, str) or not re.match(r"^\d{2}(\.\d{2})*$", it):
            continue
        if un:
            part.append({"item": it, "desc": de, "und": un, "met": float(me or 0), "pu": float(pu or 0)})
        else:
            cap[it] = de
    pie = {}
    for r in range(11, p.max_row + 1):
        b = p.cell(r, 2).value
        if isinstance(b, str):
            for k in ("GASTOS GENERALES", "UTILIDAD", "IGV"):
                if b.startswith(k):
                    m = re.search(r"\((\d+)%\)", b); pie[k] = int(m.group(1)) / 100 if m else None
    a = wb["Análisis de Precios Unitarios"]
    apu, act = {}, None
    for r in range(1, a.max_row + 1):
        A, B = a.cell(r, 1).value, a.cell(r, 2).value
        if A == "Partida":
            act = B.split()[0]; apu[act] = {"rend": None, "rec": {}}
        elif A == "Rendimiento" and act:
            apu[act]["rend"] = float(a.cell(r, 4).value or 0)
        elif act and isinstance(A, str) and re.match(r"^\d{10}$", A) and A != "0337010001":
            apu[act]["rec"][A] = (a.cell(r, 2).value, a.cell(r, 3).value, float(a.cell(r, 5).value or 0))
    return part, cap, pie, apu


# ------------------------------------------------------------------ paquetes de vaciado y sectores
VERT, HORI = {"COL", "PLA", "MUR", "ALB", "ESC", "RAM"}, {"VIG", "VCH", "LMA", "ALI"}
NOMBRE_NIVEL = {"00.01. SOTANO 2": "Sótano 2", "00.02. SOTANO 1": "Sótano 1", "00.03. SEMISOTANO": "Semisótano",
                "08. TECHO": "Azotea"}


def nombre_nivel(n):
    if n in NOMBRE_NIVEL:
        return NOMBRE_NIVEL[n]
    m = re.match(r"^0?(\d+)\. NFP", n)
    return f"Piso {int(m.group(1))}" if m else n


def sectorizar(D, filas):
    niv = sorted(D["niveles"], key=lambda l: l["cota"])
    nombres = [l["nombre"] for l in niv]
    cota = {l["nombre"]: l["cota"] for l in niv}

    def nivel_de(f):
        if f["nivel"] in cota:
            return f["nivel"]
        if not f["caja"]:
            return "07. NFP"
        z = f["caja"][2]
        c = [n for n in nombres if cota[n] <= z + 0.30]
        return c[-1] if c else nombres[0]

    for f in filas:
        f["niv"] = nivel_de(f)
        i = nombres.index(f["niv"])
        anclaje = f["g"] == "MUR" and "ANCLAJE" in (f["tipo"] or "").upper()
        if anclaje or f["g"] == "CIMM":
            # anillo 1 = semisótano, anillo 2 = sótano 1 (incluye cimientos de los muros anclados)
            f["paq"] = "AN1" if f["niv"] == "00.03. SEMISOTANO" else "AN2"
        elif f["g"] in ("ZAP", "LOSC", "SUBZ"):
            f["paq"] = "CIM"
        elif f["g"] in VERT:
            f["paq"] = f"T{min(i + 1, len(nombres) - 1):02d}"
        elif f["g"] in HORI:
            f["paq"] = f"T{max(i, 1):02d}"
        else:
            f["paq"] = "OTR"
    paquetes = {}
    for j in range(1, len(nombres)):
        paquetes[f"T{j:02d}"] = {"id": f"T{j:02d}", "nombre": f"Techo de {nombre_nivel(nombres[j - 1])}",
                                 "vertical": nombres[j - 1], "horizontal": nombres[j]}
    paquetes["CIM"] = {"id": "CIM", "nombre": "Cimentaciones", "vertical": None, "horizontal": None}
    # corte de sectores A/B: por el eje más largo, equilibrando el concreto
    for pid, pq in paquetes.items():
        s = [f for f in filas if f["paq"] == pid and f["caja"]]
        if not s:
            pq["sectores"] = []; continue
        x0, y0 = min(f["caja"][0] for f in s), min(f["caja"][1] for f in s)
        x1, y1 = max(f["caja"][3] for f in s), max(f["caja"][4] for f in s)
        eje = 0 if (x1 - x0) >= (y1 - y0) else 1
        cen = sorted(s, key=lambda f: (f["caja"][eje] + f["caja"][eje + 3]) / 2)
        tot = sum(f["conc"] for f in cen); acc = 0; corte = None
        for f in cen:
            acc += f["conc"]
            if acc >= tot / 2:
                corte = (f["caja"][eje] + f["caja"][eje + 3]) / 2 + 1e-6; break
        pq.update({"eje": "X" if eje == 0 else "Y", "corte": round(corte, 3), "caja": (x0, y0, x1, y1), "sectores": ["A", "B"]})
        for f in filas:
            if f["paq"] == pid:
                if not f["caja"]:
                    f["sec"] = "A" if f["id"] % 2 else "B"     # filas sin geometría: mitad y mitad
                else:
                    c = (f["caja"][eje] + f["caja"][eje + 3]) / 2
                    f["sec"] = "A" if c <= corte else "B"
    for f in filas:
        f.setdefault("sec", "")
    bov = [f for f in filas if f["id"] == -1]
    for f in bov:  # bovedillas sin geometría: mitad a cada sector
        f["sec"] = "AB"
    return paquetes, nombres


# ------------------------------------------------------------------ partidas del presupuesto <-> grupos del modelo
def nom_grupo(desc):
    return desc.split(" EN ", 1)[1] if " EN " in desc else ""


def _grupos():
    import presupuesto_s10 as PS
    g = {v[0]: k for k, v in PS.GRUPOS.items()}
    g["ALBAÑILERIA ARMADA"] = "ALB"
    return g


def clasifica(p):
    """('conc', grupo, f'c) · ('enc', grupo) · ('ace', grupo) · ('lad', 'ALI') · ('area', 'ALB') · ('esp', None)."""
    gp = _grupos()
    d = p["desc"].upper()
    if d.startswith("CONCRETO F'C=") and "SUB-ZAPATA" not in d:
        return ("conc", gp.get(nom_grupo(p["desc"])), int(re.search(r"F'C=(\d+)", d).group(1)))
    if d.startswith("ENCOFRADO"):
        return ("enc", gp.get(nom_grupo(p["desc"])), None)
    if d.startswith("ACERO DE REFUERZO"):
        return ("ace", gp.get(nom_grupo(p["desc"])), None)
    if d.startswith("LADRILLO DE TECHO"):
        return ("lad", "ALI", None)
    if d.startswith("MURO DE ALBAÑILERIA"):
        return ("area", "ALB", None)
    return ("esp", None, None)


def act_elem(f, clase):
    """Actividad que ejecuta la partida 'clase' (conc/enc/ace/lad/area) del elemento f."""
    if f["paq"] in ("AN1", "AN2"):
        return "MUA1" if f["paq"] == "AN1" else "MUA2"
    if f["paq"] == "CIM":
        return f"CIM-{f['sec']}"
    if f["g"] == "ALB":
        return f"{f['paq']}-ALB-{f['sec']}"
    vert = f["g"] in VERT and f["g"] not in ("ESC", "RAM")
    t = {"conc": "CV" if vert else "CH", "enc": "EV" if vert else "EH", "ace": "AV" if vert else "AH", "lad": "AH", "area": "AV"}[clase]
    return f"{f['paq']}-{t}-{f['sec']}"


def parametros_revit(ruta, D, filas, acts, part, inicio):
    """Valores VZ_ por ElementId para el Dynamo Cargar_Parametros_NEXO_NO.dyn (cuadran con el presupuesto)."""
    item, tot = {}, defaultdict(float)
    P = {p["item"]: p for p in part}
    for p in part:
        c, g, fc = clasifica(p)
        if c != "esp":
            item[(c, g, fc if c == "conc" else None)] = p["item"]
        elif "SUB-ZAPATAS" in p["desc"].upper():
            item[("conc", "SUBZ", None)] = p["item"]
    for f in filas:   # bases de reparto del presupuesto por grupo
        tot[("conc", f["g"], f["fc"])] += f["conc"]; tot[("enc", f["g"], None)] += (f["area"] if f["g"] == "ALI" else f["enc"])
        tot[("ace", f["g"], None)] += f["conc"] if f["g"] != "ALB" else f["vol"]
        tot[("lad", f["g"], None)] += f["area"]; tot[("area", f["g"], None)] += f["area"]

    def parte(clase, g, fc, q):
        it = item.get((clase, g, fc))
        if not it or tot[(clase, g, fc)] <= 0 or q <= 0:
            return it, 0.0, 0.0
        met = P[it]["met"] * q / tot[(clase, g, fc)]
        return it, met, met * P[it]["pu"]

    fmt = lambda d: d.strftime("%d/%m/%Y")
    out = []
    for f in filas:
        if f["id"] < 0:
            continue
        g = f["g"]
        if g == "ALB":
            it_c, q_c, m_c = parte("area", g, None, f["area"])
        else:
            it_c, q_c, m_c = parte("conc", g, f["fc"], f["conc"])
        _, q_e, m_e = parte("enc", g, None, f["area"] if g == "ALI" else f["enc"])
        _, q_a, m_a = parte("ace", g, None, f["conc"] if g != "ALB" else f["vol"])
        it_e, _, _ = parte("enc", g, None, f["area"] if g == "ALI" else f["enc"])
        it_a, _, _ = parte("ace", g, None, f["conc"] if g != "ALB" else f["vol"])
        partidas = [{"item": it_c, "clase": "muro" if g == "ALB" else "concreto", "metrado": round(q_c, 4)},
                    {"item": it_e, "clase": "encofrado", "metrado": round(q_e, 4)},
                    {"item": it_a, "clase": "acero", "metrado": round(q_a, 4)}]
        if g == "ALI":
            it_l, q_l, m_l = parte("lad", g, None, f["area"]); m_a += m_l      # el ladrillo se coloca con el acero de techo
            partidas.append({"item": it_l, "clase": "ladrillo", "metrado": round(q_l, 4)})
        partidas = [p for p in partidas if p["item"] and p["metrado"] > 0]
        s = f["sec"] if f["sec"] in ("A", "B") else "A"
        ff = {**f, "sec": s}
        a_c, a_e, a_a = act_elem(ff, "area" if g == "ALB" else "conc"), act_elem(ff, "enc"), act_elem(ff, "ace")
        if a_c not in acts:
            continue
        out.append({
            "id": f["id"], "paquete": f["paq"], "sector": s, "partida": it_c or "", "actividad": a_c,
            "act_acero": a_a if a_a in acts else a_c, "act_encofrado": a_e if a_e in acts else a_c,
            "fecha_prog": fmt(acts[a_c]["fin"]),
            "concreto_m3": round(q_c, 4) if g != "ALB" else 0.0, "encofrado_m2": round(q_e, 4), "acero_kg": round(q_a, 2),
            "monto_concreto": round(m_c, 2), "monto_encofrado": round(m_e, 2), "monto_acero": round(m_a, 2),
            "nivel": f["niv"], "tipo": f["tipo"], "grupo": g, "partidas": partidas,
        })
    datos = {"modelo": D["modelo"], "inicio_obra": inicio.isoformat(), "generado": date.today().isoformat(),
             "niveles": [{"nombre": l["nombre"], "cota": l["cota"]} for l in sorted(D["niveles"], key=lambda l: l["cota"])],
             "partidas": [{"item": p["item"], "desc": p["desc"], "und": p["und"], "metrado": p["met"], "pu": p["pu"]} for p in part],
             "capitulos": CAPITULOS, "pie": PIE,
             "nota": "Montos en costo directo (S/ sin IGV). Fechas dd/mm/aaaa. El monto de acero incluye el ladrillo de techo.",
             "elementos": out}
    Path(ruta).write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    return out


# ------------------------------------------------------------------ actividades y pesos de cada partida
def construir_actividades(D, filas, paquetes, part, X):
    acts, orden = {}, []

    def act(aid, nombre, tipo, paq="", sec="", preds=(), lag=0):
        acts[aid] = {"id": aid, "nombre": nombre, "tipo": tipo, "paq": paq, "sec": sec, "preds": list(preds), "lag": lag, "pesos": []}
        orden.append(aid)

    # ---------- actividades
    act("PRO", "Obras provisionales, cerco y movilización", "PRO")
    act("EXC1", "Excavación masiva capa 1 (hasta fondo de semisótano)", "EXC", "AN1", preds=["PRO"])
    act("ANC1", "Anclajes fila 1: perforación e inyección", "ANC", "AN1", preds=["EXC1"])
    act("MUA1", "Muros anclados del semisótano (paños alternados)", "MUA", "AN1", preds=["ANC1"])
    act("TEN1", "Tensado de anclajes fila 1", "TEN", "AN1", preds=["MUA1"], lag=LAG_TENSADO)
    act("EXC2", "Excavación masiva capa 2 (hasta fondo de sótano 1)", "EXC", "AN2", preds=["TEN1"])
    act("ANC2", "Anclajes fila 2: perforación e inyección", "ANC", "AN2", preds=["EXC2"])
    act("MUA2", "Muros anclados del sótano 1 y sus cimientos", "MUA", "AN2", preds=["ANC2"])
    act("TEN2", "Tensado de anclajes fila 2", "TEN", "AN2", preds=["MUA2"], lag=LAG_TENSADO)
    capas = X["capas"]
    if len(capas) >= 3:
        act("EXC3", "Excavación masiva capa 3 (sótano 2)", "EXC", "CIM", preds=["TEN2"])
        pre_cim = "EXC3"
    else:
        pre_cim = "TEN2"
    act("MAN-A", "Excavación manual, refine, solado y subzapatas · sector A", "MAN", "CIM", "A", preds=[pre_cim])
    act("MAN-B", "Excavación manual, refine, solado y subzapatas · sector B", "MAN", "CIM", "B", preds=[pre_cim, "MAN-A"])
    act("CIM-A", "Zapatas, losas de cimentación y cisterna · sector A", "CIM", "CIM", "A", preds=["MAN-A"])
    act("CIM-B", "Zapatas, losas de cimentación y cisterna · sector B", "CIM", "CIM", "B", preds=["MAN-B", "CIM-A"])
    prev = {"A": "CIM-A", "B": "CIM-B"}
    tren = ["AV", "EV", "CV", "EH", "AH", "CH"]
    paqs = sorted(k for k in paquetes if k.startswith("T"))
    ultimo_paq = {}
    for pid in paqs:
        pq = paquetes[pid]
        if not pq.get("sectores"):
            continue
        for s in ("A", "B"):
            ant = None
            for t in tren:
                aid = f"{pid}-{t}-{s}"
                nm = {"AV": "Acero vertical", "EV": "Encofrado vertical", "CV": "Concreto vertical", "EH": "Encofrado de techo",
                      "AH": "Acero de techo y ladrillo", "CH": "Concreto de techo"}[t]
                preds = []
                if ant:
                    preds.append(ant)
                else:
                    preds.append(prev[s])
                if s == "B":
                    preds.append(f"{pid}-{t}-A")
                lag = LAG_TECHO if (not ant) else 0
                act(aid, f"{pq['nombre']} · {nm} · sector {s}", t, pid, s, preds, lag)
                ant = aid
            prev[s] = ant
            if any(f["g"] == "ALB" and f["paq"] == pid and f["sec"] == s for f in filas):
                act(f"{pid}-ALB-{s}", f"{pq['nombre']} · Albañilería armada · sector {s}", "ALB", pid, s,
                    [f"{pid}-CV-{s}"] + ([f"{pid}-ALB-A"] if s == "B" else []))
            ultimo_paq[pid] = ant
    t1 = paqs[0]; t3 = paqs[min(2, len(paqs) - 1)]
    act("IMP", "Impermeabilización y waterstop de cisterna", "IMP", "CIM", preds=[f"{t1}-CH-B"])
    act("FPI", "Afirmado y falso piso de sótanos", "FPI", "CIM", preds=[f"{t3}-CH-B"])
    act("TRZ", "Trazo, niveles y replanteo (continuo)", "GLO")
    act("SST", "Seguridad y salud en el trabajo (continuo, mensual)", "GLO")

    # ---------- pesos: cada partida se reparte entre actividades según el modelo
    capa_vol = [c["area"] * max(0, (M.COTA_TERRENO if i == 0 else capas[i - 1]["fondo"]) - c["fondo"]) for i, c in enumerate(capas)]
    exc_ids = ["EXC1", "EXC2", "EXC3"][:len(capas)]
    vt = sum(capa_vol)

    def repartir(f, clase, cant, pesos):
        if f["sec"] == "AB":
            pesos[act_elem({**f, "sec": "A"}, clase)] += cant / 2; pesos[act_elem({**f, "sec": "B"}, clase)] += cant / 2
        else:
            pesos[act_elem(f, clase)] += cant

    asign = {}
    for p in part:
        it, d = p["item"], p["desc"].upper()
        pesos = defaultdict(float)
        clase, g, fc = clasifica(p)
        if clase in ("conc", "enc", "ace", "lad", "area"):
            for f in filas:
                if f["g"] != g or (clase == "conc" and f["fc"] != fc):
                    continue
                q = {"conc": f["conc"], "ace": f["conc"] if g != "ALB" else f["vol"], "enc": f["area"] if g == "ALI" else f["enc"],
                     "lad": f["area"], "area": f["area"]}[clase]
                if q > 0:
                    repartir(f, clase, q, pesos)
            if g == "ALB" and clase == "ace":
                for f in filas:
                    if f["g"] == "ALB":
                        pesos[act_elem(f, "area")] += f["vol"]
                pesos = defaultdict(float, {k: v for k, v in pesos.items() if "-ALB-" in k})
        elif "TRAZO" in d:
            pesos["TRZ"] = 1
        elif "SEGURIDAD Y SALUD" in d:
            pesos["SST"] = 1
        elif it.startswith("01."):
            pesos["PRO"] = 1
        elif d.startswith("EXCAVACION MASIVA"):
            for i, a in enumerate(exc_ids):
                pesos[a] = capa_vol[i] / vt
        elif "EXCAVACION MANUAL" in d or "REFINE" in d or "SOLADO" in d or "SUB-ZAPATAS" in d:
            for f in filas:
                if f["paq"] in ("CIM", "AN2") and f["g"] in ("ZAP", "LOSC", "SUBZ", "CIMM"):
                    s = f["sec"] if f["paq"] == "CIM" else "A"
                    pesos[f"MAN-{s}"] += f["inf"]
        elif "ELIMINACION" in d:
            exz = next(q for q in part if q["desc"].upper().startswith("EXCAVACION MANUAL"))["met"]
            exm = next(q for q in part if q["desc"].upper().startswith("EXCAVACION MASIVA"))["met"]
            for i, a in enumerate(exc_ids):
                pesos[a] = exm * capa_vol[i] / vt
            pesos["MAN-A"] += exz / 2; pesos["MAN-B"] += exz / 2
        elif "ANCLAJES" in d:
            n1 = sum(x[2] for x in X["lin_anc"] if "SEMISOTANO" in x[0]); n2 = sum(x[2] for x in X["lin_anc"]) - n1
            pesos["ANC1"], pesos["TEN1"], pesos["ANC2"], pesos["TEN2"] = 0.8 * n1, 0.2 * n1, 0.8 * n2, 0.2 * n2
        elif "AFIRMADO" in d or "FALSO PISO" in d:
            pesos["FPI"] = 1
        elif "IMPERMEABILIZACION" in d:
            pesos["IMP"] = 1
        elif "CURADO" in d or "ENSAYOS" in d:
            # se reparte entre las actividades de concreto según su volumen
            for f in filas:
                if f["conc"] > 0 and f["g"] not in ("SUBZ",):
                    repartir(f, "conc", f["enc"] if "CURADO" in d else f["conc"], pesos)
        else:
            raise ValueError(f"Partida sin asignar: {it} {p['desc']}")
        tot = sum(pesos.values())
        if tot <= 0:
            raise ValueError(f"Partida sin cantidades en el modelo: {it} {p['desc']}")
        asign[it] = {k: v / tot for k, v in pesos.items() if v > 0}
        for k in asign[it]:
            if k not in acts:
                raise ValueError(f"Actividad {k} no existe (partida {it})")
    return acts, orden, asign


# ------------------------------------------------------------------ programación (réplica exacta de las fórmulas del Excel)
def programar(acts, orden, part, asign, apu, tipos_n):
    P = {p["item"]: p for p in part}
    for a in acts.values():
        a["cd"] = 0.0; a["monto"] = 0.0
    for it, pes in asign.items():
        p = P[it]; r = apu.get(it, {}).get("rend") or 0
        for aid, w in pes.items():
            met = round(p["met"] * w, 4)
            acts[aid]["monto"] += round(met * p["pu"], 2)
            if cuenta_para_duracion(it, p, acts[aid]["tipo"]) and r > 0:
                acts[aid]["cd"] += met / (REND_ANCLAJE if "ANCLAJES" in p["desc"].upper() else r)
    # quitar actividades vacías (sin monto y sin duración fija): se reemplazan por sus predecesoras
    vac = {aid for aid, a in acts.items() if a["monto"] <= 0.005 and a["tipo"] not in DUR_FIJA and a["tipo"] != "GLO"}

    def resolver(ps):
        out = []
        for x in ps:
            out += resolver(acts[x]["preds"]) if x in vac else [x]
        return list(dict.fromkeys(out))
    orden = [a for a in orden if a not in vac]
    for aid in orden:
        acts[aid]["preds"] = resolver(acts[aid]["preds"])
    for aid in vac:
        del acts[aid]
    for aid in orden:
        a = acts[aid]
        if a["tipo"] == "GLO":
            continue
        if a["tipo"] in DUR_FIJA:
            a["dur"] = DUR_FIJA[a["tipo"]]
        else:
            a["dur"] = max(1, math.ceil(round(a["cd"], 6) / tipos_n[a["tipo"]]))
        if a["preds"]:
            a["ini"] = workday(max(acts[p]["fin"] for p in a["preds"]), 1 + a["lag"])
        else:
            a["ini"] = workday(INICIO_OBRA - timedelta(days=1), 1)
        a["fin"] = workday(a["ini"], a["dur"] - 1)
    fin_obra = max(a["fin"] for a in acts.values() if a["tipo"] != "GLO")
    for a in acts.values():
        if a["tipo"] == "GLO":
            a["ini"] = workday(INICIO_OBRA - timedelta(days=1), 1); a["fin"] = fin_obra; a["dur"] = networkdays(a["ini"], a["fin"])
    return orden


def cuenta_para_duracion(it, p, tipo):
    d = p["desc"].upper()
    if "ELIMINACION" in d or "CURADO" in d or "ENSAYOS" in d or tipo in DUR_FIJA or tipo == "GLO":
        return False
    if "ANCLAJES" in d:
        return tipo == "ANC"
    return True


def meses(ini, fin):
    out, d = [], date(ini.year, ini.month, 1)
    while d <= fin:
        nx = date(d.year + (d.month == 12), d.month % 12 + 1, 1)
        out.append((d, nx - timedelta(days=1))); d = nx
    return out


# ------------------------------------------------------------------ SVG de sectorización
def svg_sector(paq, filas, ruta):
    s = [f for f in filas if f["paq"] == paq["id"] and f["caja"]]
    if not s:
        return
    x0, y0, x1, y1 = paq["caja"]; esc = 22; m = 30
    W, H = (x1 - x0) * esc + 2 * m, (y1 - y0) * esc + 2 * m + 40
    col = {"A": "#2563EB", "B": "#E5B869"}
    out = [f'<svg xmlns="http://www.w3.org/2000/svg" width="{W:.0f}" height="{H:.0f}" viewBox="0 0 {W:.0f} {H:.0f}" font-family="Arial">',
           f'<rect width="100%" height="100%" fill="#ffffff"/>',
           f'<text x="{m}" y="20" font-size="14" font-weight="bold">{paq["nombre"]} — sectorización A/B (corte {paq["eje"]} = {paq["corte"]:.2f} m)</text>']
    for f in sorted(s, key=lambda f: f["g"] in HORI):
        a, b, c, d = f["caja"][0], f["caja"][1], f["caja"][3], f["caja"][4]
        X_ = m + (a - x0) * esc; Y_ = 40 + m + (y1 - d) * esc
        op = 0.35 if f["g"] in HORI else 0.9
        out.append(f'<rect x="{X_:.1f}" y="{Y_:.1f}" width="{max(1, (c - a) * esc):.1f}" height="{max(1, (d - b) * esc):.1f}" '
                   f'fill="{col.get(f["sec"], "#999")}" fill-opacity="{op}" stroke="#333" stroke-width="0.4"/>')
    if paq["eje"] == "X":
        cx = m + (paq["corte"] - x0) * esc
        out.append(f'<line x1="{cx:.1f}" y1="{40 + m - 10}" x2="{cx:.1f}" y2="{H - m + 10:.1f}" stroke="#d00" stroke-width="2" stroke-dasharray="8 5"/>')
    else:
        cy = 40 + m + (y1 - paq["corte"]) * esc
        out.append(f'<line x1="{m - 10}" y1="{cy:.1f}" x2="{W - m + 10:.1f}" y2="{cy:.1f}" stroke="#d00" stroke-width="2" stroke-dasharray="8 5"/>')
    out.append(f'<text x="{m}" y="{H - 8:.0f}" font-size="11">Azul = sector A · Dorado = sector B · línea roja = junta de construcción propuesta</text></svg>')
    Path(ruta).write_text("\n".join(out), encoding="utf-8")


# ------------------------------------------------------------------ Excel
def excel(ruta, D, filas, paquetes, acts, orden, part, cap, pie, apu, asign, X):
    wb = Workbook()
    wsC = wb.active; wsC.title = "Cronograma"
    wsV = wb.create_sheet("Cronograma Valorizado")
    wsR = wb.create_sheet("Adquisiciones y Recursos")
    wsS = wb.create_sheet("Sectorización")
    wsM = wb.create_sheet("Metrados por Actividad")
    wsB = wb.create_sheet("Presupuesto Base")
    wsK = wb.create_sheet("Coeficientes APU")
    wsF = wb.create_sheet("Formato Valorización")
    wsP = wb.create_sheet("Parámetros")
    for w in wb.worksheets:
        w.sheet_view.showGridLines = False
    Q = lambda n: f"'{n}'"

    # ---------------- Parámetros
    wsP["A1"] = "PARÁMETROS DE PROGRAMACIÓN"; wsP["A1"].font = f12b
    wsP["A2"] = "Celdas en azul = editables. Calendario: lunes a sábado; domingo y feriados no laborables."; wsP["A2"].font = f8
    wsP["A4"] = "Inicio de obra"; wsP["A4"].font = f9b
    x = wsP["C4"]; x.value = INICIO_OBRA; x.number_format = FD; x.font = fin9
    wsP["A5"] = "Gastos generales / Utilidad / IGV (del presupuesto)"; wsP["A5"].font = f9b
    for c, k in zip("CDE", ("GASTOS GENERALES", "UTILIDAD", "IGV")):
        x = wsP[f"{c}5"]; x.value = pie.get(k) or 0; x.number_format = "0%"; x.font = fin9
    wsP["A7"] = "CUADRILLAS / EQUIPOS POR TIPO DE ACTIVIDAD"; wsP["A7"].font = f9b
    for c, t in enumerate(["Tipo", "Descripción", "N°"], 1):
        x = wsP.cell(8, c, t); x.font = f8b; x.border = caja
    r = 9
    for k, (d, n) in TIPOS.items():
        wsP.cell(r, 1, k).font = f9; wsP.cell(r, 2, d).font = f9
        x = wsP.cell(r, 3, n if n is not None else "-"); x.font = fin9 if n is not None else f9
        r += 1
    t_ini, t_fin = 9, r - 1
    r += 1
    wsP.cell(r, 1, "FERIADOS (no laborables)").font = f9b; r += 1
    f0 = r
    for d, n in FERIADOS:
        x = wsP.cell(r, 1, date.fromisoformat(d)); x.number_format = FD; x.font = fin9
        wsP.cell(r, 2, n).font = f9; r += 1
    for k in range(6):   # filas libres para más feriados
        wsP.cell(r, 1).font = fin9; wsP.cell(r, 1).number_format = FD; r += 1
    wb.defined_names["FER"] = DefinedName("FER", attr_text=f"'Parámetros'!$A${f0}:$A${r - 1}")
    wb.defined_names["INICIO"] = DefinedName("INICIO", attr_text="'Parámetros'!$C$4")
    wsP.cell(r + 1, 1, "Feriados 2027-2028 de Perú más el aniversario de Arequipa (15/08) y el Día del Trabajador de Construcción Civil "
                       "(25/10). Verificar con el calendario oficial y agregar días no laborables decretados.").font = f8
    wsP.cell(r + 2, 1, f"Rendimiento de anclajes: {REND_ANCLAJE} und/día por equipo. Tensado a {LAG_TENSADO} días de vaciado el muro. "
                       f"Entre el vaciado de un techo y el acero vertical del nivel siguiente: {LAG_TECHO} día.").font = f8
    for c, w in zip("ABC", (14, 62, 10)):
        wsP.column_dimensions[c].width = w

    # ---------------- Presupuesto base (valores del presupuesto v2)
    wsB["A1"] = "PRESUPUESTO BASE (copiado del presupuesto S10 v2)"; wsB["A1"].font = f12b
    for c, t in enumerate(["Item", "Descripción", "Und.", "Metrado", "Precio S/.", "Parcial S/.", "Metrado repartido", "Diferencia"], 1):
        x = wsB.cell(3, c, t); x.font = f8b; x.border = caja; x.alignment = CEN
    filaB = {}
    for i, p in enumerate(part):
        rr = 4 + i; filaB[p["item"]] = rr
        wsB.cell(rr, 1, p["item"]).font = f8; wsB.cell(rr, 2, p["desc"]).font = f8; wsB.cell(rr, 3, p["und"]).font = f8
        for c, v in ((4, p["met"]), (5, p["pu"])):
            x = wsB.cell(rr, c, v); x.number_format = N2; x.font = fin8
        x = wsB.cell(rr, 6, f"=ROUND(D{rr}*E{rr},2)"); x.number_format = N2; x.font = f8
    ub = 3 + len(part)
    rcd = ub + 1
    wsB.cell(rcd, 2, "Costo directo").font = f9b
    x = wsB.cell(rcd, 6, f"=SUM(F4:F{ub})"); x.number_format = N2; x.font = f9b
    for c, w in zip("ABCDEFGH", (9, 62, 6, 12, 11, 14, 14, 11)):
        wsB.column_dimensions[c].width = w

    # ---------------- Coeficientes APU (recursos por unidad de partida)
    recursos = {}
    for it, a in apu.items():
        for cod, (d, u, q) in a["rec"].items():
            recursos.setdefault(cod, (d, u))
    claves = [c for c in sorted(recursos) if not c.startswith("0410")]
    wsK["A1"] = "COEFICIENTES DE LOS ANÁLISIS DE PRECIOS UNITARIOS (cantidad de recurso por unidad de partida)"; wsK["A1"].font = f12b
    wsK.cell(3, 1, "Item").font = f8b; wsK.cell(3, 2, "Rend./día").font = f8b
    for j, cod in enumerate(claves):
        x = wsK.cell(2, 3 + j, cod); x.font = f8
        x = wsK.cell(3, 3 + j, f"{recursos[cod][0]} ({recursos[cod][1]})"); x.font = f8b; x.alignment = CEN
        wsK.column_dimensions[CL(3 + j)].width = 13
    wsK.row_dimensions[3].height = 60
    filaK = {}
    for i, p in enumerate(part):
        rr = 4 + i; filaK[p["item"]] = rr
        wsK.cell(rr, 1, p["item"]).font = f8
        a = apu.get(p["item"], {"rend": 0, "rec": {}})
        x = wsK.cell(rr, 2, a["rend"]); x.font = fin8; x.number_format = N2
        for j, cod in enumerate(claves):
            q = a["rec"].get(cod)
            if q:
                x = wsK.cell(rr, 3 + j, q[2]); x.font = fin8; x.number_format = N4
    wsK.column_dimensions["A"].width = 9
    uk = 3 + len(part)

    # ---------------- Cronograma
    hdr = 6
    wsC["A1"] = "CRONOGRAMA DE EJECUCIÓN — CASCO ESTRUCTURAL"; wsC["A1"].font = f12b
    wsC["A2"] = f"Modelo {D['modelo']} · Arequipa · calendario lunes a sábado · inicio = 'Parámetros'!C4"; wsC["A2"].font = f8
    wsC["A3"] = "Duración = cuadrilla-días (del modelo y los rendimientos del APU) / N° de cuadrillas. Inicio = día laborable siguiente al fin de las predecesoras + desfase."
    wsC["A3"].font = f8
    cols = ["ID", "Actividad", "Tipo", "Paquete", "Sector", "Cuadrilla-días", "N° cuadr.", "Duración (días)", "Predecesoras", "Desfase", "Inicio", "Fin",
            "Monto CD S/.", "% CD"]
    for c, t in enumerate(cols, 1):
        x = wsC.cell(hdr, c, t); x.font = f8b; x.border = caja; x.alignment = CEN
    fila = {aid: hdr + 1 + i for i, aid in enumerate(orden)}
    reg = [a for a in orden if acts[a]["tipo"] != "GLO"]
    ult_reg = max(fila[a] for a in reg)
    nM = len(asign_rows := [(aid, it, w) for it, pes in asign.items() for aid, w in pes.items() if aid in acts])
    met_ini, met_fin = 5, 4 + nM
    MQ = Q("Metrados por Actividad")
    for aid in orden:
        a, rr = acts[aid], fila[aid]
        for c, v in ((1, aid), (2, a["nombre"]), (3, a["tipo"]), (4, a["paq"]), (5, a["sec"])):
            wsC.cell(rr, c, v).font = f8
        x = wsC.cell(rr, 6, f"=SUMIF({MQ}!$A${met_ini}:$A${met_fin},A{rr},{MQ}!$J${met_ini}:$J${met_fin})"); x.number_format = N2; x.font = f8
        if a["tipo"] in DUR_FIJA:
            x = wsC.cell(rr, 8, DUR_FIJA[a["tipo"]]); x.font = fin8
            wsC.cell(rr, 7, "-").font = f8
        elif a["tipo"] == "GLO":
            wsC.cell(rr, 7, "-").font = f8
            x = wsC.cell(rr, 8, f'=_xlfn.NETWORKDAYS.INTL(K{rr},L{rr},"{WEEKEND}",FER)'); x.font = f8
        else:
            x = wsC.cell(rr, 7, f"=INDEX('Parámetros'!$C${t_ini}:$C${t_fin},MATCH(C{rr},'Parámetros'!$A${t_ini}:$A${t_fin},0))"); x.font = f8
            x = wsC.cell(rr, 8, f"=MAX(1,ROUNDUP(ROUND(F{rr},6)/G{rr},0))"); x.font = f8
        wsC.cell(rr, 9, ", ".join(a["preds"])).font = f8
        x = wsC.cell(rr, 10, a["lag"]); x.font = fin8
        if a["tipo"] == "GLO":
            ini = f'=_xlfn.WORKDAY.INTL(INICIO-1,1,"{WEEKEND}",FER)'
            fin = f"=MAX($L${hdr + 1}:$L${ult_reg})"
        else:
            if a["preds"]:
                mx = ",".join(f"L{fila[p]}" for p in a["preds"])
                ini = f'=_xlfn.WORKDAY.INTL(MAX({mx}),1+J{rr},"{WEEKEND}",FER)'
            else:
                ini = f'=_xlfn.WORKDAY.INTL(INICIO-1,1,"{WEEKEND}",FER)'
            fin = f'=_xlfn.WORKDAY.INTL(K{rr},H{rr}-1,"{WEEKEND}",FER)'
        x = wsC.cell(rr, 11, ini); x.number_format = FD; x.font = f8
        x = wsC.cell(rr, 12, fin); x.number_format = FD; x.font = f8
        x = wsC.cell(rr, 13, f"=SUMIF({MQ}!$A${met_ini}:$A${met_fin},A{rr},{MQ}!$H${met_ini}:$H${met_fin})"); x.number_format = N2; x.font = f8
        x = wsC.cell(rr, 14, f"=IF($M${hdr + len(orden) + 2}=0,0,M{rr}/$M${hdr + len(orden) + 2})"); x.number_format = "0.00%"; x.font = f8
    uC = hdr + len(orden)
    rt = uC + 2
    wsC.cell(rt, 2, "Total costo directo programado").font = f9b
    x = wsC.cell(rt, 13, f"=SUM(M{hdr + 1}:M{uC})"); x.number_format = N2; x.font = f9b
    wsC.cell(rt + 1, 2, "Diferencia con el presupuesto (debe ser ~0)").font = f8
    x = wsC.cell(rt + 1, 13, f"=M{rt}-{Q('Presupuesto Base')}!F{rcd}"); x.number_format = N2; x.font = f8
    wsC.cell(rt + 2, 2, "Plazo (días calendario)").font = f9b
    x = wsC.cell(rt + 2, 13, f"=MAX(L{hdr + 1}:L{uC})-MIN(K{hdr + 1}:K{uC})+1"); x.font = f9b
    wsC.cell(rt + 3, 2, "Fin de obra (casco)").font = f9b
    x = wsC.cell(rt + 3, 13, f"=MAX(L{hdr + 1}:L{uC})"); x.number_format = FD; x.font = f9b
    # Gantt semanal
    ini_obra = min(acts[a]["ini"] for a in orden); fin_obra = max(acts[a]["fin"] for a in orden)
    lunes = ini_obra - timedelta(days=ini_obra.weekday())
    nsem = (fin_obra - lunes).days // 7 + 3
    g0 = 16
    wsC.cell(hdr - 1, g0, "Semana (lunes)").font = f8b
    for k in range(nsem):
        c = g0 + k
        x = wsC.cell(hdr, c, f"=INICIO-WEEKDAY(INICIO,3)+{7 * k}" if k == 0 else f"={CL(c - 1)}{hdr}+7")
        x.number_format = "dd/mm"; x.font = Font(name=F_, size=7); x.alignment = Alignment(text_rotation=90, horizontal="center")
        wsC.column_dimensions[CL(c)].width = 2.6
        wsC.cell(hdr - 1, c).value = None
    wsC.cell(hdr - 2, g0, "S1 = primera semana").font = f8
    rng = f"{CL(g0)}{hdr + 1}:{CL(g0 + nsem - 1)}{uC}"
    cond = lambda tipos: f'AND(OR({",".join(f"$C{hdr + 1}=" + chr(34) + t + chr(34) for t in tipos)}),$K{hdr + 1}<={CL(g0)}${hdr}+6,$L{hdr + 1}>={CL(g0)}${hdr})'
    for tipos, color in ((["EXC", "ANC", "TEN", "MUA", "MAN", "FPI"], "A0522D"), (["CIM", "AV", "EV", "CV", "ALB"], "2563EB"),
                         (["EH", "AH", "CH"], "D4A14A"), (["PRO", "IMP", "GLO"], "9CA3AF")):
        wsC.conditional_formatting.add(rng, FormulaRule(formula=[cond(tipos)], fill=PatternFill("solid", start_color=color, end_color=color)))
    for c, w in zip("ABCDEFGHIJKLMN", (12, 58, 5, 7, 6, 9, 6, 8, 18, 6, 10, 10, 12, 7)):
        wsC.column_dimensions[c].width = w
    wsC.freeze_panes = wsC.cell(hdr + 1, 3)

    # ---------------- Metrados por actividad
    wsM["A1"] = "METRADOS POR ACTIVIDAD Y SECTOR (cada partida repartida según el modelo)"; wsM["A1"].font = f12b
    wsM["A2"] = "Peso = parte de la partida que ejecuta la actividad (del modelo de Revit). Metrado = metrado del presupuesto x peso."; wsM["A2"].font = f8
    mcols = ["Actividad", "Item", "Descripción", "Und.", "Peso", "Metrado", "Precio S/.", "Monto S/.", "Rend./día", "Cuadrilla-días"]
    for c, t in enumerate(mcols, 1):
        x = wsM.cell(4, c, t); x.font = f8b; x.border = caja; x.alignment = CEN
    for j, cod in enumerate(claves):
        x = wsM.cell(4, 11 + j, f"{recursos[cod][0]} ({recursos[cod][1]})"); x.font = f8b; x.alignment = CEN
        wsM.column_dimensions[CL(11 + j)].width = 12
    wsM.row_dimensions[4].height = 60
    BQ, KQ = Q("Presupuesto Base"), Q("Coeficientes APU")
    P_ = {p["item"]: p for p in part}
    asign_rows.sort(key=lambda t: (orden.index(t[0]), t[1]))
    for i, (aid, it, w) in enumerate(asign_rows):
        rr = met_ini + i; p = P_[it]
        wsM.cell(rr, 1, aid).font = f8; wsM.cell(rr, 2, it).font = f8; wsM.cell(rr, 3, p["desc"]).font = f8; wsM.cell(rr, 4, p["und"]).font = f8
        x = wsM.cell(rr, 5, round(w, 8)); x.number_format = "0.0000%"; x.font = fin8
        x = wsM.cell(rr, 6, f"=ROUND(INDEX({BQ}!$D:$D,MATCH(B{rr},{BQ}!$A:$A,0))*E{rr},4)"); x.number_format = N2; x.font = f8
        x = wsM.cell(rr, 7, f"=INDEX({BQ}!$E:$E,MATCH(B{rr},{BQ}!$A:$A,0))"); x.number_format = N2; x.font = f8
        x = wsM.cell(rr, 8, f"=ROUND(F{rr}*G{rr},2)"); x.number_format = N2; x.font = f8
        rend = (REND_ANCLAJE if "ANCLAJES" in p["desc"].upper() else apu.get(it, {}).get("rend") or 0) \
            if cuenta_para_duracion(it, p, acts[aid]["tipo"]) else 0
        x = wsM.cell(rr, 9, rend); x.number_format = N2; x.font = fin8
        x = wsM.cell(rr, 10, f"=IF(I{rr}>0,F{rr}/I{rr},0)"); x.number_format = N2; x.font = f8
        kr = filaK[it]
        for j, cod in enumerate(claves):
            if apu.get(it, {}).get("rec", {}).get(cod):
                x = wsM.cell(rr, 11 + j, f"=F{rr}*{KQ}!{CL(3 + j)}${kr}"); x.number_format = N2; x.font = f8
    for c, w in zip("ABCDEFGHIJ", (12, 9, 48, 6, 9, 11, 10, 12, 8, 10)):
        wsM.column_dimensions[c].width = w
    wsM.freeze_panes = "C5"
    for rr in range(4, 4 + len(part)):
        it = wsB.cell(rr, 1).value
        x = wsB.cell(rr, 7, f"=SUMIF({MQ}!$B${met_ini}:$B${met_fin},A{rr},{MQ}!$F${met_ini}:$F${met_fin})"); x.number_format = N2; x.font = f8
        x = wsB.cell(rr, 8, f"=G{rr}-D{rr}"); x.number_format = N2; x.font = f8

    # ---------------- Cronograma valorizado
    MS = meses(ini_obra, fin_obra)
    wsV["A1"] = "CRONOGRAMA VALORIZADO MENSUAL (costo directo programado por valorización)"; wsV["A1"].font = f12b
    wsV["A2"] = "Cada actividad se reparte en los meses en proporción a sus días laborables."; wsV["A2"].font = f8
    vh = 6
    for c, t in enumerate(["ID", "Actividad", "Monto CD S/.", "Inicio", "Fin", "Días lab."], 1):
        x = wsV.cell(vh, c, t); x.font = f8b; x.border = caja; x.alignment = CEN
    CQ = Q("Cronograma")
    for k, (mi, mf) in enumerate(MS):
        c = 7 + k
        x = wsV.cell(vh - 2, c, mi); x.number_format = FD; x.font = f8
        x = wsV.cell(vh - 1, c, mf); x.number_format = FD; x.font = f8
        x = wsV.cell(vh, c, f"Val. N° {k + 1:02d}"); x.font = f8b; x.border = caja; x.alignment = CEN
        wsV.column_dimensions[CL(c)].width = 12
    wsV.cell(vh - 2, 6, "desde").font = f8; wsV.cell(vh - 1, 6, "hasta").font = f8
    for aid in orden:
        rr, rc = fila[aid], fila[aid]
        wsV.cell(rr, 1, f"={CQ}!A{rc}").font = f8
        wsV.cell(rr, 2, f"={CQ}!B{rc}").font = f8
        x = wsV.cell(rr, 3, f"={CQ}!M{rc}"); x.number_format = N2; x.font = f8
        x = wsV.cell(rr, 4, f"={CQ}!K{rc}"); x.number_format = FD; x.font = f8
        x = wsV.cell(rr, 5, f"={CQ}!L{rc}"); x.number_format = FD; x.font = f8
        x = wsV.cell(rr, 6, f'=_xlfn.NETWORKDAYS.INTL(D{rr},E{rr},"{WEEKEND}",FER)'); x.font = f8
        for k in range(len(MS)):
            c = CL(7 + k)
            fx = (f'=IF(MAX($D{rr},{c}${vh - 2})>MIN($E{rr},{c}${vh - 1}),0,'
                  f'ROUND($C{rr}*_xlfn.NETWORKDAYS.INTL(MAX($D{rr},{c}${vh - 2}),MIN($E{rr},{c}${vh - 1}),"{WEEKEND}",FER)/$F{rr},2))')
            x = wsV.cell(rr, 7 + k, fx); x.number_format = N2; x.font = f8
    r1 = uC + 2
    filas_res = [("Costo directo del mes", lambda c: f"=SUM({c}{hdr + 1}:{c}{uC})"),
                 ("Costo directo acumulado", lambda c: f"={c}{r1}" if c == "G" else f"={CL_prev(c)}{r1 + 1}+{c}{r1}"),
                 ("% acumulado (curva S programada)", lambda c: f"=IF($C${r1}=0,0,{c}{r1 + 1}/$C${r1})"),
                 ("Valorización del mes con GG y utilidad", lambda c: f"=ROUND({c}{r1}*(1+'Parámetros'!$C$5+'Parámetros'!$D$5),2)"),
                 ("IGV", lambda c: f"=ROUND({c}{r1 + 3}*'Parámetros'!$E$5,2)"),
                 ("Total valorización del mes (con IGV)", lambda c: f"={c}{r1 + 3}+{c}{r1 + 4}")]
    for i, (lbl, fx) in enumerate(filas_res):
        rr = r1 + i
        wsV.cell(rr, 2, lbl).font = f9b
        for k in range(len(MS)):
            c = CL(7 + k)
            x = wsV.cell(rr, 7 + k, fx(c)); x.font = f9b; x.number_format = "0.0%" if "%" in lbl else N2
    x = wsV.cell(r1, 3, f"=SUM(C{hdr + 1}:C{uC})"); x.number_format = N2; x.font = f9b
    x = wsV.cell(r1 + 5, 3, f"=SUM(G{r1 + 5}:{CL(6 + len(MS))}{r1 + 5})"); x.number_format = N2; x.font = f9b
    ch = LineChart(); ch.title = "Curva S programada (costo directo acumulado)"; ch.y_axis.title = "S/."; ch.x_axis.title = "Valorización"
    ch.add_data(Reference(wsV, min_col=7, max_col=6 + len(MS), min_row=r1 + 1, max_row=r1 + 1), from_rows=True, titles_from_data=False)
    ch.set_categories(Reference(wsV, min_col=7, max_col=6 + len(MS), min_row=vh, max_row=vh))
    ch.height, ch.width = 8, 18; ch.legend = None; ch.roundedCorners = False
    ch.title.overlay = False; ch.x_axis.title.overlay = False; ch.y_axis.title.overlay = False
    ch.x_axis.delete = False; ch.y_axis.delete = False
    wsV.add_chart(ch, f"B{r1 + 8}")
    for c, w in zip("ABCDEF", (12, 52, 13, 10, 10, 7)):
        wsV.column_dimensions[c].width = w
    wsV.freeze_panes = wsV.cell(vh + 1, 3)

    # ---------------- Adquisiciones y recursos (por mes)
    wsR["A1"] = "CALENDARIO DE ADQUISICIONES Y RECURSOS (cantidad requerida en cada mes)"; wsR["A1"].font = f12b
    wsR["A2"] = "Pedir los materiales con 2 semanas de anticipación (acero y premezclado: confirmar plazos con el proveedor)."; wsR["A2"].font = f8
    for c, t in enumerate(["Código", "Recurso", "Und.", "Total"], 1):
        x = wsR.cell(4, c, t); x.font = f8b; x.border = caja; x.alignment = CEN
    for k in range(len(MS)):
        x = wsR.cell(4, 5 + k, f"Val. N° {k + 1:02d}"); x.font = f8b; x.border = caja; x.alignment = CEN
        x = wsR.cell(3, 5 + k, MS[k][0]); x.number_format = "mmm-yy"; x.font = f8
        wsR.column_dimensions[CL(5 + k)].width = 11
    VQ = Q("Cronograma Valorizado")
    # columna auxiliar en Cronograma: recursos por actividad (SUMIF de metrados por actividad)
    rc0 = g0 + nsem + 2
    wsC.cell(hdr, rc0 - 1, "Recursos por actividad →").font = f8b
    for j, cod in enumerate(claves):
        c = rc0 + j
        x = wsC.cell(hdr, c, f"{recursos[cod][0]} ({recursos[cod][1]})"); x.font = f8b; x.alignment = CEN
        wsC.column_dimensions[CL(c)].width = 11
        for aid in orden:
            rr = fila[aid]
            x = wsC.cell(rr, c, f"=SUMIF({MQ}!$A${met_ini}:$A${met_fin},$A{rr},{MQ}!{CL(11 + j)}${met_ini}:{CL(11 + j)}${met_fin})")
            x.number_format = N2; x.font = fgr8
    rr = 5
    grupos_rec = [("MANO DE OBRA (hh)", [c for c in claves if c.startswith("0147")]),
                  ("MATERIALES", [c for c in claves if c.startswith("02")]),
                  ("EQUIPOS (hm / m3)", [c for c in claves if c.startswith("03")])]
    filas_hh = []
    for titulo, cods in grupos_rec:
        wsR.cell(rr, 1, titulo).font = f9b; rr += 1
        for cod in cods:
            j = claves.index(cod); cc = CL(rc0 + j)
            wsR.cell(rr, 1, cod).font = f8; wsR.cell(rr, 2, recursos[cod][0]).font = f8; wsR.cell(rr, 3, recursos[cod][1]).font = f8
            for k in range(len(MS)):
                vc = CL(7 + k)
                fx = (f"=SUMPRODUCT({VQ}!{vc}${hdr + 1}:{vc}${uC}/{VQ}!$C${hdr + 1}:$C${uC},{CQ}!{cc}${hdr + 1}:{cc}${uC})")
                x = wsR.cell(rr, 5 + k, fx); x.number_format = N0; x.font = f8
            x = wsR.cell(rr, 4, f"=SUM(E{rr}:{CL(4 + len(MS))}{rr})"); x.number_format = N0; x.font = f8b
            if cod.startswith("0147"):
                filas_hh.append(rr)
            rr += 1
        rr += 1
    wsR.cell(rr, 2, "Días laborables del mes").font = f9b
    for k in range(len(MS)):
        x = wsR.cell(rr, 5 + k, f'=_xlfn.NETWORKDAYS.INTL({VQ}!{CL(7 + k)}${vh - 2},{VQ}!{CL(7 + k)}${vh - 1},"{WEEKEND}",FER)'); x.font = f8
    rdl = rr; rr += 1
    wsR.cell(rr, 2, "Personal promedio en obra (obreros/día) — histograma").font = f9b
    for k in range(len(MS)):
        c = CL(5 + k)
        x = wsR.cell(rr, 5 + k, f"=IF({c}{rdl}=0,0,ROUND(({'+'.join(f'{c}{h}' for h in filas_hh)})/8/{c}{rdl},0))"); x.font = f9b; x.number_format = N0
    rpe = rr
    for c, w in zip("ABCD", (12, 50, 6, 12)):
        wsR.column_dimensions[c].width = w
    ch2 = LineChart(); ch2.title = "Histograma de personal (promedio diario por mes)"; ch2.legend = None; ch2.roundedCorners = False
    ch2.add_data(Reference(wsR, min_col=5, max_col=4 + len(MS), min_row=rpe, max_row=rpe), from_rows=True, titles_from_data=False)
    ch2.set_categories(Reference(wsR, min_col=5, max_col=4 + len(MS), min_row=4, max_row=4))
    ch2.title.overlay = False; ch2.x_axis.delete = False; ch2.y_axis.delete = False; ch2.height, ch2.width = 7, 18
    wsR.add_chart(ch2, f"B{rpe + 3}")

    # ---------------- Sectorización
    wsS["A1"] = "SECTORIZACIÓN Y PAQUETES DE VACIADO"; wsS["A1"].font = f12b
    wsS["A2"] = ("Cada paquete = elementos verticales de un nivel + techo (vigas, losas, escaleras) del nivel superior. Se divide en 2 sectores (A y B) "
                 "por el eje más largo, con una junta que equilibra el concreto. Los muros anclados se ejecutan por anillos (paños alternados).")
    wsS["A2"].font = f8
    for c, t in enumerate(["Paquete", "Nombre", "Verticales del nivel", "Techo del nivel", "Eje de corte", "Junta (m)", "Sector",
                           "N° elementos", "Concreto m3", "Encofrado m2", "% concreto"], 1):
        x = wsS.cell(4, c, t); x.font = f8b; x.border = caja; x.alignment = CEN
    rr = 5
    for pid in sorted(paquetes, key=lambda k: (k != "CIM", k)):
        pq = paquetes[pid]
        for s in pq.get("sectores", []):
            el = [f for f in filas if f["paq"] == pid and f["sec"] == s]
            tot = sum(f["conc"] for f in filas if f["paq"] == pid)
            vals = (pid, pq["nombre"], nombre_nivel(pq["vertical"]) if pq["vertical"] else "-", nombre_nivel(pq["horizontal"]) if pq["horizontal"] else "-",
                    pq["eje"], pq["corte"], s, len(el), round(sum(f["conc"] for f in el), 2), round(sum(f["enc"] for f in el), 2),
                    sum(f["conc"] for f in el) / tot if tot else 0)
            for c, v in enumerate(vals, 1):
                x = wsS.cell(rr, c, v); x.font = f8
                if c in (9, 10): x.number_format = N2
                if c == 11: x.number_format = "0.0%"
            rr += 1
    for anillo in ("AN1", "AN2"):
        el = [f for f in filas if f["paq"] == anillo]
        vals = (anillo, "Anillo de muros anclados " + ("1 (semisótano)" if anillo == "AN1" else "2 (sótano 1)"), "-", "-", "-", "-", "paños alternados",
                len(el), round(sum(f["conc"] for f in el), 2), round(sum(f["enc"] for f in el), 2), 1)
        for c, v in enumerate(vals, 1):
            x = wsS.cell(rr, c, v); x.font = f8
            if c in (9, 10): x.number_format = N2
            if c == 11: x.number_format = "0.0%"
        rr += 1
    rr += 1
    wsS.cell(rr, 1, "Planos de sectorización (SVG) en 5_Documentacion\\Sectorizacion\\. Criterios: la junta de construcción va en el tercio central "
                    "del paño de losa o viga, nunca sobre columna/placa; confirmar con el ingeniero estructural.").font = f8
    for c, w in zip("ABCDEFGHIJK", (8, 26, 14, 14, 8, 9, 15, 10, 11, 11, 9)):
        wsS.column_dimensions[c].width = w

    # ---------------- Formato de valorización (plantilla para la app)
    wsF["A1"] = "VALORIZACIÓN DE OBRA N°"; wsF["A1"].font = f12b
    x = wsF["D1"]; x.value = 1; x.font = fin9
    wsF["A2"] = "Periodo"; x = wsF["B2"]; x.value = MS[0][0]; x.number_format = FD; x.font = fin9
    x = wsF["C2"]; x.value = MS[0][1]; x.number_format = FD; x.font = fin9
    wsF["A3"] = ("Cómo llenar: escribe solo en las columnas azules «Metrado anterior» y «Metrado del mes». Lo demás se calcula. "
                 "La app NEXO OBRA (NO) llenará estas columnas desde los sectores y elementos ejecutados.")
    wsF["A3"].font = f8
    fc_ = ["Item", "Descripción", "Und.", "Metrado contrato", "Precio S/.", "Presupuesto S/.", "Metrado anterior", "Metrado del mes",
           "Metrado acumulado", "Saldo metrado", "% avance", "Monto anterior", "Monto del mes", "Monto acumulado", "Saldo S/."]
    for c, t in enumerate(fc_, 1):
        x = wsF.cell(5, c, t); x.font = f8b; x.border = caja; x.alignment = CEN
    for i, p in enumerate(part):
        rr = 6 + i; br = filaB[p["item"]]
        wsF.cell(rr, 1, p["item"]).font = f8; wsF.cell(rr, 2, p["desc"]).font = f8; wsF.cell(rr, 3, p["und"]).font = f8
        for c, fx in ((4, f"={BQ}!D{br}"), (5, f"={BQ}!E{br}"), (6, f"=ROUND(D{rr}*E{rr},2)")):
            x = wsF.cell(rr, c, fx); x.number_format = N2; x.font = f8
        for c in (7, 8):
            x = wsF.cell(rr, c); x.font = fin8; x.number_format = N2
        for c, fx, fm in ((9, f"=G{rr}+H{rr}", N2), (10, f"=D{rr}-I{rr}", N2), (11, f"=IF(D{rr}=0,0,I{rr}/D{rr})", "0.0%"),
                          (12, f"=ROUND(G{rr}*E{rr},2)", N2), (13, f"=ROUND(H{rr}*E{rr},2)", N2), (14, f"=L{rr}+M{rr}", N2),
                          (15, f"=F{rr}-N{rr}", N2)):
            x = wsF.cell(rr, c, fx); x.number_format = fm; x.font = f8
    uf = 5 + len(part)
    rr = uf + 2
    tot_rows = {}
    for lbl, key, fx in (("Costo directo", "cd", lambda c: f"=SUM({c}6:{c}{uf})"),
                         ("Gastos generales", "gg", lambda c: f"=ROUND({c}{tot_rows['cd']}*'Parámetros'!$C$5,2)"),
                         ("Utilidad", "ut", lambda c: f"=ROUND({c}{tot_rows['cd']}*'Parámetros'!$D$5,2)"),
                         ("Subtotal", "st", lambda c: f"={c}{tot_rows['cd']}+{c}{tot_rows['gg']}+{c}{tot_rows['ut']}"),
                         ("IGV", "igv", lambda c: f"=ROUND({c}{tot_rows['st']}*'Parámetros'!$E$5,2)"),
                         ("Valorización bruta", "vb", lambda c: f"={c}{tot_rows['st']}+{c}{tot_rows['igv']}")):
        tot_rows[key] = rr
        wsF.cell(rr, 2, lbl).font = f9b
        for c in ("F", "L", "M", "N", "O"):
            x = wsF[f"{c}{rr}"]; x.value = fx(c); x.number_format = N2; x.font = f9b
        rr += 1
    rr += 1
    wsF.cell(rr, 2, "Amortización de adelanto directo (% de la valorización del mes)").font = f9
    x = wsF.cell(rr, 7, 0); x.number_format = "0.00%"; x.font = fin9
    x = wsF.cell(rr, 13, f"=-ROUND(M{tot_rows['vb']}*G{rr},2)"); x.number_format = N2; x.font = f9
    ra = rr; rr += 1
    wsF.cell(rr, 2, "Amortización de adelanto de materiales (S/.)").font = f9
    x = wsF.cell(rr, 13, 0); x.number_format = N2; x.font = fin9
    rb = rr; rr += 1
    wsF.cell(rr, 2, "Retención por fondo de garantía (% de la valorización del mes)").font = f9
    x = wsF.cell(rr, 7, 0); x.number_format = "0.00%"; x.font = fin9
    x = wsF.cell(rr, 13, f"=-ROUND(M{tot_rows['vb']}*G{rr},2)"); x.number_format = N2; x.font = f9
    rc = rr; rr += 1
    wsF.cell(rr, 2, "NETO A PAGAR EN ESTA VALORIZACIÓN").font = f9b
    x = wsF.cell(rr, 13, f"=M{tot_rows['vb']}+M{ra}-ABS(M{rb})+M{rc}"); x.number_format = N2; x.font = f9b
    for c, w in zip("ABCDEFGHIJKLMNO", (9, 52, 6, 11, 10, 13, 11, 11, 11, 11, 8, 12, 12, 13, 13)):
        wsF.column_dimensions[c].width = w
    wsF.freeze_panes = "D6"

    wb.calculation.fullCalcOnLoad = True
    wb.save(ruta)
    return {"fila": fila, "hdr": hdr, "uC": uC, "meses": MS, "vh": vh, "r1": r1}


def CL_prev(c):
    from openpyxl.utils import column_index_from_string as ci
    return CL(ci(c) - 1)


# ------------------------------------------------------------------ JSON para la app
def exportar_json(ruta, D, filas, paquetes, acts, orden, part, asign, MS):
    P = {p["item"]: p for p in part}
    el = []
    for f in filas:
        if f["id"] < 0:
            continue
        ac = {}
        for it, pes in asign.items():
            pass
        el.append({"elementId": f["id"], "grupo": f["g"], "tipo": f["tipo"], "nivel": f["niv"], "paquete": f["paq"], "sector": f["sec"],
                   "concreto_m3": round(f["conc"], 4), "encofrado_m2": round(f["enc"], 4)})
    pv = []
    acum = 0
    for k, (mi, mf) in enumerate(MS):
        s = 0
        for aid in orden:
            a = acts[aid]
            lo, hi = max(a["ini"], mi), min(a["fin"], mf)
            if lo <= hi:
                s += round(a["monto"] * networkdays(lo, hi) / networkdays(a["ini"], a["fin"]), 2)
        acum += s
        pv.append({"n": k + 1, "desde": mi.isoformat(), "hasta": mf.isoformat(), "pv_mes": round(s, 2), "pv_acum": round(acum, 2)})
    datos = {
        "proyecto": {"modelo": D["modelo"], "lugar": "Arequipa", "moneda": "PEN", "calendario": {"laborables": "lunes a sábado",
                     "feriados": [d for d, _ in FERIADOS]}, "inicio": INICIO_OBRA.isoformat()},
        "paquetes": [{k: (list(v) if isinstance(v, tuple) else v) for k, v in pq.items()} for pq in paquetes.values()],
        "actividades": [{"id": a, "nombre": acts[a]["nombre"], "tipo": acts[a]["tipo"], "paquete": acts[a]["paq"], "sector": acts[a]["sec"],
                         "predecesoras": acts[a]["preds"], "inicio": acts[a]["ini"].isoformat(), "fin": acts[a]["fin"].isoformat(),
                         "duracion": acts[a]["dur"], "monto_cd": round(acts[a]["monto"], 2),
                         "partidas": [{"item": it, "metrado": round(P[it]["met"] * w, 4)} for it, pes in asign.items() for x, w in pes.items() if x == a]}
                        for a in orden],
        "partidas": [{"item": p["item"], "descripcion": p["desc"], "und": p["und"], "metrado": p["met"], "pu": p["pu"]} for p in part],
        "periodos": pv,
        "elementos": el,
        "reglas_de_medicion": {
            "concreto": "100 % del elemento al vaciarlo", "encofrado": "100 % al encofrar (incluye desencofrado)",
            "acero": "100 % al colocarlo y ser aprobado por la supervisión", "excavacion": "m3 ejecutados por capa (topografía)",
            "anclajes": "80 % al perforar e inyectar, 20 % al tensar", "globales": "obras provisionales al instalarse; SST por mes"},
    }
    Path(ruta).write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    return pv


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    mj, pres, out = sys.argv[1], sys.argv[2], Path(sys.argv[3])
    D, filas, X = M.calcular(mj)
    part, cap, pie, apu = leer_presupuesto(pres)
    CAPITULOS, PIE = cap, {"gg": pie.get("GASTOS GENERALES"), "ut": pie.get("UTILIDAD"), "igv": pie.get("IGV")}
    paquetes, nombres = sectorizar(D, filas)
    acts, orden, asign = construir_actividades(D, filas, paquetes, part, X)
    orden = programar(acts, orden, part, asign, apu, {k: v[1] for k, v in TIPOS.items()})
    (out / "Sectorizacion").mkdir(parents=True, exist_ok=True)
    for pid, pq in paquetes.items():
        if pq.get("sectores"):
            svg_sector(pq, filas, out / "Sectorizacion" / f"sectorizacion_{pid}.svg")
    info = excel(out / f"Cronograma_Valorizado_{D['modelo']}.xlsx", D, filas, paquetes, acts, orden, part, cap, pie, apu, asign, X)
    pv = exportar_json(out / "nexo_no_datos.json", D, filas, paquetes, acts, orden, part, asign, info["meses"])
    pr = parametros_revit(Path(mj).parent / f"parametros_revit_{D['modelo']}.json", D, filas, acts, part, INICIO_OBRA)
    print(f"Parámetros VZ_ para {len(pr)} elementos | monto total S/ {sum(e['monto_concreto'] + e['monto_encofrado'] + e['monto_acero'] for e in pr):,.2f}")
    json.dump({a: [acts[a]["ini"].isoformat(), acts[a]["fin"].isoformat(), acts[a]["dur"], round(acts[a]["monto"], 2)] for a in orden},
              open(out / "_programa_python.json", "w", encoding="utf-8"))
    fin = max(a["fin"] for a in acts.values())
    print(f"Actividades: {len(orden)} | inicio {INICIO_OBRA} | fin {fin} | {len(info['meses'])} valorizaciones")
    print("CD programado:", round(sum(a["monto"] for a in acts.values()), 2), "| CD presupuesto:", round(sum(round(p['met'] * p['pu'], 2) for p in part), 2))
    for p in pv:
        print(f"  Val {p['n']:02d} {p['desde']}  PV mes {p['pv_mes']:>13,.2f}  acum {p['pv_acum']:>14,.2f}")
