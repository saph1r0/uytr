"""NEXO OBRA (NO) · Metrados de estructuras a partir de metrados_<modelo>.json (escrito por NexoMetrados en Revit).

Clasifica cada elemento en una partida (zapatas, columnas, placas, vigas, losas, ...) y calcula:
concreto (m3), encofrado (m2), base para acero (m3 x cuantía) y otros metrados (solado, excavación, ladrillo).
Los criterios y supuestos están en SUPUESTOS y se copian al Excel.
"""
import json
import re
import sys
from collections import defaultdict

ESP_ALIGERADO = 0.0875      # m3 de concreto por m2 de aligerado h=0.20 (ladrillo 15x30x30)
LADRILLO_M2 = 8.33          # ladrillos de techo por m2
E_LOSA = 0.20               # espesor de losa para descontar encofrado lateral de vigas
FC_DEFECTO = 280            # f'c cuando ni el tipo ni el material lo indican
COTA_TERRENO = 0.00         # nivel de terreno natural asumido (m)
ESPONJAMIENTO = 1.25

SUPUESTOS = [
    ("Concreto en losa aligerada h=0.20", f"{ESP_ALIGERADO} m3/m2 (ladrillo de techo 15x30x30), no el volumen sólido del modelo"),
    ("Ladrillo de techo", f"{LADRILLO_M2} und/m2"),
    ("Encofrado de vigas", f"Fondo + caras laterales x (h - {E_LOSA}) / h; las vigas de h=0.20 (chatas/viguetas) solo fondo"),
    ("Encofrado de columnas, placas y muros", "Caras laterales del sólido de Revit (incluye cabezas); no descuenta contactos entre elementos"),
    ("Encofrado de losas", "Cara inferior + bordes del sólido"),
    ("f'c sin dato", f"Si el tipo o material no indica f'c se asume {FC_DEFECTO} kg/cm2"),
    ("Terreno natural", f"Cota {COTA_TERRENO:+.2f} m; excavación masiva = huella del sótano x (terreno - fondo de cimentación)"),
    ("Esponjamiento", f"{ESPONJAMIENTO} para eliminación de material"),
    ("Acero de refuerzo", "Cuantías kg/m3 por tipo de elemento (hoja Supuestos, editables); el modelo solo tiene 197 barras"),
    ("Excavación masiva por capas", "Cada sótano se excava hasta su nivel - 0.10 m (falso piso), con la huella de los elementos de ese sótano"),
    ("Excavación manual", "Cimentaciones (zapatas, cimientos, losas de cimentación y cisterna, subzapatas) por debajo del fondo de la excavación masiva, + 0.10 m de solado"),
    ("Encofrado de muros anclados", "Una sola cara (el muro se vacía contra el terreno)"),
    ("Anclajes postensados", f"Uno cada {2.5} m de muro anclado por nivel (largo de muros del modelo); precio de subcontrato a cotizar"),
    ("Aligerado con bovedillas genéricas", f"Las bovedillas del modelo se convierten a m2 de aligerado ({LADRILLO_M2} und/m2)"),
    ("Excluido", "Calzaduras, juntas sísmicas, tanque elevado y caja de ascensor no modelados, instalaciones, acabados"),
]
SEP_ANCLAJE = 2.5


def capas_excavacion(D, filas):
    """Capas de excavación masiva: un sótano por capa (huella de sus elementos y fondo = nivel - 0.10)."""
    niv = {l["nombre"]: l["cota"] for l in D["niveles"]}
    capas = []
    for n in sorted((n for n, c in niv.items() if c < COTA_TERRENO), key=lambda n: -niv[n]):
        s = [f for f in filas if f["nivel"] == n and f["caja"]]
        if not s:
            continue
        x0, y0 = min(f["caja"][0] for f in s), min(f["caja"][1] for f in s)
        x1, y1 = max(f["caja"][3] for f in s), max(f["caja"][4] for f in s)
        capas.append({"nivel": n, "caja": (x0, y0, x1, y1), "area": (x1 - x0) * (y1 - y0), "fondo": round(niv[n] - 0.10, 3)})
    return capas


def fc_de(e):
    textos = [e.get("tipo") or ""] + [m["material"] or "" for m in e["materiales"]]
    for t in textos:
        m = re.search(r"f'?c\s*=?\s*(\d{3})", t, re.I)
        if m:
            return int(m.group(1))
    return FC_DEFECTO


def bh_de(tipo):
    m = re.search(r"\(\s*(\d+(?:\.\d+)?)\s*x\s*(\d+(?:\.\d+)?)", tipo or "")
    return (float(m.group(1)), float(m.group(2))) if m else (None, None)


def grupo_de(e):
    """Devuelve (código de grupo, nombre) o None si no se presupuesta."""
    b, fam, tipo = e["bic"], (e["familia"] or "").upper(), (e["tipo"] or "").upper()
    mats = " ".join(m["material"] or "" for m in e["materiales"]).upper()
    if b == "OST_StructuralFoundation":
        if "SUB-ZAPATA" in tipo or "CONCRETO SIMPLE" in mats:
            return "SUBZ"
        if "MURO" in tipo or "CIMENTACIÓN DE MURO" in fam:
            return "CIMM"
        if "LOSA" in fam or "LOSA" in tipo:
            return "LOSC"
        return "ZAP"
    if b in ("OST_StructuralColumns", "OST_Columns"):
        return "PLA" if "PLACA" in fam or "PLACA" in tipo else "COL"
    if b == "OST_StructuralFraming":
        _, h = bh_de(e["tipo"])
        return "VCH" if h is not None and h <= E_LOSA + 1e-6 else "VIG"
    if b == "OST_Floors":
        return "ALI" if "ALIGERAD" in tipo else "LMA"
    if b == "OST_Walls":
        if "ALBAÑILERIA" in tipo or "ALBAÑILERIA" in mats:
            return "ALB"
        return "PLA" if "PLACA" in tipo else "MUR"
    if b == "OST_Stairs":
        return "ESC" if (e["volumen_solido_m3"] or e["volumen_m3"]) > 0 else None
    if b == "OST_Ramps":
        return "RAM"
    return None   # modelos genéricos (bovedillas, viguetas sueltas) no se presupuestan


def encofrado(e, g):
    c = e["caras_m2"]
    if g == "VIG":
        _, h = bh_de(e["tipo"])
        f = max(0.0, (h - E_LOSA) / h) if h else 1.0
        return c["inferior"] + c["laterales"] * f
    if g == "VCH":
        return c["inferior"]
    if g in ("LMA", "ALI", "ESC", "RAM"):
        return c["inferior"] + c["laterales"]
    if g == "MUR" and "ANCLAJE" in (e["tipo"] or "").upper():
        return e["area_m2"] or c["laterales"] / 2      # muro anclado: se vacía contra el terreno, encofrado por una cara
    if g in ("COL", "PLA", "MUR", "ZAP", "CIMM", "LOSC"):
        return c["laterales"]
    return 0.0


def calcular(ruta_json):
    D = json.load(open(ruta_json, encoding="utf-8"))
    filas = []
    for e in D["elementos"]:
        g = grupo_de(e)
        if not g:
            continue
        vol = e["volumen_m3"] or e["volumen_solido_m3"]
        area = e["area_m2"] or e["caras_m2"]["inferior"]
        conc = ESP_ALIGERADO * area if g == "ALI" else vol
        filas.append({
            "id": e["id"], "g": g, "fc": fc_de(e) if g not in ("SUBZ", "ALB") else None,
            "tipo": e["tipo"], "familia": e["familia"], "nivel": e["nivel"] or "(sin nivel)",
            "vol": vol, "conc": conc, "area": area, "enc": encofrado(e, g),
            "inf": e["caras_m2"]["inferior"], "caja": e["caja_m"], "largo": e["largo_m"],
        })
    # Huella del sótano y fondo de cimentación (para excavación masiva)
    cim = [f for f in filas if f["g"] in ("ZAP", "CIMM", "LOSC", "SUBZ")]
    zmin = min(f["caja"][2] for f in cim)
    sot = [f for f in filas if f["caja"] and f["caja"][2] < -1.0]
    x0, y0 = min(f["caja"][0] for f in sot), min(f["caja"][1] for f in sot)
    x1, y1 = max(f["caja"][3] for f in sot), max(f["caja"][4] for f in sot)
    huella = (x1 - x0) * (y1 - y0)
    zap_fondo = sum(f["inf"] for f in cim)
    zap_prof = sum(f["inf"] * (f["caja"][5] - f["caja"][2] + 0.10) for f in cim if f["g"] in ("ZAP", "CIMM", "SUBZ"))

    # bovedillas genéricas -> m2 de aligerado
    nb = sum(1 for e in D["elementos"] if e["bic"] == "OST_GenericModel" and "BOVEDILLA" in (e["tipo"] or "").upper())
    if nb:
        a = nb / LADRILLO_M2
        filas.append({"id": -1, "g": "ALI", "fc": FC_DEFECTO, "tipo": f"Aligerado con {nb} bovedillas genéricas", "familia": "",
                      "nivel": "(modelos genéricos)", "vol": 0, "conc": ESP_ALIGERADO * a, "area": a, "enc": a, "inf": a,
                      "caja": None, "largo": 0})

    capas = capas_excavacion(D, filas)

    def fondo_de(f):
        cx, cy = (f["caja"][0] + f["caja"][3]) / 2, (f["caja"][1] + f["caja"][4]) / 2
        z = COTA_TERRENO
        for c in capas:
            x0, y0, x1, y1 = c["caja"]
            if x0 <= cx <= x1 and y0 <= cy <= y1:
                z = min(z, c["fondo"])
        return z
    lin_exz = defaultdict(lambda: [0, 0.0])
    for f in cim:
        prof = fondo_de(f) - (f["caja"][2] - 0.10)
        if prof > 0:
            k = f"{f['tipo']} — {f['nivel']} (fondo masiva {fondo_de(f):.2f})"
            lin_exz[k][0] += 1; lin_exz[k][1] += f["inf"] * prof
    # falso piso: piso de cada capa no cubierto por la capa de abajo, menos losas de cimentación/cisterna
    lin_fp = []
    for i, c in enumerate(capas):
        deb = capas[i + 1]["area"] if i + 1 < len(capas) else 0.0
        if c["area"] - deb > 0.5:
            lin_fp.append((f"Piso de {c['nivel']} (huella {c['area']:.1f} - sótano inferior {deb:.1f})", 1, c["area"] - deb))
    losc = sum(f["inf"] for f in filas if f["g"] == "LOSC")
    lin_fp.append(("Menos losas de cimentación y cisterna (ya son piso)", 1, -losc))
    # anclajes
    E = {e["id"]: e for e in D["elementos"]}
    lar = defaultdict(float)
    for f in filas:
        if f["g"] == "MUR" and "ANCLAJE" in (f["tipo"] or "").upper():
            lar[f["nivel"]] += E[f["id"]]["largo_m"]
    lin_anc = [(f"Muro anclado {n}: {l:.2f} m / {SEP_ANCLAJE} m", 1, -(-l // SEP_ANCLAJE)) for n, l in sorted(lar.items())]
    extra = {
        "capas": capas, "lin_exz": sorted((k, v[0], v[1]) for k, v in lin_exz.items()), "lin_fp": lin_fp, "lin_anc": lin_anc,
        "huella_m2": huella, "dim_huella": (x1 - x0, y1 - y0), "fondo_cim": zmin,
        "excav_masiva": huella * (COTA_TERRENO - zmin - 0.0) if zmin < COTA_TERRENO else 0.0,
        "excav_local": zap_prof, "solado_m2": zap_fondo,
        "ladrillo_und": LADRILLO_M2 * sum(f["area"] for f in filas if f["g"] == "ALI"),
        "alb_m2": sum(f["area"] for f in filas if f["g"] == "ALB"),
        "acero_modelado_m": sum(b["largo_total_m"] for b in D["acero"]),
    }
    return D, filas, extra


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    D, filas, x = calcular(sys.argv[1])
    t = defaultdict(lambda: [0, 0.0, 0.0, 0.0])
    for f in filas:
        r = t[(f["g"], f["fc"])]; r[0] += 1; r[1] += f["conc"]; r[2] += f["enc"]; r[3] += f["area"]
    for k in sorted(t, key=str):
        r = t[k]; print(f"{k[0]:5} f'c={k[1]}  n={r[0]:4}  concreto={r[1]:9.2f} m3  encofrado={r[2]:9.2f} m2  area={r[3]:9.2f}")
    print({k: (round(v, 2) if isinstance(v, float) else v) for k, v in x.items()})
    print("Concreto total:", round(sum(f["conc"] for f in filas), 2))
