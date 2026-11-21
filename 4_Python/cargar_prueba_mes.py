"""PRUEBA: registra en Revit (vía el panel, http://127.0.0.1:8770) el metrado PROGRAMADO de un mes, como si se hubiera ejecutado.

Sirve para comprobar la valorización y el expediente. Cada actividad aporta lo programado en ese mes
(días laborables del mes / duración); se registra en su último día laborable del mes.
Uso: python cargar_prueba_mes.py 2027-01 [--simular]
Para deshacer: en Revit, Ctrl+Z una vez por cada fecha registrada (antes de guardar el modelo).
"""
import json
import sys
import urllib.request
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path

RAIZ = Path(__file__).parent.parent
mes, simular = sys.argv[1], "--simular" in sys.argv
crono = json.loads((RAIZ / "1_Panel" / "datos" / "nexo_no_datos.json").read_text(encoding="utf-8"))
pr = json.loads((RAIZ / "3_Modelo_y_parametros" / f"parametros_revit_{crono['proyecto']['modelo']}.json").read_text(encoding="utf-8"))
FER = {date.fromisoformat(f) for f in crono["proyecto"]["calendario"]["feriados"]}
lab = lambda d: d.weekday() != 6 and d not in FER
y, m = map(int, mes.split("-"))
m0 = date(y, m, 1); m1 = date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1)


def dias(a, b):
    n, d = 0, a
    while d <= b:
        n += lab(d); d += timedelta(days=1)
    return n


con_elementos = {p["item"] for e in pr["elementos"] for p in e["partidas"]}
registros = defaultdict(list)          # fecha -> [{id, item, cantidad}]
pv = 0.0
for a in crono["actividades"]:
    ini, fin = date.fromisoformat(a["inicio"]), date.fromisoformat(a["fin"])
    lo, hi = max(ini, m0), min(fin, m1)
    if lo > hi:
        continue
    f = dias(lo, hi) / max(1, a["duracion"])
    pv += a["monto_cd"] * f
    fecha = hi
    while not lab(fecha):
        fecha -= timedelta(days=1)
    k = fecha.isoformat()
    for p in a["partidas"]:
        if p["item"] not in con_elementos:                   # partidas sin elementos: total de la partida
            registros[k].append({"id": 0, "item": p["item"], "cantidad": round(p["metrado"] * f, 6)})
    for e in pr["elementos"]:                                # partidas con elementos: cada elemento de la actividad
        for p in e["partidas"]:
            act = e["actividad"] if p["clase"] in ("concreto", "muro") else e["act_encofrado"] if p["clase"] == "encofrado" else e["act_acero"]
            if act == a["id"]:
                registros[k].append({"id": e["id"], "item": p["item"], "cantidad": round(p["metrado"] * f, 6)})

print(f"Mes {mes}: programado (CD) S/ {pv:,.2f} · {sum(len(v) for v in registros.values())} registros en {len(registros)} fechas · simular={simular}")
tot_ok = 0
for fecha in sorted(registros):
    cuerpo = json.dumps({"fecha": fecha, "simular": simular, "registros": registros[fecha]}).encode()
    req = urllib.request.Request("http://127.0.0.1:8770/api/metrado", data=cuerpo, method="POST", headers={"Content-Type": "application/json"})
    r = json.loads(urllib.request.urlopen(req, timeout=600).read())
    tot_ok += len(r["aplicados"])
    print(f"  {fecha}: {len(r['aplicados'])} aplicados, {len(r['rechazados'])} rechazados, {len(r['alertas'])} alertas"
          + (f" · {r['rechazados'][:2]}" if r["rechazados"] else ""))
print("Total aplicados:", tot_ok)
