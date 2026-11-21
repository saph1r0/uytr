"""Arma 1_Panel/index.html: toma panel_fuente.html e incrusta el cronograma programado y los coeficientes de materiales del APU.

Fuentes (no se inventa nada):
  datos/nexo_no_datos.json                         actividades, periodos, feriados (4_Python/cronograma_nexo.py)
  ../3_Modelo_y_parametros/parametros_revit_*.json actividades de concreto/acero/encofrado de cada elemento
  ../5_Documentacion/Presupuesto_*_v2.xlsx          coeficientes de materiales de cada análisis de precios unitarios
Uso: python armar_panel.py   (volver a correrlo cuando cambie el cronograma o el presupuesto)
"""
import json
import sys
from pathlib import Path

AQUI = Path(__file__).parent
sys.path.insert(0, str(AQUI.parent / "4_Python"))
from cronograma_nexo import leer_presupuesto  # noqa: E402

crono = json.loads((AQUI / "datos" / "nexo_no_datos.json").read_text(encoding="utf-8"))
modelo = crono["proyecto"]["modelo"]
pr = json.loads((AQUI.parent / "3_Modelo_y_parametros" / f"parametros_revit_{modelo}.json").read_text(encoding="utf-8"))
pres = AQUI.parent / "5_Documentacion" / f"Presupuesto_Estructuras_{modelo}_v2.xlsx"
_, _, _, apu = leer_presupuesto(pres)
recursos, coef = {}, {}
for it, a in apu.items():
    for cod, (desc, und, q) in a["rec"].items():
        if cod.startswith("02") and q:              # solo materiales
            recursos[cod] = [desc, und]
            coef.setdefault(it, {})[cod] = q
datos = {
    "modelo": modelo,
    "inicio": crono["proyecto"]["inicio"],
    "feriados": crono["proyecto"]["calendario"]["feriados"],
    "paquetes": {p["id"]: p["nombre"] for p in crono["paquetes"]},
    "actividades": [{k: a[k] for k in ("id", "nombre", "tipo", "paquete", "sector", "predecesoras", "inicio", "fin", "duracion", "monto_cd", "partidas")}
                    for a in crono["actividades"]],
    "periodos": crono["periodos"],
    "elementos": {e["id"]: [e["actividad"], e["act_acero"], e["act_encofrado"]] for e in pr["elementos"]},
    "recursos": recursos, "coef": coef, "fuentePresupuesto": pres.name,
}
fuente = (AQUI / "panel_fuente.html").read_text(encoding="utf-8")
marca = "/*__CRONOGRAMA__*/null"
assert marca in fuente, "falta la marca del cronograma en panel_fuente.html"
html = fuente.replace(marca, json.dumps(datos, ensure_ascii=False, separators=(",", ":")))
(AQUI / "index.html").write_text(html, encoding="utf-8")
print(f"index.html listo: {len(datos['actividades'])} actividades, {len(datos['elementos'])} elementos, "
      f"{len(recursos)} materiales, {len(html) // 1024} KB")
