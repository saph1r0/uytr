"""Arma la instantánea de solo lectura del panel para publicarla como página con comentarios.

Lee el panel y los datos desde el complemento de Revit (http://127.0.0.1:8770, Revit abierto con el modelo),
incrusta datos y geometría, y quita las etiquetas <!doctype>/<html>/<head>/<body> (la página publicada las pone sola).
Uso: python armar_instantanea.py [salida.html]
"""
import json
import re
import sys
import urllib.request
from datetime import date
from pathlib import Path

BASE = "http://127.0.0.1:8770"
leer = lambda r: urllib.request.urlopen(BASE + r, timeout=600).read().decode("utf-8")
salida = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).parent.parent / "5_Documentacion" / "Compartir" / "NEXO_OBRA_solo_lectura.html"
html, datos, geo = leer("/"), json.loads(leer("/api/datos")), json.loads(leer("/api/geometria"))
snap = json.dumps({"fecha": date.today().isoformat(), "datos": datos, "geo": geo}, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
html = re.sub(r"(?is)<!doctype[^>]*>|</?html[^>]*>|</?head>|</?body>", "", html)
html = re.sub(r'(?i)<meta charset="utf-8">\s*<meta name="viewport"[^>]*>', "", html, count=1)
html, n = re.subn(r'<script>\s*"use strict";', lambda m: f"<script>window.NEXO_SNAPSHOT={snap};</script>\n" + m.group(0), html, count=1)
assert n == 1, "no se encontró el script principal"
salida.parent.mkdir(parents=True, exist_ok=True)
salida.write_text(html.strip(), encoding="utf-8", newline="\n")
print(f"Instantánea: {salida} · {len(html) / 1048576:.1f} MB · {len(datos['elementos'])} elementos · {len(geo['elementos'])} con geometría")
