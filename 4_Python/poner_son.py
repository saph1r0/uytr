"""Escribe la línea «SON : ... SOLES» del presupuesto con el total ya recalculado (como el reporte de S10)."""
import sys
from openpyxl import load_workbook
from openpyxl.comments import Comment

U = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE",
     "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUNO", "VEINTIDOS",
     "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE"]
D = ["", "", "", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"]
C = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS",
     "OCHOCIENTOS", "NOVECIENTOS"]


def cien(n):
    if n == 0: return ""
    if n == 100: return "CIEN"
    c, r = divmod(n, 100)
    t = C[c]
    if r < 30: s = U[r]
    else:
        d, u = divmod(r, 10); s = D[d] + (" Y " + U[u] if u else "")
    return (t + " " + s).strip()


def letras(n):
    if n == 0: return "CERO"
    mill, resto = divmod(n, 1_000_000)
    miles, uni = divmod(resto, 1000)
    p = []
    if mill: p.append("UN MILLON" if mill == 1 else cien(mill) + " MILLONES")
    if miles: p.append("MIL" if miles == 1 else cien(miles) + " MIL")
    if uni: p.append(cien(uni))
    return " ".join(p).replace("UNO MIL", "UN MIL").replace("UNO MILLONES", "UN MILLONES")


ruta, fila_tot, fila_son = sys.argv[1], int(sys.argv[2]), int(sys.argv[3])
total = load_workbook(ruta, data_only=True)["Presupuesto"].cell(fila_tot, 6).value
wb = load_workbook(ruta)
ent = int(round(total * 100)) // 100; cts = int(round(total * 100)) % 100
c = wb["Presupuesto"].cell(fila_son, 2, f"{letras(ent)} Y {cts:02d}/100 SOLES")
c.comment = Comment("Texto fijo con el total al generar el archivo. Si cambias precios, vuelve a correr poner_son.py.", "NEXO")
wb.save(ruta)
print(c.value)
