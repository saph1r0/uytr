# Plan de ejecución y control · NEXO OBRA (NO)

Edificio multifamiliar con 3 sótanos (modelo **REP0-MOD2-R02**), casco estructural. Arequipa.
Archivos: `Cronograma_Valorizado_REP0-MOD2-R02.xlsx`, `nexo_no_datos.json` (también en `1_Panel\datos\`), `Sectorizacion\*.svg`.
Generado con `4_Python\cronograma_nexo.py` a partir del modelo y del presupuesto `..._v2.xlsx`.

## 1. Resumen
| Dato | Valor |
|---|---|
| Inicio (supuesto) | lunes 04/01/2027 |
| Fin del casco | 11/11/2027 |
| Plazo | 312 días calendario · 11 valorizaciones mensuales |
| Calendario | **lunes a sábado**; domingo y feriados no laborables (incluye 15/08 Arequipa y 25/10 Construcción Civil) |
| Costo directo programado | S/ 2,410,991.60 (cuadra con el presupuesto; diferencia S/ 0.23 por redondeo) |
| Personal promedio | 2 a 22 obreros/día (más subcontratos de excavación y anclajes) |

## 2. Estructura de desglose (EDT)
1. **Obras provisionales** — cerco, oficinas, instalaciones, cartel, movilización · SST y trazo continuos.
2. **Sostenimiento y excavación** (de arriba hacia abajo)
   - Anillo 1 (semisótano): excavación capa 1 → anclajes fila 1 → muros anclados → tensado (7 días después).
   - Anillo 2 (sótano 1): excavación capa 2 → anclajes fila 2 → muros anclados y sus cimientos → tensado.
   - Capa 3: excavación del sótano 2.
3. **Cimentación** (sectores A y B): excavación manual, solado, subzapatas, zapatas, losas de cimentación y cisterna.
4. **Estructura** (de abajo hacia arriba), 10 paquetes `T01…T10`. Cada paquete = verticales de un nivel + techo del nivel superior.
5. **Cierre de sótanos**: impermeabilización de cisterna, afirmado y falso piso.

## 3. Sectorización
- Cada paquete se divide en **2 sectores (A y B)** con una junta perpendicular al eje largo (eje Y del modelo).
- La junta se ubica donde el concreto queda más parejo; por las placas grandes (no se cortan) el reparto queda ~56/44 %.
- Muros anclados: **paños alternados** dentro de cada anillo (no se sectorizan en A/B).
- Planos: `Sectorizacion\sectorizacion_T01.svg … T10.svg` y `_CIM.svg` (azul = A, dorado = B, rojo = junta).
- **Confirmar con el ingeniero estructural**: la junta va en el tercio central del paño de losa o viga, nunca sobre una placa o columna.

## 4. Tren de trabajo por sector
`Acero vertical → Encofrado vertical → Concreto vertical → Encofrado de techo → Acero de techo y ladrillo → Concreto de techo`
- El sector B sigue al A en cada actividad (la misma cuadrilla pasa de A a B).
- El nivel siguiente empieza en un sector 1 día después de vaciar su techo.
- Cuadrillas por defecto: acero 6, encofrado 8, vaciado 1 (editables en la hoja **Parámetros**). Ciclo típico ≈ 14 días por piso.
- La duración de cada actividad = cuadrilla-días (metrado del modelo ÷ rendimiento del APU) ÷ N° de cuadrillas.

## 5. Contenido del Excel
| Hoja | Para qué |
|---|---|
| Cronograma | 140 actividades con predecesoras, fechas por fórmula (`WORKDAY.INTL`, lunes a sábado), Gantt semanal y recursos por actividad |
| Cronograma Valorizado | Costo directo por mes (valorización N° 1…11), acumulado, curva S, con GG, utilidad e IGV |
| Adquisiciones y Recursos | Materiales, horas-hombre y equipos por mes; histograma de personal |
| Sectorización | Paquetes, sectores, junta, concreto y encofrado por sector |
| Metrados por Actividad | Cada partida repartida en actividades/sectores según el modelo (base para valorizar por sector) |
| Presupuesto Base | Copia del presupuesto v2 con control de que el metrado repartido cuadra |
| Coeficientes APU | Recursos por unidad de partida (del APU) |
| Formato Valorización | Plantilla: anterior, del mes, acumulado, saldo, % avance, amortizaciones, retención y neto |
| Parámetros | Inicio, cuadrillas, feriados, GG/UT/IGV |

## 6. Reglas de medición para valorizar (propuesta)
| Partida | Se valoriza cuando |
|---|---|
| Concreto | 100 % del elemento al vaciarlo |
| Encofrado | 100 % al encofrar (incluye desencofrado) |
| Acero | 100 % al colocarlo y ser aprobado por la supervisión |
| Excavación / eliminación | m3 ejecutados por capa (topografía) |
| Anclajes | 80 % al perforar e inyectar · 20 % al tensar |
| Obras provisionales | al instalarse · SST por mes |

## 7. Qué necesita la app NEXO OBRA (NO)
1. **Datos base** (`nexo_no_datos.json`): partidas con PU, paquetes y sectores, 140 actividades con fechas y monto, periodos con valor programado (PV) y **cada ElementId de Revit con su paquete y sector**.
2. **Registro de avance**: por elemento (desde el visor 3D) o por sector/actividad, con fecha y responsable. Regla de medición de la sección 6.
3. **Valorización mensual**: metrado del mes por partida = suma de lo ejecutado en el periodo; formato de la hoja «Formato Valorización».
4. **Control de costos**: costo real (AC) por partida/actividad (facturas, planillas, partes diarios) frente al valor ganado (EV) y al programado (PV): SPI, CPI, EAC.
5. **Visor 3D** coloreado por: programado / ejecutado / valorizado / por valorizar, y por sector.
6. **Calendario configurable** (días laborables y feriados), sin el error del panel ND (sábado no laborable).
7. **Reportes**: valorización N° X en PDF/Excel con formato S10, curva S real vs. programada, adquisiciones del mes siguiente.

## 8. Supuestos y riesgos
- Fecha de inicio, cuadrillas y feriados son supuestos (hoja Parámetros).
- La partida de SST del presupuesto está a 12 meses; el cronograma da ~10.5: ajustar uno de los dos.
- Excavación a 350 m3/día con 1 excavadora y anclajes a 2 und/día por equipo: confirmar con los subcontratistas.
- El mayor riesgo de plazo es el sostenimiento (anclajes y tensado) y la autorización municipal de excavación.
- No incluye calzaduras, juntas sísmicas, tanque elevado ni caja de ascensor (no modelados).
