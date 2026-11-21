# Especificación del panel NEXO OBRA (NO)

Regla: todo lo que se muestra sale de datos existentes — presupuesto S10 v2, cronograma (`cronograma_nexo.py`),
modelo de Revit (parámetros VZ_ y geometría) y metrados diarios registrados (`VZ_Registro`, `VZ_Registro_Global`,
`VZ_Observaciones`). Lo que no existe (fotos, costo real, cuaderno, reajuste) no se muestra; los datos que escribe
el usuario (empresa, firmas, amortizaciones, estado de valorizaciones) se marcan como tales.

## Distribución general
- Riel izquierdo con 8 páginas. Barra superior: proyecto, **fecha de corte** (manda en todos los cálculos), estado de Revit, recargar.
- Un solo visor 3D (WebGL) que se monta en Inicio, Visor y Partidas.

| # | Página | Contenido |
|---|---|---|
| 1 | Inicio | 5 tarjetas (presupuesto, valorizado, programado, SPI, plazo) · curva S programado vs valorizado · avance por capítulo · 3D por avance · resumen del mes |
| 2 | Visor 3D | Colores: avance, sector, programado al corte, atraso · niveles · timeline programado · avance por bloque · detalle del elemento |
| 3 | Partidas | Lista con filtros (capítulo, estado) · detalle · registro por elemento o por partida completa (con confirmación) · historial · 3D |
| 4 | Parte diario | Calendario del mes (días con registro, domingos y feriados) · parte del día por partida · ver en 3D · imprimir |
| 5 | Cronograma | Gantt con avance real dentro de cada barra · hitos · detalle de actividad con sus partidas y elementos |
| 6 | Valorizaciones | Control N° 1..n (programado, valorizado, acumulado, %, SPI, estado) · curva S y barras · formato · hoja de metrados · expediente PDF · CSV |
| 7 | Control | Atrasados (fecha programada vencida sin metrado) · mayor metrado · desvío del mes por partida · materiales del mes siguiente (APU) |
| 8 | Configuración | Conexión, fuentes, calendario, datos de la empresa para el formato, instantánea de solo lectura |

## Expediente de valorización (PDF)
Carátula · resumen (indicadores y curva S) · formato de valorización · hoja de metrados (elemento, nivel, fecha) ·
observaciones de mayor metrado · firmas.

## Compartir para que otros vean y sugieran (sin modificar)
1. **Instantánea de solo lectura** (implementada): Configuración → «Descargar instantánea». Es un HTML con los datos y el 3D
   incrustados, sin botones de registro. Se puede publicar como página privada de claude.ai y compartir el enlace; quienes
   la abren pueden dejar **comentarios** (sugerencias) que luego se leen y se responden. Se actualiza volviendo a publicar.
2. **Red local en vivo**: el complemento atendería también la red de la oficina en modo solo lectura y guardaría las
   sugerencias en un archivo. Requiere permiso de administrador (reserva de URL y firewall) — decisión del usuario.
3. **Nube en vivo** (por ejemplo una base de datos en línea con usuarios y roles): acceso desde cualquier lugar; requiere cuenta y costo.
