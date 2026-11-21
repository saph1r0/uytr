# NEXO OBRA (NO) · Valorizaciones de obra

Panel para preparar, controlar y reportar **valorizaciones** a partir del modelo de Revit.
Parte de NEXO OBRA (ND), cuyo respaldo intacto está en:
`D:\PANEL DE OBRA BIM\Respaldos\NEXO_OBRA_ND_v1_2026-11-19\` (panel, complemento 0.6 instalado y Python).

## Carpetas

| Carpeta | Para qué |
|---|---|
| `1_Panel` | Página del panel (HTML/JS) |
| `2_Complemento_Revit` | Complemento C# / .NET 10 para Revit 2027 |
| `3_Modelo_y_parametros` | Modelo de prueba y parámetros compartidos |
| `4_Python` | Motor de cálculo, emulador y pruebas |
| `5_Documentacion` | Guías, referencia técnica, PDF |
| `6_Reutilizado_de_ND` | Piezas copiadas de ND antes de adaptarlas |
| `Respaldos` | Versiones anteriores |

## Qué reutilizar de ND

Leyenda: ✅ tal cual · 🔧 adaptar · ❌ no se usa

### Visualización
| Pieza | Dónde está en ND | Uso | Nota para NO |
|---|---|---|---|
| Visor 3D propio (`V3`) | dashboard.html ~L2257 | ✅ | Colorear por **valorizado / por valorizar / valorizado este periodo** |
| Geometría desde Revit (`/api/geometria`) + caché local | `Consultas.cs`, `GEO_KEY` | ✅ | |
| Reloj de simulación y línea de tiempo (`crearReloj`, `TL`) | ~L2207, ~L2476 | 🔧 | Avanzar **por valorización (mes)**, no por día |
| Simulación 4D (`G4`) | ~L2573 | 🔧 | Pasa a ser «5D»: monto acumulado en el tiempo |
| Gráficos Chart.js incrustado + `baseOpts` | ~L1669 | ✅ | Curva S de valorizaciones, barras por partida |
| Anillo de avance (`ringSvg`) | ~L1377 | ✅ | Avance valorizado vs. programado |

### Enlace con Revit
| Pieza | Uso | Nota para NO |
|---|---|---|
| Servidor local 127.0.0.1 + `ExternalEvent` (`ServidorHttp.cs`, `PuenteRevit`) | ✅ | Base sólida y probada |
| Lectura `/api/elementos`, `/api/estado`, `/api/ping` | ✅ | |
| Escritura `POST /api/cambios` con lista blanca y una sola transacción (`Escritura.cs`) | 🔧 | Cambiar la lista: metrado ejecutado, % valorizado, N.º de valorización |
| Mapeo de parámetros y perfiles (`MapeoParametros.cs`, perfiles SC/AV) | ✅ | Crear perfil **VZ_** (valorización) |
| Arranque automático, validación de configuración, registro | ✅ | |
| Instalador `instalar.ps1` | 🔧 | Nombre y carpeta nuevos para no chocar con ND |
| Conflictos «gana Revit y se avisa», sincronización idempotente | ✅ | |

### Datos y cálculo
| Pieza | Uso | Nota para NO |
|---|---|---|
| Lectura de fechas y números (`parseDate`, `parseNum`) | ✅ | Añadir aviso si una fecha no se lee |
| Calendario laborable (`lab`, `diaLab`, `addLab`) | 🔧 | **Corregir antes**: días configurables (lun–sáb por defecto) |
| Partidas con precio unitario y BAC | 🔧 | Base del **presupuesto contractual** |
| Valor ganado (EV, CPI, SPI) | 🔧 | Solo como indicador; **una sola fuente de PV** |
| Estados por elemento | 🔧 | Simplificar: Por valorizar / Parcial / Valorizado |
| Importar tabla exportada de Revit (`.txt`) | ✅ | Modo sin enlace |
| Guardado local + respaldo descargable/cargable | ✅ | Clave nueva (`nexoNO_…`) para no mezclar con ND |
| Fecha de corte editable (`setCorte`) | ✅ | Pasa a ser **fecha de cierre de la valorización** |

### Interfaz y reportes
| Pieza | Uso | Nota para NO |
|---|---|---|
| Diseño NEXO (Midnight Navy & Champagne Gold), riel izquierdo, modo claro/oscuro | ✅ | Mostrar los subtítulos de tarjetas (en ND quedaron ocultos) |
| Tarjetas KPI, tablas, píldoras de estado, diálogos, avisos | ✅ | |
| Ayuda ⓘ por indicador | ✅ | Completar todas desde el inicio |
| Informe PDF imprimible (`buildReport`, `chartImg`) | 🔧 | Formato de **valorización** (ver abajo) |
| Réplica Python + pruebas de paridad + emulador | ✅ | Probar cada fórmula nueva en JS y Python |

### Lo que NO se lleva
- ❌ Productividad (HH), Incidencias, Programación detallada: no son de valorizaciones.
- ❌ «Culminado en sábado» y la regla del sábado como colchón.
- ❌ Doble ingreso del PV (partida + semana).
- ❌ Textos fijos de «Modelo_1» y «SC - SEGUIMIENTO Y CONTROL».

## Lo nuevo que necesita NO (no existe en ND)
1. **Metrado ejecutado por elemento y periodo** (avance parcial, no solo 0/100 %).
2. **Valorización N.º X**: anterior, actual, acumulado y saldo por partida.
3. **Presupuesto contractual** con adicionales y deductivos.
4. **Reajuste por fórmula polinómica** (índices del INEI) — opcional.
5. **Amortización de adelantos** (directo y de materiales).
6. **Retención de garantía, gastos generales, utilidad e IGV**.
7. **Cuadro resumen** tipo hoja de valorización para firmar y exportar a PDF/Excel.
8. **Validación metrado modelo vs. contractual** con umbral (ej. ±3 %).

## Orden sugerido
1. Copiar a `6_Reutilizado_de_ND` las piezas ✅ y separar el HTML de ND en módulos (visor, enlace, gráficos, diseño).
2. Corregir calendario y fuente única de PV.
3. Definir los parámetros VZ_ y el perfil de mapeo.
4. Construir la hoja de valorización (puntos 1, 2, 6 y 7).
5. Visor 3D coloreado por valorización.
6. Adelantos, reajustes y adicionales.
