// NEXO OBRA (NO) · Lector de metrados (solo lectura).
// Al abrir un modelo cuyo nombre figura en metrados.config.json, recorre los elementos estructurales y escribe
// metrados_<modelo>.json en la carpeta de salida. NO abre transacciones ni modifica el modelo.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Encodings.Web;
using System.Text.Json;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;

namespace NexoMetrados
{
    public class App : IExternalApplication
    {
        const double FT = 0.3048, FT2 = 0.09290304, FT3 = 0.028316846592;
        static string Dir => Path.GetDirectoryName(typeof(App).Assembly.Location);
        static string Log => Path.Combine(Dir, "metrados.log");

        static void Registrar(string m)
        {
            try { File.AppendAllText(Log, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + m + Environment.NewLine); } catch { }
        }

        class Cfg { public string carpetaSalida { get; set; } public List<string> modelos { get; set; } }

        static Cfg Leer()
        {
            var ruta = Path.Combine(Dir, "metrados.config.json");
            if (!File.Exists(ruta)) return null;
            return JsonSerializer.Deserialize<Cfg>(File.ReadAllText(ruta).TrimStart('﻿'));
        }

        static readonly HashSet<string> Hechos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public Result OnStartup(UIControlledApplication app)
        {
            app.ControlledApplication.DocumentOpened += AlAbrir;
            app.Idling += AlEsperar;   // también exporta los modelos que ya estaban abiertos cuando se cargó el complemento
            Registrar("Lector de metrados cargado (v" + typeof(App).Assembly.GetName().Version + ")");
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            app.ControlledApplication.DocumentOpened -= AlAbrir;
            app.Idling -= AlEsperar;
            return Result.Succeeded;
        }

        static void AlEsperar(object s, Autodesk.Revit.UI.Events.IdlingEventArgs e)
        {
            try
            {
                if (s is UIApplication ui)
                    foreach (Document d in ui.Application.Documents) Exportar(d);
            }
            catch (Exception ex) { Registrar("ERROR (espera): " + ex.Message); }
        }

        static void AlAbrir(object s, DocumentOpenedEventArgs e) => Exportar(e.Document);

        static void Exportar(Document doc)
        {
            try
            {
                if (doc == null || doc.IsFamilyDocument || doc.IsLinked) return;
                var clave = doc.PathName ?? doc.Title;
                if (Hechos.Contains(clave)) return;
                Hechos.Add(clave);
                var cfg = Leer();
                var nombre = Path.GetFileNameWithoutExtension(doc.PathName ?? doc.Title);
                if (cfg?.modelos == null || !cfg.modelos.Any(m => string.Equals(m, nombre, StringComparison.OrdinalIgnoreCase))) return;
                var salida = Path.Combine(cfg.carpetaSalida, "metrados_" + nombre + ".json");
                var datos = Extraer(doc, nombre);
                var op = new JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
                File.WriteAllText(salida, JsonSerializer.Serialize(datos, op));
                Registrar($"Metrados escritos: {salida} ({((List<object>)datos["elementos"]).Count} elementos)");
                var sa = Path.Combine(cfg.carpetaSalida, "auditoria_" + nombre + ".json");
                File.WriteAllText(sa, JsonSerializer.Serialize(Auditar(doc), op));
                Registrar($"Auditoría escrita: {sa}");
            }
            catch (Exception ex) { Registrar("ERROR: " + ex); }
        }

        static readonly BuiltInCategory[] CATS = {
            BuiltInCategory.OST_StructuralFoundation, BuiltInCategory.OST_StructuralColumns, BuiltInCategory.OST_StructuralFraming,
            BuiltInCategory.OST_Floors, BuiltInCategory.OST_Walls, BuiltInCategory.OST_Stairs, BuiltInCategory.OST_Columns,
            BuiltInCategory.OST_GenericModel, BuiltInCategory.OST_Roofs, BuiltInCategory.OST_Ramps };

        // ---------------- AUDITORÍA (solo lectura): censo, vínculos, fases, opciones de diseño y solapes reales entre sólidos
        static Solid SolidoUnido(Element el, Options opt)
        {
            Solid u = null;
            void Rec(GeometryElement ge)
            {
                if (ge == null) return;
                foreach (var g in ge)
                {
                    if (g is Solid s && s.Volume > 1e-9)
                    {
                        try { u = u == null ? SolidUtils.Clone(s) : BooleanOperationsUtils.ExecuteBooleanOperation(u, s, BooleanOperationsType.Union); }
                        catch { if (u == null) u = s; }
                    }
                    else if (g is GeometryInstance gi) Rec(gi.GetInstanceGeometry());
                }
            }
            Rec(el.get_Geometry(opt));
            return u;
        }

        static Dictionary<string, object> Auditar(Document doc)
        {
            var r = new Dictionary<string, object>();
            // Censo de todas las categorías de modelo
            var censo = new Dictionary<string, int>();
            foreach (var el in new FilteredElementCollector(doc).WhereElementIsNotElementType())
            {
                var c = el.Category;
                if (c == null || c.CategoryType != CategoryType.Model) continue;
                censo[c.Name] = censo.TryGetValue(c.Name, out var n) ? n + 1 : 1;
            }
            r["censo"] = censo.OrderByDescending(k => k.Value).ToDictionary(k => k.Key, k => k.Value);
            r["vinculos"] = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>()
                .Select(l => new { nombre = l.Name, cargado = l.GetLinkDocument() != null }).ToList();
            r["importados_cad"] = new FilteredElementCollector(doc).OfClass(typeof(ImportInstance)).Count();
            r["opciones_diseno"] = new FilteredElementCollector(doc).OfClass(typeof(DesignOption)).Select(o => o.Name).ToList();
            r["fases"] = doc.Phases.Cast<Phase>().Select(p => p.Name).ToList();

            var opt = new Options { DetailLevel = ViewDetailLevel.Fine };
            var est = new List<Element>();
            foreach (var bic in CATS.Where(b => b != BuiltInCategory.OST_GenericModel))
                est.AddRange(new FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType());
            var info = new List<object>();
            foreach (var el in est)
            {
                var pc = el.get_Parameter(BuiltInParameter.PHASE_CREATED);
                var pd = el.get_Parameter(BuiltInParameter.PHASE_DEMOLISHED);
                info.Add(new
                {
                    id = el.Id.Value,
                    fase = pc?.AsElementId() is ElementId a && a != ElementId.InvalidElementId ? doc.GetElement(a)?.Name : null,
                    demolida = pd?.AsElementId() is ElementId b && b != ElementId.InvalidElementId ? doc.GetElement(b)?.Name : null,
                    opcion = el.DesignOption?.Name,
                    grupo = el.GroupId != ElementId.InvalidElementId ? doc.GetElement(el.GroupId)?.Name : null
                });
            }
            r["elementos"] = info;

            // Solapes: sólidos que se cruzan sin unirse (el volumen común se estaría contando dos veces)
            var sol = new Dictionary<long, Solid>();
            foreach (var el in est) { try { var s = SolidoUnido(el, opt); if (s != null) sol[el.Id.Value] = s; } catch { } }
            var ids = est.Select(e => e.Id).ToList();
            var solapes = new List<object>();
            var vistos = new HashSet<(long, long)>();
            foreach (var el in est)
            {
                if (!sol.TryGetValue(el.Id.Value, out var sa)) continue;
                var bb = el.get_BoundingBox(null);
                if (bb == null) continue;
                var f = new BoundingBoxIntersectsFilter(new Outline(bb.Min, bb.Max));
                foreach (var otro in new FilteredElementCollector(doc, ids).WherePasses(f))
                {
                    long a = el.Id.Value, b = otro.Id.Value;
                    if (a >= b || !vistos.Add((a, b)) || !sol.TryGetValue(b, out var sb)) continue;
                    try
                    {
                        var i = BooleanOperationsUtils.ExecuteBooleanOperation(sa, sb, BooleanOperationsType.Intersect);
                        var v = (i?.Volume ?? 0) * FT3;
                        if (v > 0.001)
                            solapes.Add(new { a, b, volumen_m3 = Math.Round(v, 4), unidos = JoinGeometryUtils.AreElementsJoined(doc, el, otro) });
                    }
                    catch { }
                }
            }
            r["solapes"] = solapes;
            return r;
        }

        static double P(Element el, BuiltInParameter bip)
        {
            var p = el.get_Parameter(bip);
            return p != null && p.HasValue && p.StorageType == StorageType.Double ? p.AsDouble() : 0;
        }

        static Dictionary<string, object> Extraer(Document doc, string nombre)
        {
            var pi = doc.ProjectInformation;
            var r = new Dictionary<string, object>
            {
                ["modelo"] = nombre, ["ruta"] = doc.PathName, ["fecha"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
                ["proyecto"] = pi?.Name, ["cliente"] = pi?.ClientName, ["direccion"] = pi?.Address, ["numero"] = pi?.Number,
                ["niveles"] = new FilteredElementCollector(doc).OfClass(typeof(Level)).Cast<Level>()
                    .OrderBy(l => l.Elevation).Select(l => new { nombre = l.Name, cota = Math.Round(l.Elevation * FT, 4) }).ToList()
            };
            var opt = new Options { DetailLevel = ViewDetailLevel.Fine, ComputeReferences = false, IncludeNonVisibleObjects = false };
            var lista = new List<object>();
            foreach (var bic in CATS)
            {
                foreach (var el in new FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType())
                {
                    try { lista.Add(Elemento(doc, el, opt)); }
                    catch (Exception ex) { Registrar($"Elemento {el.Id.Value}: {ex.Message}"); }
                }
            }
            r["elementos"] = lista;

            var barras = new List<object>();
            foreach (var el in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType())
            {
                try
                {
                    var tipo = doc.GetElement(el.GetTypeId()) as RebarBarType;
                    var host = (el as Rebar)?.GetHostId();
                    barras.Add(new
                    {
                        id = el.Id.Value, tipo = tipo?.Name, diametro_mm = Math.Round((tipo?.BarNominalDiameter ?? 0) * FT * 1000, 2),
                        largo_total_m = Math.Round(P(el, BuiltInParameter.REBAR_ELEM_LENGTH) * FT, 3),
                        cantidad = el.get_Parameter(BuiltInParameter.REBAR_ELEM_QUANTITY_OF_BARS)?.AsInteger() ?? 0,
                        host = host?.Value, hostCat = host != null ? doc.GetElement(host)?.Category?.Name : null
                    });
                }
                catch (Exception ex) { Registrar($"Barra {el.Id.Value}: {ex.Message}"); }
            }
            r["acero"] = barras;
            return r;
        }

        static object Elemento(Document doc, Element el, Options opt)
        {
            var tipo = doc.GetElement(el.GetTypeId());
            double top = 0, bot = 0, lat = 0, vSol = 0;
            void Sol(GeometryElement ge, Transform t)
            {
                if (ge == null) return;
                foreach (var g in ge)
                {
                    if (g is Solid so && so.Volume > 1e-9)
                    {
                        vSol += so.Volume;
                        foreach (Face f in so.Faces)
                        {
                            var bb = f.GetBoundingBox();
                            var uv = (bb.Min + bb.Max) / 2;
                            var n = f.ComputeNormal(uv);
                            if (t != null) n = t.OfVector(n);
                            if (n.Z > 0.7) top += f.Area; else if (n.Z < -0.7) bot += f.Area; else lat += f.Area;
                        }
                    }
                    else if (g is GeometryInstance gi) Sol(gi.GetInstanceGeometry(), null);
                }
            }
            Sol(el.get_Geometry(opt), null);
            var bx = el.get_BoundingBox(null);
            string nivel = null;
            if (el.LevelId != ElementId.InvalidElementId) nivel = doc.GetElement(el.LevelId)?.Name;
            if (nivel == null)
            {
                foreach (var bip in new[] { BuiltInParameter.FAMILY_BASE_LEVEL_PARAM, BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
                                            BuiltInParameter.SCHEDULE_LEVEL_PARAM, BuiltInParameter.WALL_BASE_CONSTRAINT })
                {
                    var p = el.get_Parameter(bip);
                    if (p != null && p.StorageType == StorageType.ElementId && p.AsElementId() != ElementId.InvalidElementId)
                    { nivel = doc.GetElement(p.AsElementId())?.Name; break; }
                }
            }
            var mats = new List<object>();
            foreach (var mid in el.GetMaterialIds(false))
            {
                var v = el.GetMaterialVolume(mid);
                if (v > 0) mats.Add(new { material = doc.GetElement(mid)?.Name, volumen_m3 = Math.Round(v * FT3, 4) });
            }
            var tp = new Dictionary<string, string>();
            foreach (var src in new[] { tipo, el })
            {
                if (src == null) continue;
                foreach (Parameter p in src.Parameters)
                {
                    if (!p.HasValue || p.Definition == null) continue;
                    var k = (src == tipo ? "T:" : "I:") + p.Definition.Name;
                    var val = p.StorageType == StorageType.String ? p.AsString() : p.AsValueString();
                    if (!string.IsNullOrWhiteSpace(val) && !tp.ContainsKey(k) && tp.Count < 120) tp[k] = val;
                }
            }
            return new
            {
                id = el.Id.Value, categoria = el.Category?.Name, bic = ((BuiltInCategory)el.Category.Id.Value).ToString(),
                familia = (tipo as ElementType)?.FamilyName, tipo = tipo?.Name, nivel,
                estructural = (el as FamilyInstance)?.StructuralType.ToString() ?? (el.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT)?.AsInteger() == 1 ? "Estructural" : null),
                volumen_m3 = Math.Round(P(el, BuiltInParameter.HOST_VOLUME_COMPUTED) * FT3, 4),
                volumen_solido_m3 = Math.Round(vSol * FT3, 4),
                area_m2 = Math.Round(P(el, BuiltInParameter.HOST_AREA_COMPUTED) * FT2, 4),
                largo_m = Math.Round(Math.Max(P(el, BuiltInParameter.STRUCTURAL_FRAME_CUT_LENGTH), P(el, BuiltInParameter.CURVE_ELEM_LENGTH)) * FT, 4),
                caras_m2 = new { superior = Math.Round(top * FT2, 4), inferior = Math.Round(bot * FT2, 4), laterales = Math.Round(lat * FT2, 4) },
                caja_m = bx == null ? null : new[] { bx.Min.X * FT, bx.Min.Y * FT, bx.Min.Z * FT, bx.Max.X * FT, bx.Max.Y * FT, bx.Max.Z * FT }.Select(x => Math.Round(x, 4)).ToArray(),
                materiales = mats, parametros = tp
            };
        }
    }
}
