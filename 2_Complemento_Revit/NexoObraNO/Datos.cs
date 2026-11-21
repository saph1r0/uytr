// Lectura: datos del cronograma/presupuesto (JSON) + valores VZ_ actuales del modelo + geometría triangulada.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json.Nodes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace NexoObraNO
{
    public static class Datos
    {
        static JsonNode cache; static DateTime cacheFecha;

        /// <summary>parametros_revit_<modelo>.json (se vuelve a leer si cambió en disco).</summary>
        public static JsonNode Base(Config cfg)
        {
            var f = File.GetLastWriteTimeUtc(cfg.datos);
            if (cache == null || f != cacheFecha) { cache = JsonNode.Parse(File.ReadAllText(cfg.datos)); cacheFecha = f; }
            return cache;
        }

        public static Document Documento(UIApplication app, JsonNode b)
        {
            var doc = app.ActiveUIDocument?.Document ?? throw new Exception("No hay un modelo abierto en Revit.");
            var nombre = Path.GetFileNameWithoutExtension(doc.PathName ?? doc.Title);
            var modelo = (string)b["modelo"];
            if (!string.Equals(nombre, modelo, StringComparison.OrdinalIgnoreCase))
                throw new Exception($"El modelo activo es «{nombre}» y los datos son de «{modelo}». Activa la ventana del modelo correcto.");
            return doc;
        }

        public static string Texto(Element e, string p) { var x = e?.LookupParameter(p); return x == null ? null : (x.AsString() ?? ""); }

        public static object Leer(UIApplication app, Config cfg)
        {
            var b = Base(cfg); var doc = Documento(app, b);
            var faltan = new List<string>();
            var muestra = doc.GetElement(new ElementId((long)b["elementos"][0]["id"]));
            foreach (var p in Escritura.Nuevos.Select(n => n.nombre).Where(n => n != "VZ_Registro_Global").Concat(new[] { "VZ_Paquete" }))
                if (muestra?.LookupParameter(p) == null) faltan.Add(p);
            var els = new List<object>();
            foreach (var n in b["elementos"].AsArray())
            {
                var e = doc.GetElement(new ElementId((long)n["id"]));
                els.Add(new
                {
                    id = (long)n["id"], existe = e != null, nivel = (string)n["nivel"], tipo = (string)n["tipo"], grupo = (string)n["grupo"],
                    paquete = (string)n["paquete"], sector = (string)n["sector"], actividad = (string)n["actividad"], fecha_prog = (string)n["fecha_prog"],
                    partidas = n["partidas"], registro = Texto(e, "VZ_Registro") ?? "", obs = Texto(e, "VZ_Observaciones") ?? ""
                });
            }
            return new
            {
                ok = true, modelo = (string)b["modelo"], inicio = (string)b["inicio_obra"], pie = b["pie"], niveles = b["niveles"],
                capitulos = b["capitulos"], partidas = b["partidas"], elementos = els,
                registroGlobal = Texto(doc.ProjectInformation, "VZ_Registro_Global") ?? "", parametrosFaltan = faltan
            };
        }

        /// <summary>Malla de cada elemento en metros: {"elementos":[{"id":..,"v":[x,y,z, ...]}]} (triángulos sin índice).</summary>
        public static string Geometria(UIApplication app, Config cfg)
        {
            var b = Base(cfg); var doc = Documento(app, b);
            const double FT = 0.3048; var ci = CultureInfo.InvariantCulture;
            var opt = new Options { DetailLevel = ViewDetailLevel.Medium };
            var sb = new StringBuilder("{\"elementos\":[");
            bool primero = true;
            foreach (var n in b["elementos"].AsArray())
            {
                var e = doc.GetElement(new ElementId((long)n["id"]));
                if (e == null) continue;
                var v = new StringBuilder();
                void Rec(GeometryElement ge)
                {
                    if (ge == null) return;
                    foreach (var g in ge)
                    {
                        if (g is Solid s && s.Volume > 1e-9)
                            foreach (Face f in s.Faces)
                            {
                                Mesh m; try { m = f.Triangulate(); } catch { continue; }
                                for (int i = 0; i < m.NumTriangles; i++)
                                {
                                    var t = m.get_Triangle(i);
                                    for (int k = 0; k < 3; k++)
                                    {
                                        var p = t.get_Vertex(k);
                                        if (v.Length > 0) v.Append(',');
                                        v.Append(Math.Round(p.X * FT, 3).ToString(ci)).Append(',').Append(Math.Round(p.Y * FT, 3).ToString(ci)).Append(',').Append(Math.Round(p.Z * FT, 3).ToString(ci));
                                    }
                                }
                            }
                        else if (g is GeometryInstance gi) Rec(gi.GetInstanceGeometry());
                    }
                }
                try { Rec(e.get_Geometry(opt)); } catch (Exception ex) { Registro.Escribir("Geometría " + n["id"] + ": " + ex.Message); }
                if (v.Length == 0) continue;
                if (!primero) sb.Append(','); primero = false;
                sb.Append("{\"id\":").Append((long)n["id"]).Append(",\"v\":[").Append(v).Append("]}");
            }
            return sb.Append("]}").ToString();
        }
    }
}
