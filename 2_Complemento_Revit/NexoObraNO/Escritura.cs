// Escritura del metrado diario en los parámetros VZ_ (una sola transacción por envío; «simular» la revierte).
//  · VZ_Registro (texto): historial «aaaa-mm-dd|ítem|cantidad;...» — fuente de las valorizaciones mensuales.
//  · VZ_Ejec_Concreto_m3 / VZ_Ejec_Encofrado_m2 / VZ_Ejec_Acero_kg: acumulados (se suman a lo ya existente).
//  · VZ_Observaciones: alertas de mayor metrado (cuando el acumulado supera el contrato).
//  · VZ_Registro_Global (Información de proyecto): partidas sin elementos en el modelo (provisionales, excavación, ...).
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json.Nodes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace NexoObraNO
{
    public class Registro1 { public long id { get; set; } public string item { get; set; } public double cantidad { get; set; } }
    public class PedidoMetrado { public string fecha { get; set; } public bool simular { get; set; } public List<Registro1> registros { get; set; } }

    public static class Escritura
    {
        static readonly CultureInfo CI = CultureInfo.InvariantCulture;
        public static readonly (string nombre, ForgeTypeId tipo, string desc)[] Nuevos =
        {
            ("VZ_Ejec_Concreto_m3", SpecTypeId.Volume, "Concreto ejecutado acumulado (suma de los metrados diarios)"),
            ("VZ_Ejec_Encofrado_m2", SpecTypeId.Area, "Encofrado ejecutado acumulado (suma de los metrados diarios)"),
            ("VZ_Ejec_Acero_kg", SpecTypeId.Number, "Acero ejecutado acumulado en kg (suma de los metrados diarios)"),
            ("VZ_Registro", SpecTypeId.String.Text, "Historial de metrados diarios: aaaa-mm-dd|ítem|cantidad;..."),
            ("VZ_Observaciones", SpecTypeId.String.Text, "Alertas de mayor metrado y notas"),
            ("VZ_Registro_Global", SpecTypeId.String.Text, "Metrados diarios de partidas sin elementos en el modelo"),
        };
        static readonly BuiltInCategory[] CATS = { BuiltInCategory.OST_StructuralFoundation, BuiltInCategory.OST_StructuralColumns,
            BuiltInCategory.OST_StructuralFraming, BuiltInCategory.OST_Floors, BuiltInCategory.OST_Walls, BuiltInCategory.OST_Stairs, BuiltInCategory.OST_Ramps };

        public static Dictionary<string, double> Sumas(string reg)
        {
            var d = new Dictionary<string, double>();
            foreach (var x in (reg ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var p = x.Split('|');
                if (p.Length == 3 && double.TryParse(p[2], NumberStyles.Float, CI, out var q)) d[p[1]] = (d.TryGetValue(p[1], out var a) ? a : 0) + q;
            }
            return d;
        }

        static void AsegurarParametros(UIApplication app, Document doc, List<string> notas)
        {
            var existentes = new HashSet<string>();
            var it = doc.ParameterBindings.ForwardIterator(); it.Reset();
            while (it.MoveNext()) existentes.Add(it.Key.Name);
            var falta = Nuevos.Where(n => !existentes.Contains(n.nombre)).ToList();
            if (falta.Count == 0) return;
            var a = app.Application; var original = a.SharedParametersFilename;
            try
            {
                var cfgSp = Config.Leer().parametrosCompartidos;
                if (!File.Exists(cfgSp)) File.WriteAllText(cfgSp, "");
                a.SharedParametersFilename = cfgSp;
                var archivo = a.OpenSharedParameterFile();
                var grupo = archivo.Groups.get_Item("NEXO OBRA (NO)") ?? archivo.Groups.Create("NEXO OBRA (NO)");
                var cats = a.Create.NewCategorySet();
                foreach (var b in CATS) { var c = Category.GetCategory(doc, b); if (c != null && c.AllowsBoundParameters) cats.Insert(c); }
                var catProy = a.Create.NewCategorySet(); catProy.Insert(Category.GetCategory(doc, BuiltInCategory.OST_ProjectInformation));
                foreach (var n in falta)
                {
                    var def = grupo.Definitions.get_Item(n.nombre) ?? grupo.Definitions.Create(new ExternalDefinitionCreationOptions(n.nombre, n.tipo) { Description = n.desc });
                    var vinc = a.Create.NewInstanceBinding(n.nombre == "VZ_Registro_Global" ? catProy : cats);
                    doc.ParameterBindings.Insert(def, vinc, GroupTypeId.IdentityData);
                    notas.Add("Parámetro creado: " + n.nombre);
                }
            }
            finally { a.SharedParametersFilename = original; }
            doc.Regenerate();
        }

        static void Poner(Element e, string n, object v)
        {
            var p = e.LookupParameter(n);
            if (p == null || p.IsReadOnly) return;
            if (p.StorageType == StorageType.String) p.Set(Convert.ToString(v, CI));
            else if (p.StorageType == StorageType.Double) p.Set(Convert.ToDouble(v, CI));
            else if (p.StorageType == StorageType.Integer) p.Set(Convert.ToInt32(v, CI));
        }

        public static object Registrar(UIApplication app, Config cfg, PedidoMetrado ped)
        {
            var b = Datos.Base(cfg); var doc = Datos.Documento(app, b);
            if (!DateTime.TryParseExact(ped?.fecha ?? "", "yyyy-MM-dd", CI, DateTimeStyles.None, out var fecha)) throw new Exception("Fecha no válida (aaaa-mm-dd).");
            var inicio = DateTime.ParseExact((string)b["inicio_obra"], "yyyy-MM-dd", CI);
            var partidas = b["partidas"].AsArray().ToDictionary(p => (string)p["item"], p => p);
            var conElementos = new HashSet<string>();
            var elems = new Dictionary<long, JsonNode>();
            foreach (var n in b["elementos"].AsArray())
            {
                elems[(long)n["id"]] = n;
                foreach (var p in n["partidas"].AsArray()) conElementos.Add((string)p["item"]);
            }
            var aplicados = new List<object>(); var rechazados = new List<object>(); var alertas = new List<string>(); var notas = new List<string>();
            var actualizados = new Dictionary<long, Element>();
            string f = fecha.ToString("yyyy-MM-dd"), fdmy = fecha.ToString("dd/MM/yyyy");
            int nVal = (fecha.Year - inicio.Year) * 12 + fecha.Month - inicio.Month + 1;

            using var t = new Transaction(doc, $"NEXO OBRA: metrado diario {fdmy}");
            t.Start();
            AsegurarParametros(app, doc, notas);
            foreach (var r in ped.registros ?? new List<Registro1>())
            {
                if (r.cantidad == 0 || double.IsNaN(r.cantidad)) { rechazados.Add(new { r.id, r.item, motivo = "Cantidad 0 o no válida." }); continue; }
                if (!partidas.TryGetValue(r.item ?? "", out var part)) { rechazados.Add(new { r.id, r.item, motivo = "La partida no existe en el presupuesto." }); continue; }
                string und = (string)part["und"];
                if (r.id == 0)
                {   // partida sin elementos en el modelo
                    if (conElementos.Contains(r.item)) { rechazados.Add(new { r.id, r.item, motivo = "Esta partida tiene elementos: registra por elemento o por partida completa." }); continue; }
                    var pi = doc.ProjectInformation; var reg = Datos.Texto(pi, "VZ_Registro_Global") ?? "";
                    double ant = Sumas(reg).TryGetValue(r.item, out var a0) ? a0 : 0, nuevo = ant + r.cantidad, contrato = (double)part["metrado"];
                    if (nuevo < -1e-9) { rechazados.Add(new { r.id, r.item, motivo = $"El acumulado quedaría negativo ({nuevo:0.###})." }); continue; }
                    Poner(pi, "VZ_Registro_Global", reg + $"{f}|{r.item}|{r.cantidad.ToString("R", CI)};");
                    if (nuevo > contrato * (1 + 1e-6) + 1e-9) alertas.Add($"{r.item}: acumulado {nuevo:0.###} {und} supera el contrato ({contrato:0.###} {und}).");
                    aplicados.Add(new { r.id, r.item, r.cantidad, acumulado = nuevo, contrato });
                    continue;
                }
                if (!elems.TryGetValue(r.id, out var de)) { rechazados.Add(new { r.id, r.item, motivo = "El elemento no está en los datos del cronograma." }); continue; }
                var pe = de["partidas"].AsArray().FirstOrDefault(p => (string)p["item"] == r.item);
                if (pe == null) { rechazados.Add(new { r.id, r.item, motivo = "Esa partida no corresponde a este elemento." }); continue; }
                var e = doc.GetElement(new ElementId(r.id));
                if (e == null) { rechazados.Add(new { r.id, r.item, motivo = "El elemento ya no existe en el modelo." }); continue; }
                var regE = Datos.Texto(e, "VZ_Registro") ?? "";
                var sum = Sumas(regE);
                double antE = sum.TryGetValue(r.item, out var a1) ? a1 : 0, nuevoE = antE + r.cantidad, contE = (double)pe["metrado"];
                if (nuevoE < -1e-9) { rechazados.Add(new { r.id, r.item, motivo = $"El acumulado quedaría negativo ({nuevoE:0.###})." }); continue; }
                Poner(e, "VZ_Registro", regE + $"{f}|{r.item}|{r.cantidad.ToString("R", CI)};");
                sum[r.item] = nuevoE;
                string clase = (string)pe["clase"];
                // acumulados (se suman a lo existente) y fecha de la primera ejecución de cada etapa
                if (clase == "concreto") Poner(e, "VZ_Ejec_Concreto_m3", UnitUtils.ConvertToInternalUnits(nuevoE, UnitTypeId.CubicMeters));
                if (clase == "encofrado") Poner(e, "VZ_Ejec_Encofrado_m2", UnitUtils.ConvertToInternalUnits(nuevoE, UnitTypeId.SquareMeters));
                if (clase == "acero") Poner(e, "VZ_Ejec_Acero_kg", nuevoE);
                string campoFecha = clase == "acero" || clase == "ladrillo" ? "VZ_Fecha_Acero" : clase == "encofrado" ? "VZ_Fecha_Encofrado" : "VZ_Fecha_Concreto";
                if (string.IsNullOrWhiteSpace(Datos.Texto(e, campoFecha))) Poner(e, campoFecha, fdmy);
                if (nuevoE > contE * (1 + 1e-6) + 1e-9)
                {
                    string msg = $"{fdmy} · mayor metrado {r.item}: acumulado {nuevoE:0.###} {und} > contrato {contE:0.###} {und}";
                    alertas.Add($"Elemento {r.id} ({de["tipo"]}): {msg}");
                    var obs = Datos.Texto(e, "VZ_Observaciones") ?? "";
                    Poner(e, "VZ_Observaciones", (obs.Length > 0 ? obs + " | " : "") + msg);
                }
                // monto valorizado y estado del elemento
                double monto = 0; bool principalCompleto = false, alguno = false;
                foreach (var p in de["partidas"].AsArray())
                {
                    string it = (string)p["item"]; double q = sum.TryGetValue(it, out var s) ? s : 0;
                    monto += q * (double)partidas[it]["pu"]; alguno |= q > 1e-9;
                    if (((string)p["clase"] == "concreto" || (string)p["clase"] == "muro") && q >= (double)p["metrado"] * 0.999) principalCompleto = true;
                }
                Poner(e, "VZ_Monto_Valorizado", Math.Round(monto, 2));
                Poner(e, "VZ_Estado", principalCompleto ? "Ejecutado" : alguno ? "En ejecución" : "Programado");
                Poner(e, "VZ_Valorizacion_N", nVal);
                actualizados[r.id] = e;
                aplicados.Add(new { r.id, r.item, r.cantidad, acumulado = nuevoE, contrato = contE });
            }
            if (ped.simular || aplicados.Count == 0) t.RollBack(); else t.Commit();
            Registro.Escribir($"Metrado {fdmy}: {aplicados.Count} aplicados, {rechazados.Count} rechazados, simular={ped.simular}");
            return new
            {
                ok = true, simulado = ped.simular, guardado = !ped.simular && aplicados.Count > 0, aplicados, rechazados, alertas, notas,
                elementos = ped.simular ? new List<object>() : actualizados.Select(kv => (object)new { id = kv.Key, registro = Datos.Texto(kv.Value, "VZ_Registro"), obs = Datos.Texto(kv.Value, "VZ_Observaciones") }).ToList(),
                registroGlobal = Datos.Texto(doc.ProjectInformation, "VZ_Registro_Global") ?? ""
            };
        }
    }
}
