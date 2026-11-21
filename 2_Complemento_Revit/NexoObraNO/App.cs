// NEXO OBRA (NO) · complemento de Revit: arranque, configuración, registro y botón «Abrir panel».
using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace NexoObraNO
{
    public class Config
    {
        public int puerto { get; set; } = 8770;
        public string datos { get; set; }                  // parametros_revit_<modelo>.json (cronograma + presupuesto)
        public string panel { get; set; }                  // index.html del panel
        public string parametrosCompartidos { get; set; }  // Parametros_NEXO_NO.txt
        public int esperaRevitMs { get; set; } = 60000;

        public static string Carpeta => Path.GetDirectoryName(typeof(Config).Assembly.Location);
        public static Config Leer()
        {
            var ruta = Path.Combine(Carpeta, "nexo_no.config.json");
            return JsonSerializer.Deserialize<Config>(File.ReadAllText(ruta).TrimStart('﻿'));
        }
    }

    public static class Registro
    {
        public static void Escribir(string m)
        {
            try { File.AppendAllText(Path.Combine(Config.Carpeta, "nexo_no.log"), DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + m + Environment.NewLine); }
            catch { }
        }
    }

    public class App : IExternalApplication
    {
        public static Servidor Servidor;
        public static Puente Puente;

        public Result OnStartup(UIControlledApplication app)
        {
            try
            {
                var cfg = Config.Leer();
                Puente = new Puente(); Puente.Crear();
                Servidor = new Servidor(cfg, Puente);
                Servidor.Iniciar();
                try
                {
                    const string tab = "NEXO OBRA (NO)";
                    try { app.CreateRibbonTab(tab); } catch { }
                    var panel = app.CreateRibbonPanel(tab, "Valorizaciones");
                    panel.AddItem(new PushButtonData("NexoAbrir", "Abrir\npanel", typeof(App).Assembly.Location, typeof(CmdAbrir).FullName)
                    { ToolTip = "Abre el panel NEXO OBRA (NO) en el navegador: " + Servidor.Url });
                }
                catch (Exception ex) { Registro.Escribir("Cinta: " + ex.Message); }
                Registro.Escribir("Complemento cargado (v" + typeof(App).Assembly.GetName().Version + ") en " + Servidor.Url);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                Registro.Escribir("ERROR al iniciar: " + ex);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            Servidor?.Detener();
            Puente?.Liberar();
            return Result.Succeeded;
        }
    }

    [Transaction(TransactionMode.ReadOnly)]
    public class CmdAbrir : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string msg, ElementSet els)
        {
            Process.Start(new ProcessStartInfo(App.Servidor.Url) { UseShellExecute = true });
            return Result.Succeeded;
        }
    }
}
