// Servidor HTTP local (solo 127.0.0.1): panel, /api/ping, /api/datos, /api/geometria y POST /api/metrado.
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Threading.Tasks;

namespace NexoObraNO
{
    public class Servidor
    {
        public static readonly JsonSerializerOptions Json = new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
        readonly Config cfg; readonly Puente puente; HttpListener l;
        string geometriaCache;   // la geometría se calcula una vez por sesión

        public Servidor(Config cfg, Puente puente) { this.cfg = cfg; this.puente = puente; }
        public string Url => $"http://127.0.0.1:{cfg.puerto}/";

        public void Iniciar()
        {
            l = new HttpListener(); l.Prefixes.Add(Url); l.Start();
            Task.Run(Bucle);
        }
        public void Detener() { try { l?.Stop(); l?.Close(); } catch { } l = null; }

        async Task Bucle()
        {
            while (l != null && l.IsListening)
            {
                HttpListenerContext c;
                try { c = await l.GetContextAsync(); } catch { break; }
                _ = Task.Run(() => Atender(c));
            }
        }

        async Task Atender(HttpListenerContext ctx)
        {
            var req = ctx.Request; var res = ctx.Response;
            try
            {
                if (req.RemoteEndPoint == null || !IPAddress.IsLoopback(req.RemoteEndPoint.Address)) { await Enviar(res, 403, Error("Solo conexiones locales.")); return; }
                string ruta = req.Url.AbsolutePath.TrimEnd('/').ToLowerInvariant();
                if (req.HttpMethod == "GET" && (ruta == "" || ruta == "/index.html"))
                    await EnviarBytes(res, 200, "text/html; charset=utf-8", await File.ReadAllBytesAsync(cfg.panel));
                else if (req.HttpMethod == "GET" && ruta == "/api/ping")
                    await Enviar(res, 200, new { ok = true, version = typeof(Servidor).Assembly.GetName().Version.ToString(3) });
                else if (req.HttpMethod == "GET" && ruta == "/api/datos")
                    await Enviar(res, 200, await puente.EjecutarAsync(app => Datos.Leer(app, cfg), cfg.esperaRevitMs));
                else if (req.HttpMethod == "GET" && ruta == "/api/geometria")
                {
                    if (geometriaCache == null || req.QueryString["recargar"] == "1")
                        geometriaCache = (string)await puente.EjecutarAsync(app => Datos.Geometria(app, cfg), cfg.esperaRevitMs * 3);
                    await EnviarBytes(res, 200, "application/json; charset=utf-8", Encoding.UTF8.GetBytes(geometriaCache));
                }
                else if (req.HttpMethod == "POST" && ruta == "/api/metrado")
                {
                    string origen = req.Headers["Origin"];
                    if (!string.IsNullOrEmpty(origen) && !string.Equals(origen.TrimEnd('/'), Url.TrimEnd('/'), StringComparison.OrdinalIgnoreCase))
                    { await Enviar(res, 403, Error("Origen no permitido: " + origen)); return; }
                    if (req.ContentType == null || !req.ContentType.StartsWith("application/json")) { await Enviar(res, 415, Error("El cuerpo debe ser JSON.")); return; }
                    string cuerpo; using (var sr = new StreamReader(req.InputStream, Encoding.UTF8)) cuerpo = await sr.ReadToEndAsync();
                    var pedido = JsonSerializer.Deserialize<PedidoMetrado>(cuerpo);
                    await Enviar(res, 200, await puente.EjecutarAsync(app => Escritura.Registrar(app, cfg, pedido), cfg.esperaRevitMs));
                }
                else await Enviar(res, 404, Error("No existe " + req.HttpMethod + " " + req.Url.AbsolutePath));
            }
            catch (RevitOcupadoException ex) { await Enviar(res, 503, Error(ex.Message)); }
            catch (HttpListenerException) { }
            catch (Exception ex)
            {
                Registro.Escribir("Error " + req.Url + ": " + ex);
                try { await Enviar(res, 500, Error(ex.Message)); } catch { }
            }
        }

        static object Error(string m) => new { ok = false, error = m };
        static Task Enviar(HttpListenerResponse r, int c, object o) => EnviarBytes(r, c, "application/json; charset=utf-8", Encoding.UTF8.GetBytes(JsonSerializer.Serialize(o, Json)));
        static async Task EnviarBytes(HttpListenerResponse r, int c, string tipo, byte[] b)
        {
            r.StatusCode = c; r.ContentType = tipo; r.Headers["Cache-Control"] = "no-store"; r.ContentLength64 = b.Length;
            await r.OutputStream.WriteAsync(b, 0, b.Length); r.OutputStream.Close();
        }
    }
}
