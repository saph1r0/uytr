// Pasa los pedidos del servidor HTTP al hilo de Revit con un ExternalEvent (la API solo se usa dentro de Execute).
using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using Autodesk.Revit.UI;

namespace NexoObraNO
{
    public class RevitOcupadoException : Exception { public RevitOcupadoException(string m) : base(m) { } }

    public class Puente : IExternalEventHandler
    {
        const int EnCola = 0, Ejecutando = 1, Cancelado = 2;
        class Pedido
        {
            public Func<UIApplication, object> Accion;
            public readonly TaskCompletionSource<object> Resultado = new TaskCompletionSource<object>(TaskCreationOptions.RunContinuationsAsynchronously);
            public int Estado = EnCola;
        }
        readonly ConcurrentQueue<Pedido> cola = new ConcurrentQueue<Pedido>();
        ExternalEvent evento;

        public void Crear() => evento = ExternalEvent.Create(this);
        public void Liberar() { evento?.Dispose(); evento = null; }

        public async Task<object> EjecutarAsync(Func<UIApplication, object> accion, int esperaMs)
        {
            var p = new Pedido { Accion = accion };
            cola.Enqueue(p);
            var r = evento.Raise();
            if (r == ExternalEventRequest.Denied || r == ExternalEventRequest.TimedOut)
            {
                Interlocked.CompareExchange(ref p.Estado, Cancelado, EnCola);
                throw new RevitOcupadoException("Revit no aceptó el pedido (" + r + ").");
            }
            var t = await Task.WhenAny(p.Resultado.Task, Task.Delay(esperaMs));
            if (t != p.Resultado.Task && Interlocked.CompareExchange(ref p.Estado, Cancelado, EnCola) == EnCola)
                throw new RevitOcupadoException("Revit ocupado: cierra cualquier diálogo o termina el comando en curso y vuelve a intentar.");
            return await p.Resultado.Task;
        }

        public void Execute(UIApplication app)
        {
            while (cola.TryDequeue(out Pedido p))
            {
                if (Interlocked.CompareExchange(ref p.Estado, Ejecutando, EnCola) != EnCola) continue;   // cancelado: nunca se ejecuta
                try { p.Resultado.SetResult(p.Accion(app)); }
                catch (Exception ex) { Registro.Escribir("Error en Revit: " + ex); p.Resultado.SetException(ex); }
            }
        }

        public string GetName() => "NEXO OBRA (NO)";
    }
}
