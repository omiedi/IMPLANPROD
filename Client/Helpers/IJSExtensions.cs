
using Microsoft.JSInterop;

namespace IMPLANPROD.Client.Helpers
{

    namespace MantenimientoImp.Client.Helpers
    {
        public static class IJSExtensions
        {
            public static Task GuardarComo(this IJSRuntime js, string nombreArchivo, byte[] archivo)
            {
                return js.InvokeAsync<object>("saveAsFile",
                    nombreArchivo,
                    Convert.ToBase64String(archivo)).AsTask();
            }
        }
    }
}


