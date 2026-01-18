using IMPLANPROD.Client.Auth;
using IMPLANPROD.Client.Repositories;
using IMPLANPROD.Shared.DTOs;
using Microsoft.JSInterop;

namespace IMPLANPROD.Client.Services
{
    /// <summary>
    /// Servicio cliente para verificación de derechos de usuario
    /// Replica la funcionalidad DERACCE del VB6
    /// </summary>
    public class DerechosUsuarioClientService
    {
        private readonly IRepository _repository;
        private readonly IJSRuntime _jsRuntime;

        public DerechosUsuarioClientService(
            IRepository repository,
            IJSRuntime jsRuntime)
        {
            _repository = repository;
            _jsRuntime = jsRuntime;
        }

        /// <summary>
        /// Verifica los derechos del usuario activo sobre una aplicación
        /// Replica la funcionalidad DERACCE del VB6
        /// </summary>
        /// <param name="codApli">Código de aplicación (ej: "APROPEDI")</param>
        /// <param name="desApli">Descripción de la aplicación (ej: "APROBAR PEDIDOS PENDIENTES")</param>
        /// <param name="mostrarMensaje">Si es false, no muestra mensajes de error (equivalente a "NOMESS" en VB6)</param>
        /// <returns>
        /// 0 = Sin acceso
        /// 1 = Solo lectura
        /// 2 = Lectura y escritura
        /// </returns>
        public async Task<int> VerificarDerechoAsync(string codApli, string desApli, bool mostrarMensaje = true)
        {
            try
            {
                // Obtener el número de usuario activo desde UserInfo
                var userInfo = await _repository.GetUserInfoAsync<Auth.UserInfo>();
                
                // Logging detallado para diagnosticar el problema
                Console.WriteLine("=== DEBUG: Verificación de Derechos ===");
                Console.WriteLine($"UserInfo obtenido: {userInfo != null}");
                
                if (userInfo != null)
                {
                    Console.WriteLine($"IdFlexoft: {userInfo.IdFlexoft}");
                    Console.WriteLine($"Id: {userInfo.Id}");
                    Console.WriteLine($"NombreUsuario: {userInfo.NombreUsuario}");
                    Console.WriteLine($"Email: {userInfo.Email}");
                    Console.WriteLine($"Codiempresa: {userInfo.Codiempresa}");
                }
                
                if (userInfo == null || !userInfo.IdFlexoft.HasValue || userInfo.IdFlexoft.Value <= 0)
                {
                    Console.WriteLine($"❌ ERROR: No se pudo obtener el número de usuario activo");
                    Console.WriteLine($"   - userInfo es null: {userInfo == null}");
                    if (userInfo != null)
                    {
                        Console.WriteLine($"   - IdFlexoft.HasValue: {userInfo.IdFlexoft.HasValue}");
                        Console.WriteLine($"   - IdFlexoft.Value: {userInfo.IdFlexoft}");
                    }
                    
                    if (mostrarMensaje)
                    {
                        await MostrarMensajeSinSesion();
                    }
                    return 0;
                }

                var numUser = (short)userInfo.IdFlexoft.Value;

                // Llamar al endpoint de verificación
                var url = $"api/DerechosUsuario/verificar/{numUser}/{codApli}/{desApli}";
                var responseHttp = await _repository.Get<int>(url);

                if (responseHttp.Error)
                {
                    Console.WriteLine($"Error al verificar derechos: {responseHttp.HttpResponseMessage?.StatusCode}");
                    return 0;
                }

                var nivelAcceso = responseHttp.Response;

                // Mostrar mensajes según el nivel de acceso
                if (mostrarMensaje)
                {
                    if (nivelAcceso == 0)
                    {
                        await MostrarMensajeSinAcceso(codApli, desApli);
                    }
                    else if (nivelAcceso == 1)
                    {
                        await MostrarMensajeSoloLectura(codApli, desApli);
                    }
                }

                return nivelAcceso;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al verificar derechos: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Método simplificado para verificar si el usuario tiene acceso de escritura
        /// Retorna true solo si tiene nivel 2 (lectura y escritura)
        /// </summary>
        public async Task<bool> TieneAccesoEscrituraAsync(string codApli, string desApli, bool mostrarMensaje = true)
        {
            var nivel = await VerificarDerechoAsync(codApli, desApli, mostrarMensaje);
            return nivel >= 2;
        }

        /// <summary>
        /// Método simplificado para verificar si el usuario tiene al menos acceso de lectura
        /// Retorna true si tiene nivel 1 o 2
        /// </summary>
        public async Task<bool> TieneAccesoLecturaAsync(string codApli, string desApli, bool mostrarMensaje = true)
        {
            var nivel = await VerificarDerechoAsync(codApli, desApli, mostrarMensaje);
            return nivel >= 1;
        }

        /// <summary>
        /// Muestra mensaje cuando el usuario no tiene acceso
        /// Replica el mensaje del VB6
        /// </summary>
        private async Task MostrarMensajeSinAcceso(string codApli, string desApli)
        {
            var mensaje = $@"
🔒 USUARIO NO HABILITADO

Usuario no Habilitado para Ejecutar esta Opción.

Módulo: {codApli.ToUpper()}
Aplicación: {desApli}

Solicite la Autorización Correspondiente al Supervisor del Sistema.
";

            await _jsRuntime.InvokeVoidAsync("alert", mensaje);
        }

        /// <summary>
        /// Muestra mensaje cuando el usuario solo tiene acceso de lectura
        /// </summary>
        private async Task MostrarMensajeSoloLectura(string codApli, string desApli)
        {
            var mensaje = $@"
⚠️ ACCESO DE SOLO LECTURA

Usted solo tiene permisos de LECTURA para esta opción.
No puede realizar modificaciones.

Módulo: {codApli.ToUpper()}
Aplicación: {desApli}

Si necesita permisos de escritura, solicite la autorización al Supervisor del Sistema.
";

            await _jsRuntime.InvokeVoidAsync("alert", mensaje);
        }

        /// <summary>
        /// Muestra mensaje cuando no hay sesión activa
        /// </summary>
        private async Task MostrarMensajeSinSesion()
        {
            var mensaje = @"
🔐 SESIÓN REQUERIDA

No se pudo verificar su sesión de usuario.
Por favor, inicie sesión nuevamente.
";

            await _jsRuntime.InvokeVoidAsync("alert", mensaje);
        }
    }
}
