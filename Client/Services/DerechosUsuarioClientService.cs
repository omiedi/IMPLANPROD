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
        private readonly CurrentUserService _currentUserService;

        public DerechosUsuarioClientService(
            IRepository repository,
            IJSRuntime jsRuntime,
            CurrentUserService currentUserService)
        {
            _repository = repository;
            _jsRuntime = jsRuntime;
            _currentUserService = currentUserService;
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
                // Si el usuario es SUPERADMIN, tiene acceso total a todo
                var esSuperAdmin = await _currentUserService.EsSuperAdminAsync();
                if (esSuperAdmin)
                {
                    Console.WriteLine($"✅ Usuario SUPERADMIN: Acceso total concedido para {codApli}");
                    return 2; // Nivel 2: Lectura y escritura
                }

                var numUser = await _currentUserService.GetIdFlexoftAsync();

                if (numUser <= 0)
                {
                    Console.WriteLine($"❌ ERROR: No se pudo obtener el número de usuario activo");

                    if (mostrarMensaje)
                    {
                        await MostrarMensajeSinSesion();
                    }
                    return 0;
                }

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
