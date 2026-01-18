using IMPLANPROD.Client.Repositories;
using Microsoft.AspNetCore.Components;
using IMPLANPROD.Client.Shared.Components;
using CurrieTechnologies.Razor.SweetAlert2;

namespace IMPLANPROD.Client.Services
{
    /// <summary>
    /// Servicio para manejar errores de autenticación de forma centralizada
    /// Proporciona mensajes informativos y redirección automática al login
    /// </summary>
    public class AuthErrorHandlerService
    {
        private readonly NavigationManager _navigationManager;
        private readonly SweetAlertService _sweetAlertService;

        public AuthErrorHandlerService(NavigationManager navigationManager, SweetAlertService sweetAlertService)
        {
            _navigationManager = navigationManager;
            _sweetAlertService = sweetAlertService;
        }

        /// <summary>
        /// Maneja errores HTTP y muestra mensajes informativos específicos para errores de autenticación
        /// </summary>
        /// <typeparam name="T">Tipo de respuesta del HttpResponseWrapper</typeparam>
        /// <param name="response">Respuesta HTTP a evaluar</param>
        /// <param name="redirectToLogin">Si debe redirigir automáticamente al login en caso de error 401</param>
        /// <returns>True si se manejó un error de autenticación, False en caso contrario</returns>
        public async Task<bool> HandleAuthenticationErrorAsync<T>(HttpResponseWrapper<T> response, bool redirectToLogin = true)
        {
            if (response.IsUnauthorized)
            {
                // Mostrar mensaje informativo específico para errores de autenticación
                var authMessage = response.GetAuthenticationErrorMessage();
                
                await _sweetAlertService.FireAsync(new SweetAlertOptions
                {
                    Title = "Sesión Requerida",
                    Text = authMessage,
                    Icon = SweetAlertIcon.Warning,
                    ConfirmButtonText = "Ir al Login",
                    ConfirmButtonColor = "#007bff",
                    ShowCancelButton = !redirectToLogin,
                    CancelButtonText = redirectToLogin ? null : "Cancelar",
                    AllowOutsideClick = false,
                    AllowEscapeKey = false
                });

                if (redirectToLogin)
                {
                    // Limpiar datos de sesión y redirigir al login
                    await ClearSessionAndRedirectToLogin();
                }

                return true; // Se manejó un error de autenticación
            }

            return false; // No era un error de autenticación
        }

        /// <summary>
        /// Maneja errores HTTP con mensaje genérico si no es de autenticación
        /// </summary>
        /// <typeparam name="T">Tipo de respuesta del HttpResponseWrapper</typeparam>
        /// <param name="response">Respuesta HTTP a evaluar</param>
        /// <param name="customErrorTitle">Título personalizado para errores no relacionados con autenticación</param>
        /// <param name="redirectToLogin">Si debe redirigir automáticamente al login en caso de error 401</param>
        /// <returns>True si se manejó algún error, False si no hubo errores</returns>
        public async Task<bool> HandleErrorAsync<T>(HttpResponseWrapper<T> response, string customErrorTitle = "Error", bool redirectToLogin = true)
        {
            if (!response.Error)
            {
                return false; // No hay error
            }

            // Primero intentar manejar como error de autenticación
            if (await HandleAuthenticationErrorAsync(response, redirectToLogin))
            {
                return true;
            }

            // Si no es error de autenticación, mostrar mensaje genérico
            var errorMessage = await response.GetErrorMessageAsync();
            await _sweetAlertService.FireAsync(customErrorTitle, errorMessage, SweetAlertIcon.Error);

            return true; // Se manejó un error genérico
        }

        /// <summary>
        /// Muestra un mensaje informativo cuando se detecta que el usuario necesita iniciar sesión
        /// </summary>
        public async Task ShowSessionRequiredMessageAsync()
        {
            await _sweetAlertService.FireAsync(new SweetAlertOptions
            {
                Title = "🔐 Sesión Requerida",
                Text = "Para acceder a esta funcionalidad debe iniciar sesión en el sistema.\n\n" +
                       "📋 PASOS A SEGUIR:\n" +
                       "1. Haga clic en 'Ir al Login'\n" +
                       "2. Ingrese sus credenciales\n" +
                       "3. Regrese a esta página una vez autenticado\n\n" +
                       "💡 Si ya tiene una sesión activa, es posible que haya expirado.",
                Icon = SweetAlertIcon.Info,
                ConfirmButtonText = "Ir al Login",
                ConfirmButtonColor = "#007bff",
                ShowCancelButton = true,
                CancelButtonText = "Cancelar"
            });

            await ClearSessionAndRedirectToLogin();
        }

        /// <summary>
        /// Limpia los datos de sesión y redirige al login
        /// </summary>
        private async Task ClearSessionAndRedirectToLogin()
        {
            try
            {
                // Aquí se podría agregar lógica para limpiar el localStorage si es necesario
                // Por ejemplo, usando IJSRuntime para limpiar tokens expirados
                
                // Redirigir al login
                _navigationManager.NavigateTo("/login", true);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error al limpiar sesión y redirigir: {ex.Message}");
                // Forzar redirección incluso si hay error
                _navigationManager.NavigateTo("/login", true);
            }
        }

        /// <summary>
        /// Verifica si una respuesta HTTP indica un problema de autenticación
        /// </summary>
        /// <typeparam name="T">Tipo de respuesta del HttpResponseWrapper</typeparam>
        /// <param name="response">Respuesta HTTP a evaluar</param>
        /// <returns>True si es un error de autenticación, False en caso contrario</returns>
        public static bool IsAuthenticationError<T>(HttpResponseWrapper<T> response)
        {
            return response.IsUnauthorized;
        }
    }
}
