using System.Net;
using System.Net.Http;

namespace IMPLANPROD.Client.Repositories
{
    public class HttpResponseWrapper<T>
    {
        public HttpResponseWrapper(T? response, bool error, HttpResponseMessage httpResponseMessage)
        {
            Error = error;
            Response = response;
            HttpResponseMessage = httpResponseMessage ?? throw new ArgumentNullException(nameof(httpResponseMessage));
        }

        public bool Error { get; set; }
        public T? Response { get; set; }
        public HttpResponseMessage HttpResponseMessage { get; }

        public async Task<string> GetErrorMessageAsync()
        {
            if (!Error)
            {
                return string.Empty;
            }

            var statusCode = HttpResponseMessage.StatusCode;
            
            if (statusCode == HttpStatusCode.NotFound)
            {
                return "Recurso no encontrado";
            }
            else if (statusCode == HttpStatusCode.BadRequest)
            {
                return await HttpResponseMessage.Content.ReadAsStringAsync();
            }
            else if (statusCode == HttpStatusCode.Unauthorized)
            {
                // Para errores 401, intentar obtener el mensaje detallado del servidor
                try
                {
                    var serverMessage = await HttpResponseMessage.Content.ReadAsStringAsync();
                    if (!string.IsNullOrWhiteSpace(serverMessage) && serverMessage.Contains("OPERACIÓN BLOQUEADA"))
                    {
                        // Es un mensaje del AuditInterceptor, mostrarlo completo
                        return serverMessage;
                    }
                }
                catch
                {
                    // Si falla la lectura, usar mensaje por defecto
                }
                
                return "Su sesión ha expirado o no ha iniciado sesión.\n\n" +
                       "Por favor, inicie sesión nuevamente para continuar.\n\n" +
                       "Si el problema persiste, contacte al administrador del sistema.";
            }
            else if (statusCode == HttpStatusCode.Forbidden)
            {
                return "No tiene permisos para realizar esta operación";
            }
            else if (statusCode == HttpStatusCode.InternalServerError)
            {
                // Para errores 500, intentar obtener el mensaje detallado del servidor
                try
                {
                    var serverMessage = await HttpResponseMessage.Content.ReadAsStringAsync();
                    if (!string.IsNullOrWhiteSpace(serverMessage))
                    {
                        // Si el mensaje contiene información de auditoría, mostrarlo
                        if (serverMessage.Contains("OPERACIÓN BLOQUEADA") || 
                            serverMessage.Contains("SEGURIDAD DE AUDITORÍA") ||
                            serverMessage.Contains("Usuario no logueado"))
                        {
                            return serverMessage;
                        }
                        
                        // Para otros errores 500, mostrar el mensaje del servidor
                        return serverMessage;
                    }
                }
                catch
                {
                    // Si falla la lectura, usar mensaje por defecto
                }
                
                return "Ha ocurrido un error interno del servidor";
            }

            // Para otros códigos de estado, intentar leer el mensaje del servidor
            try
            {
                var serverMessage = await HttpResponseMessage.Content.ReadAsStringAsync();
                if (!string.IsNullOrWhiteSpace(serverMessage))
                {
                    return serverMessage;
                }
            }
            catch
            {
                // Si falla la lectura, usar mensaje por defecto
            }

            return "Ha ocurrido un error inesperado";
        }

        /// <summary>
        /// Verifica si el error es de tipo Unauthorized (401)
        /// </summary>
        public bool IsUnauthorized => Error && HttpResponseMessage.StatusCode == HttpStatusCode.Unauthorized;

        /// <summary>
        /// Obtiene un mensaje de error específico para problemas de autenticación
        /// con instrucciones claras para el usuario
        /// </summary>
        public string GetAuthenticationErrorMessage()
        {
            if (IsUnauthorized)
            {
                return "🔐 SESIÓN REQUERIDA\n\n" +
                       "Su sesión ha expirado o no ha iniciado sesión en el sistema.\n\n" +
                       "📋 PASOS A SEGUIR:\n" +
                       "1. Haga clic en 'Aceptar' para continuar\n" +
                       "2. Será redirigido a la página de inicio de sesión\n" +
                       "3. Ingrese sus credenciales para acceder nuevamente\n\n" +
                       "💡 CONSEJO: Para evitar que esto suceda, mantenga actividad en el sistema.\n\n" +
                       "Si tiene problemas para iniciar sesión, contacte al administrador del sistema.";
            }
            
            return string.Empty;
        }
    }
}
