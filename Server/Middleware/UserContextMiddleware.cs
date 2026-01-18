using IMPLANPROD.Server.Services;
using Microsoft.AspNetCore.Http;
using System.Threading.Tasks;

namespace IMPLANPROD.Server.Middleware
{
    /// <summary>
    /// Middleware para verificar y establecer el contexto del usuario en cada solicitud HTTP
    /// </summary>
    /// <remarks>
    /// Este middleware se ejecuta en cada solicitud y almacena información del usuario
    /// en el contexto HTTP para que esté disponible en toda la aplicación
    /// </remarks>
    public class UserContextMiddleware
    {
        private readonly RequestDelegate _next;

        /// <summary>
        /// Constructor del middleware de contexto de usuario
        /// </summary>
        /// <param name="next">Siguiente middleware en la pipeline</param>
        public UserContextMiddleware(RequestDelegate next)
        {
            _next = next;
        }

        /// <summary>
        /// Método que se ejecuta en cada solicitud HTTP
        /// </summary>
        /// <param name="context">Contexto HTTP actual</param>
        /// <param name="usuarioFlexoftService">Servicio para obtener información del usuario Flexoft</param>
        /// <param name="codiempresa">Servicio para obtener el código de empresa</param>
        /// <returns>Tarea asíncrona</returns>
        public async Task InvokeAsync(
            HttpContext context, 
            IUsuarioFlexoftService usuarioFlexoftService, 
            ICodiempresa codiempresa)
        {
            // Solo verificar en rutas de API para optimizar rendimiento
            if (context.Request.Path.StartsWithSegments("/api"))
            {
                var user = context.User;
                
                // Verificar si el usuario está autenticado
                if (user?.Identity?.IsAuthenticated == true)
                {
                    // Obtener y almacenar en el contexto HTTP para uso posterior
                    var idFlexoft = usuarioFlexoftService.InterObtenerIdFlexoft(user);
                    var codigoEmpresa = codiempresa.IObtenerCodiempresa(user);
                    
                    if (idFlexoft.HasValue && codigoEmpresa.HasValue)
                    {
                        // Guardar en los Items del contexto para acceso en controladores
                        context.Items["UsuarioId"] = idFlexoft.Value;
                        context.Items["CodigoEmpresa"] = codigoEmpresa.Value;
                    }
                }
            }

            // Continuar con el siguiente middleware en la pipeline
            await _next(context);
        }
    }
}
// TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Este archivo es parte de la implementación del servicio centralizado 
// para acceder al usuario actual y su código de empresa desde cualquier parte de la aplicación.
// Buscar esta etiqueta para encontrar todos los archivos relacionados con esta funcionalidad.
