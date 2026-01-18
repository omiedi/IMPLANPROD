using Microsoft.AspNetCore.Builder;

namespace IMPLANPROD.Server.Middleware
{
    /// <summary>
    /// Extensiones para registrar los middleware personalizados en la pipeline de la aplicación
    /// </summary>
    public static class MiddlewareExtensions
    {
        /// <summary>
        /// Agrega el middleware de contexto de usuario a la pipeline de la aplicación
        /// </summary>
        /// <param name="builder">IApplicationBuilder</param>
        /// <returns>IApplicationBuilder con el middleware configurado</returns>
        public static IApplicationBuilder UseUserContext(this IApplicationBuilder builder)
        {
            return builder.UseMiddleware<UserContextMiddleware>();
        }
    }
}
// TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Este archivo es parte de la implementación del servicio centralizado 
// para acceder al usuario actual y su código de empresa desde cualquier parte de la aplicación.
// Buscar esta etiqueta para encontrar todos los archivos relacionados con esta funcionalidad.
