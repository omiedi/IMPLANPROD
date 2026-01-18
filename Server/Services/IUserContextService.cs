using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Interfaz para el servicio de contexto del usuario que proporciona acceso centralizado
    /// a la información del usuario actual y su código de empresa
    /// </summary>
    /// <remarks>
    /// Este servicio permite obtener información del usuario autenticado desde cualquier
    /// parte de la aplicación sin necesidad de inyectar múltiples servicios
    /// </remarks>
    public interface IUserContextService
    {
        /// <summary>
        /// Obtiene el ID del usuario actual en el sistema Flexoft (IdFlexoft)
        /// </summary>
        /// <returns>IdFlexoft del usuario o null si no está autenticado</returns>
        int? GetCurrentUserId();

        /// <summary>
        /// Obtiene el ID real del usuario actual (usuarios.Id) para obtener datos de empresa/dependencia
        /// </summary>
        /// <returns>usuarios.Id del usuario o null si no está autenticado</returns>
        int? GetCurrentUsuarioId();
        
        /// <summary>
        /// Obtiene el código de empresa del usuario actual
        /// </summary>
        /// <returns>Código de empresa o null si no está autenticado o no tiene empresa asignada</returns>
        short? GetCurrentCodigoEmpresa();
        
        /// <summary>
        /// Verifica si el usuario está autenticado y tiene valores válidos de ID y código de empresa
        /// </summary>
        /// <returns>True si el usuario está autenticado y tiene valores válidos, False en caso contrario</returns>
        bool IsAuthenticated();

        /// <summary>
        /// Obtiene el objeto ClaimsPrincipal del usuario actual
        /// </summary>
        /// <returns>ClaimsPrincipal del usuario actual o null si no hay contexto HTTP</returns>
        ClaimsPrincipal? GetCurrentUser();
    }
}
// TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Este archivo es parte de la implementación del servicio centralizado 
// para acceder al usuario actual y su código de empresa desde cualquier parte de la aplicación.
// Buscar esta etiqueta para encontrar todos los archivos relacionados con esta funcionalidad.
