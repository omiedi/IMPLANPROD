using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Interfaz para obtener el código de empresa
    /// </summary>
    public interface ICodiempresa
    {
        /// <summary>
        /// Obtiene el código de empresa a partir de las claims del usuario
        /// </summary>
        /// <param name="user">Usuario autenticado</param>
        /// <returns>Código de empresa</returns>
        short? IObtenerCodiempresa(ClaimsPrincipal user);

        /// <summary>
        /// Obtiene el nombre de la empresa (razón social) a partir del ID del usuario
        /// </summary>
        /// <param name="usuarioId">ID del usuario</param>
        /// <returns>Nombre de la empresa (razón social)</returns>
        string? ObtenerNombreEmpresa(int usuarioId);

        /// <summary>
        /// Obtiene el nombre de la dependencia a partir del ID del usuario
        /// </summary>
        /// <param name="usuarioId">ID del usuario</param>
        /// <returns>Nombre de la dependencia</returns>
        string? ObtenerNombreDependencia(int usuarioId);
    }
}
