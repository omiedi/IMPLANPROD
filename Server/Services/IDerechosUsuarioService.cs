using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Interfaz para el servicio de gestión de derechos de usuarios
    /// </summary>
    public interface IDerechosUsuarioService
    {
        /// <summary>
        /// Obtiene todos los derechos de un usuario específico
        /// Replica la lógica del código VB6 con LEFT JOIN
        /// </summary>
        /// <param name="numUser">Número de usuario</param>
        /// <returns>Lista de derechos del usuario sobre todas las aplicaciones</returns>
        Task<List<DerechoUsuarioDTO>> ObtenerDerechosUsuarioAsync(short numUser);

        /// <summary>
        /// Guarda o actualiza los derechos de un usuario
        /// </summary>
        /// <param name="numUser">Número de usuario</param>
        /// <param name="derechos">Lista de derechos a guardar</param>
        /// <returns>True si se guardó correctamente</returns>
        Task<bool> GuardarDerechosUsuarioAsync(short numUser, List<DerechoUsuarioDTO> derechos);

        /// <summary>
        /// Verifica si un usuario tiene derecho de lectura sobre una aplicación
        /// </summary>
        /// <param name="numUser">Número de usuario</param>
        /// <param name="codApli">Código de aplicación</param>
        /// <returns>True si tiene derecho de lectura</returns>
        Task<bool> TieneDerechoLecturaAsync(short numUser, string codApli);

        /// <summary>
        /// Verifica si un usuario tiene derecho de escritura sobre una aplicación
        /// </summary>
        /// <param name="numUser">Número de usuario</param>
        /// <param name="codApli">Código de aplicación</param>
        /// <returns>True si tiene derecho de escritura</returns>
        Task<bool> TieneDerechoEscrituraAsync(short numUser, string codApli);

        /// <summary>
        /// Verifica los derechos de un usuario sobre una aplicación (replica funcionalidad DERACCE del VB6)
        /// Retorna:
        /// 0 = Sin acceso (crea el registro en INDI_APLI si no existe)
        /// 1 = Solo lectura
        /// 2 = Lectura y escritura (o solo escritura)
        /// </summary>
        /// <param name="numUser">Número de usuario</param>
        /// <param name="codApli">Código de aplicación (máximo 32 caracteres)</param>
        /// <param name="desApli">Descripción de la aplicación (máximo 128 caracteres)</param>
        /// <returns>Nivel de acceso del usuario</returns>
        Task<int> VerificarDerechoAsync(short numUser, string codApli, string desApli);
    }
}
