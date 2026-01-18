using Microsoft.AspNetCore.Http;
using System.Security.Claims;
using IMPLANPROD.Server.Data;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Implementación del servicio de contexto de usuario que proporciona acceso centralizado
    /// a la información del usuario actual y su código de empresa
    /// </summary>
    public class UserContextService : IUserContextService
    {
        private readonly IHttpContextAccessor _httpContextAccessor;
        private readonly IUsuarioFlexoftService _usuarioFlexoftService;
        private readonly ICodiempresa _codiempresa;
        private readonly DataContext _context;

        /// <summary>
        /// Constructor del servicio de contexto de usuario
        /// </summary>
        /// <param name="httpContextAccessor">Acceso al contexto HTTP actual</param>
        /// <param name="usuarioFlexoftService">Servicio para obtener información del usuario Flexoft</param>
        /// <param name="codiempresa">Servicio para obtener el código de empresa</param>
        /// <param name="context">Contexto de base de datos</param>
        public UserContextService(
            IHttpContextAccessor httpContextAccessor,
            IUsuarioFlexoftService usuarioFlexoftService,
            ICodiempresa codiempresa,
            DataContext context)
        {
            _httpContextAccessor = httpContextAccessor;
            _usuarioFlexoftService = usuarioFlexoftService;
            _codiempresa = codiempresa;
            _context = context;
        }

        /// <summary>
        /// Obtiene el ID del usuario actual en el sistema Flexoft (IdFlexoft)
        /// </summary>
        /// <returns>IdFlexoft del usuario o null si no está autenticado</returns>
        public int? GetCurrentUserId()
        {
            var user = GetCurrentUser();
            return user != null ? _usuarioFlexoftService.InterObtenerIdFlexoft(user) : null;
        }

        /// <summary>
        /// Obtiene el ID real del usuario actual (usuarios.Id) para obtener datos de empresa/dependencia
        /// </summary>
        /// <returns>usuarios.Id del usuario o null si no está autenticado</returns>
        public int? GetCurrentUsuarioId()
        {
            var user = GetCurrentUser();
            if (user?.Identity?.IsAuthenticated ?? false)
            {
                var username = user.Identity?.Name;
                var email = user.FindFirst(System.Security.Claims.ClaimTypes.Email)?.Value;

                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(email))
                {
                    var usuario = _context.Usuarios.FirstOrDefault(u =>
                        u.NombreUsuario == username && u.Email == email);

                    return usuario?.Id;
                }
            }

            return null;
        }

        /// <summary>
        /// Obtiene el código de empresa del usuario actual
        /// </summary>
        /// <returns>Código de empresa o null si no está autenticado o no tiene empresa asignada</returns>
        public short? GetCurrentCodigoEmpresa()
        {
            var user = GetCurrentUser();
            return user != null ? _codiempresa.IObtenerCodiempresa(user) : null;
        }

        /// <summary>
        /// Verifica si el usuario está autenticado y tiene valores válidos de ID y código de empresa
        /// </summary>
        /// <returns>True si el usuario está autenticado y tiene valores válidos, False en caso contrario</returns>
        public bool IsAuthenticated()
        {
            return GetCurrentUserId().HasValue && GetCurrentCodigoEmpresa().HasValue;
        }

        /// <summary>
        /// Obtiene el objeto ClaimsPrincipal del usuario actual
        /// </summary>
        /// <returns>ClaimsPrincipal del usuario actual o null si no hay contexto HTTP</returns>
        public ClaimsPrincipal? GetCurrentUser()
        {
            return _httpContextAccessor.HttpContext?.User;
        }
    }
}
// TODO: CENTRALIZACIÓN_CONTEXTO_USUARIO - Este archivo es parte de la implementación del servicio centralizado 
// para acceder al usuario actual y su código de empresa desde cualquier parte de la aplicación.
// Buscar esta etiqueta para encontrar todos los archivos relacionados con esta funcionalidad.
