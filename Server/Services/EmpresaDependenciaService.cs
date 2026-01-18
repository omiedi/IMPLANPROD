using IMPLANPROD.Server.Data;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{
    /// <summary>
    /// Servicio para validaciones específicas por empresa y dependencia
    /// </summary>
    public class EmpresaDependenciaService : IEmpresaDependenciaService
    {
        private readonly DataContext _context;
        private readonly IHttpContextAccessor _httpContextAccessor;

        public EmpresaDependenciaService(DataContext context, IHttpContextAccessor httpContextAccessor)
        {
            _context = context;
            _httpContextAccessor = httpContextAccessor;
        }

        /// <summary>
        /// Obtiene el usuario actual desde el contexto HTTP
        /// </summary>
        private ClaimsPrincipal? ObtenerUsuarioActual()
        {
            return _httpContextAccessor.HttpContext?.User;
        }

        /// <summary>
        /// Obtiene la información completa del usuario actual desde la base de datos
        /// </summary>
        private Shared.Entities.Usuario? ObtenerUsuarioCompleto()
        {
            var user = ObtenerUsuarioActual();
            if (user?.Identity?.IsAuthenticated ?? false)
            {
                var username = user.Identity?.Name;
                var email = user.FindFirst(ClaimTypes.Email)?.Value;

                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(email))
                {
                    return _context.Usuarios
                        .Include(u => u.Empresa)
                        .Include(u => u.Dependencia)
                        .FirstOrDefault(u => u.NombreUsuario == username && u.Email == email);
                }
            }
            return null;
        }

        /// <summary>
        /// Obtiene el nombre de la empresa del usuario actual
        /// </summary>
        public string? ObtenerNombreEmpresaActual()
        {
            var usuario = ObtenerUsuarioCompleto();
            if (usuario == null) return null;
            
            // Si la navegación no funciona, buscar directamente por ID
            if (usuario.Empresa == null && usuario.EmpresaId > 0)
            {
                var empresa = _context.Empresas.FirstOrDefault(e => e.Id == usuario.EmpresaId);
                return empresa?.razon;
            }
            
            return usuario?.Empresa?.razon;
        }

        /// <summary>
        /// Obtiene el nombre de la dependencia del usuario actual
        /// </summary>
        public string? ObtenerNombreDependenciaActual()
        {
            var usuario = ObtenerUsuarioCompleto();
            if (usuario == null) return null;
            
            // Si la navegación no funciona, buscar directamente por ID
            if (usuario.Dependencia == null && usuario.DependenciaId > 0)
            {
                var dependencia = _context.Dependencias.FirstOrDefault(d => d.Id == usuario.DependenciaId);
                return dependencia?.dependencia;
            }
            
            return usuario?.Dependencia?.dependencia;
        }

        /// <summary>
        /// Verifica si el usuario pertenece a una empresa específica
        /// </summary>
        public bool EsEmpresa(string nombreEmpresa)
        {
            var empresaActual = ObtenerNombreEmpresaActual();
            if (string.IsNullOrEmpty(empresaActual))
                return false;

            return empresaActual.Equals(nombreEmpresa, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Verifica si el usuario pertenece a una dependencia específica
        /// </summary>
        public bool EsDependencia(string nombreDependencia)
        {
            var dependenciaActual = ObtenerNombreDependenciaActual();
            if (string.IsNullOrEmpty(dependenciaActual))
                return false;

            return dependenciaActual.Equals(nombreDependencia, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Verifica si el usuario pertenece a una empresa Y dependencia específicas
        /// </summary>
        public bool EsEmpresaYDependencia(string nombreEmpresa, string nombreDependencia)
        {
            return EsEmpresa(nombreEmpresa) && EsDependencia(nombreDependencia);
        }

        /// <summary>
        /// Obtiene el ID de la empresa del usuario actual
        /// </summary>
        public int? ObtenerEmpresaIdActual()
        {
            var usuario = ObtenerUsuarioCompleto();
            return usuario?.EmpresaId;
        }

        /// <summary>
        /// Obtiene el ID de la dependencia del usuario actual
        /// </summary>
        public int? ObtenerDependenciaIdActual()
        {
            var usuario = ObtenerUsuarioCompleto();
            return usuario?.DependenciaId;
        }
    }
}
