using System.Security.Claims;
using IMPLANPROD.Client.Repositories;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Authorization;

namespace IMPLANPROD.Client.Services
{
    /// <summary>
    /// Servicio que verifica si el usuario tiene permiso para acceder a una página del RBAC.
    /// - SUPERADMIN y ADMIN pasan siempre (sin consulta).
    /// - USUARIO: se carga una vez la lista de rutas permitidas y se cachea por sesión.
    /// Uso en cada página: await PageAccess.RequiereAccesoAsync("/mi-ruta");
    /// </summary>
    public class PageAccessService
    {
        private readonly AuthenticationStateProvider _authProvider;
        private readonly IRepository _repository;
        private readonly NavigationManager _navigationManager;

        private HashSet<string>? _rutasPermitidas;
        private int _cachedUserId = -1;

        public PageAccessService(
            AuthenticationStateProvider authProvider,
            IRepository repository,
            NavigationManager navigationManager)
        {
            _authProvider = authProvider;
            _repository = repository;
            _navigationManager = navigationManager;
        }

        /// <summary>
        /// Verifica acceso a la ruta indicada.
        /// Si el usuario no tiene permiso redirige a /acceso-denegado.
        /// SUPERADMIN y ADMIN no son verificados (tienen acceso total).
        /// </summary>
        public async Task RequiereAccesoAsync(string ruta)
        {
            var auth = await _authProvider.GetAuthenticationStateAsync();
            var user = auth.User;

            if (!(user.Identity?.IsAuthenticated ?? false))
            {
                _navigationManager.NavigateTo("/login", replace: true);
                return;
            }

            // Roles privilegiados acceden a todo sin consulta
            if (user.IsInRole("SUPERADMIN") || user.IsInRole("ADMINISTRADOR") || user.IsInRole("ADMIN"))
                return;

            var userId = int.Parse(user.FindFirst("UserId")?.Value ?? "0");

            if (_rutasPermitidas == null || _cachedUserId != userId)
                await CargarRutasPermitidasAsync(userId);

            // Prefix match: /cities/edit/5 es accesible si el usuario tiene /cities/edit
            var rutaNorm = ruta.TrimEnd('/').ToLowerInvariant();
            var tieneAcceso = _rutasPermitidas!.Any(r =>
                rutaNorm.StartsWith(r, StringComparison.OrdinalIgnoreCase));

            if (!tieneAcceso)
                _navigationManager.NavigateTo("/acceso-denegado", replace: true);
        }

        /// <summary>
        /// Invalida la caché. Llamar tras logout o cuando se modifican permisos del usuario actual.
        /// </summary>
        public void InvalidarCache()
        {
            _rutasPermitidas = null;
            _cachedUserId = -1;
        }

        private async Task CargarRutasPermitidasAsync(int userId)
        {
            _cachedUserId = userId;
            _rutasPermitidas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            var respIds     = await _repository.Get<List<int>>($"api/seguridad-rutas/usuarios/{userId}/paginas");
            var respModulos = await _repository.Get<List<SeguridadModuloDto>>("api/seguridad-rutas/modulos?incluirConfiguracion=true");

            if (respIds.Error || respModulos.Error) return;

            var allowedIds = new HashSet<int>(respIds.Response ?? new());

            foreach (var modulo in respModulos.Response ?? new())
            {
                foreach (var pagina in modulo.Paginas.Where(p => allowedIds.Contains(p.Id)))
                {
                    _rutasPermitidas.Add(pagina.Ruta.TrimEnd('/').ToLowerInvariant());
                }
            }
        }
    }
}
