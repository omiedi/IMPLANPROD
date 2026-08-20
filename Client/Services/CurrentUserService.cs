using Microsoft.AspNetCore.Components.Authorization;
using System.Security.Claims;

namespace IMPLANPROD.Client.Services
{
    public class CurrentUserService
    {
        private readonly AuthenticationStateProvider _authProvider;

        public CurrentUserService(AuthenticationStateProvider authProvider)
        {
            _authProvider = authProvider;
        }

        public async Task<short> GetIdFlexoftAsync()
        {
            var auth = await _authProvider.GetAuthenticationStateAsync();
            var claimValue = auth.User.FindFirst("IdFlexoft")?.Value;

            if (!short.TryParse(claimValue, out var idFlexoft))
            {
                return 0;
            }

            return idFlexoft;
        }

        public async Task<int> GetUserIdAsync()
        {
            var auth = await _authProvider.GetAuthenticationStateAsync();
            var claimValue = auth.User.FindFirst("UserId")?.Value;

            if (!int.TryParse(claimValue, out var userId))
            {
                return 0;
            }

            return userId;
        }

        public async Task<string> GetNombreUsuarioAsync()
        {
            var auth = await _authProvider.GetAuthenticationStateAsync();
            return auth.User.FindFirst(ClaimTypes.Name)?.Value ?? string.Empty;
        }

        public async Task<int> GetCodiempresaAsync()
        {
            var auth = await _authProvider.GetAuthenticationStateAsync();
            var claimValue = auth.User.FindFirst("Codiempresa")?.Value;

            if (!int.TryParse(claimValue, out var codiempresa))
            {
                return 0;
            }

            return codiempresa;
        }

        /// <summary>
        /// Obtiene el perfil/rol del usuario actual
        /// </summary>
        public async Task<string> GetPerfilAsync()
        {
            var auth = await _authProvider.GetAuthenticationStateAsync();
            return auth.User.FindFirst(ClaimTypes.Role)?.Value ?? string.Empty;
        }

        /// <summary>
        /// Verifica si el usuario actual tiene el perfil SUPERADMIN
        /// </summary>
        public async Task<bool> EsSuperAdminAsync()
        {
            var perfil = await GetPerfilAsync();
            return perfil.Equals("SUPERADMIN", StringComparison.OrdinalIgnoreCase);
        }
    }
}
