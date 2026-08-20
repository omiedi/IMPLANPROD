using IMPLANPROD.Server.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Data.SqlClient;
using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{

    public class UsuarioFlexoftService : IUsuarioFlexoftService
    {
        private readonly DataContext _context;
        private readonly IMemoryCache _cache;
        private readonly ILogger<UsuarioFlexoftService> _logger;


        public UsuarioFlexoftService(
            DataContext context,
            IMemoryCache cache,
            ILogger<UsuarioFlexoftService> logger)
        {
            _context = context;
            _cache = cache;
            _logger = logger;
        }

        public int? InterObtenerIdFlexoft(ClaimsPrincipal user)
        {
            if (user?.Identity?.IsAuthenticated ?? false)
            {
                var username = user.Identity?.Name;
                var email = user.FindFirst(ClaimTypes.Email)?.Value;

                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(email))
                {
                    var cacheKey = $"idflexoft:{username}:{email}";
                    if (_cache.TryGetValue(cacheKey, out int? idFlexoftEnCache))
                    {
                        return idFlexoftEnCache;
                    }

                    try
                    {
                        var idFlexoft = _context.Usuarios
                            .AsNoTracking()
                            .Where(u => u.NombreUsuario == username && u.Email == email)
                            .Select(u => (int?)u.IdFlexoft)
                            .FirstOrDefault();

                        _cache.Set(cacheKey, idFlexoft, TimeSpan.FromMinutes(10));
                        return idFlexoft;
                    }
                    catch (SqlException ex)
                    {
                        _logger.LogWarning(ex,
                            "[InterObtenerIdFlexoft] Timeout/SQL error obteniendo IdFlexoft para usuario {Username}",
                            username);
                        return null;
                    }
                    catch (Exception ex)
                    {
                        _logger.LogWarning(ex,
                            "[InterObtenerIdFlexoft] Error obteniendo IdFlexoft para usuario {Username}",
                            username);
                        return null;
                    }
                }
            }

            return null;
        }
    }

}

