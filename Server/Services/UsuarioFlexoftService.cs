using IMPLANPROD.Server.Data;
using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{

    public class UsuarioFlexoftService : IUsuarioFlexoftService
    {
        private readonly DataContext _context;


        public UsuarioFlexoftService(DataContext context)
        {
            _context = context;
        }

        public int? InterObtenerIdFlexoft(ClaimsPrincipal user)
        {
            if (user?.Identity?.IsAuthenticated ?? false)
            {
                var username = user.Identity?.Name;
                var email = user.FindFirst(ClaimTypes.Email)?.Value;

                if (!string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(email))
                {
                    var usuario = _context.Usuarios.FirstOrDefault(u =>
                        u.NombreUsuario == username && u.Email == email);

                    return usuario?.IdFlexoft;
                }
            }

            return null;
        }
    }

}

