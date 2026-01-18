using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{
    public interface IUsuarioFlexoftService
    {
        int? InterObtenerIdFlexoft(ClaimsPrincipal user);
    }
}
