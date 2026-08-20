using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Mvc;
using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{
    public interface IMoviBrecepenService
    {
        Task<ActionResult> RegistrarMovimientoAsync(MoviStockDTO movimiento, ClaimsPrincipal user);
    }
}
