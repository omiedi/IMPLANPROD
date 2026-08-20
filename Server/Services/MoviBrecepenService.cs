using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Security.Claims;

namespace IMPLANPROD.Server.Services
{
    public class MoviBrecepenService : IMoviBrecepenService
    {
        private readonly DataContext _context;
        private readonly IUsuarioFlexoftService _usuarioFlexoftService;
        private readonly ICodiempresa _codiempresa;
        private readonly ILogger<MoviBrecepenService> _logger;

        public MoviBrecepenService(
            DataContext context,
            IUsuarioFlexoftService usuarioFlexoftService,
            ICodiempresa codiempresa,
            ILogger<MoviBrecepenService> logger)
        {
            _context = context;
            _usuarioFlexoftService = usuarioFlexoftService;
            _codiempresa = codiempresa;
            _logger = logger;
        }

        public async Task<ActionResult> RegistrarMovimientoAsync(MoviStockDTO movimiento, ClaimsPrincipal user)
        {
            try
            {
                int? idFlexoft = _usuarioFlexoftService.InterObtenerIdFlexoft(user);
                short? codigoEmpresa = _codiempresa.IObtenerCodiempresa(user);

                if (!idFlexoft.HasValue)
                {
                    _logger.LogWarning("Intento de registro BRECEPEN sin usuario válido");
                    return new BadRequestObjectResult("Sesión caducada - Debe cerrar sesión y loguearse nuevamente.");
                }

                var ahora = DateTime.Now;
                int claveHp = (ahora.Hour * 100) + ahora.Minute;

                short? monedaPesos = await ObtenerCodigoMonedaPesosAsync();

                var numeroCliente = movimiento.NumeroCliente ?? 0;
                var numeClie = 0;
                var numeProv = 0;
                if (numeroCliente < 0)
                {
                    numeProv = Math.Abs(numeroCliente);
                }
                else
                {
                    numeClie = numeroCliente;
                }

                var entidad = new Brecepen
                {
                    Fechmov = movimiento.Fecha.Date,
                    Feremovi = ahora,
                    COD_INTE = movimiento.CodigoInterno,
                    Idenlote = string.IsNullOrWhiteSpace(movimiento.Lote) ? null : movimiento.Lote,
                    Codimovi = movimiento.CodigoMovimiento,
                    Doprilar = movimiento.DocumentoPrimarioLargo,
                    Dopricor = movimiento.DocumentoPrimarioLargo,
                    Doseclar = movimiento.DocumentoSecundarioLargo,
                    Doseccor = movimiento.DocumentoSecundarioLargo,
                    Refercor = movimiento.Referencia,
                    Valount = movimiento.ValorUnitario ?? 0,
                    Moneval = monedaPesos,
                    Depomovi = (short?)movimiento.Deposito,
                    NumeClie = numeClie,
                    NumeProv = numeProv,
                    Cantingre = movimiento.Cantidad,
                    Cantsalid = 0,
                    ClaveHp = claveHp,
                    AddRecord = ahora,
                    LastUpdate = ahora,
                    CodiEmprNet = codigoEmpresa,
                    NumUsuar = idFlexoft
                };

                _context.Brecepens.Add(entidad);
                await _context.SaveChangesAsync();

                return new OkObjectResult(new { Id = entidad.Id, Mensaje = "Movimiento registrado en BRECEPEN" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar movimiento en BRECEPEN: {Message}", ex.Message);
                return new BadRequestObjectResult($"Error al registrar movimiento en BRECEPEN: {ex.Message}");
            }
        }

        private async Task<short?> ObtenerCodigoMonedaPesosAsync()
        {
            var moneda = await _context.Tamoneds
                .AsNoTracking()
                .Where(m => m.Status >= 0 && m.Descripcion != null)
                .OrderBy(m => m.CodigoMoneda)
                .FirstOrDefaultAsync(m => m.Descripcion.ToUpper() == "PESOS" || m.Descripcion.ToUpper().Contains("PESO"));

            return moneda != null ? (short?)moneda.CodigoMoneda : (short?)0;
        }
    }
}
