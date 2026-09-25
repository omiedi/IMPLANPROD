using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para la tabla EXTRAS_REMITO.
    /// Almacena datos adicionales asociados a un remito de cliente.
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    public class ExtrasRemitoController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly IUserContextService _userContextService;
        private readonly ILogger<ExtrasRemitoController> _logger;

        public ExtrasRemitoController(
            DataContext context,
            IUserContextService userContextService,
            ILogger<ExtrasRemitoController> logger)
        {
            _context = context;
            _userContextService = userContextService;
            _logger = logger;
        }

        /// <summary>
        /// Guarda o actualiza los datos adicionales de un remito en EXTRAS_REMITO.
        /// Si ya existe un registro para el mismo DOPRILAR, se actualiza.
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<int>> GuardarAsync([FromBody] ExtrasRemito extrasRemito)
        {
            try
            {
                if (extrasRemito == null)
                {
                    return BadRequest("Los datos de ExtrasRemito son inválidos.");
                }

                if (string.IsNullOrWhiteSpace(extrasRemito.Doprilar))
                {
                    return BadRequest("El número de remito (Doprilar) es obligatorio.");
                }

                if (!extrasRemito.NumeClie.HasValue || extrasRemito.NumeClie.Value <= 0)
                {
                    return BadRequest("El número de cliente es obligatorio.");
                }

                var ahora = DateTime.Now;
                var userId = _userContextService.GetCurrentUserId();

                // Normalizar longitudes
                var doprilar = extrasRemito.Doprilar.Trim();
                if (doprilar.Length > 24)
                {
                    doprilar = doprilar.Substring(0, 24);
                }

                var observaciones = (extrasRemito.Observaciones ?? string.Empty).Trim();
                if (observaciones.Length > 256)
                {
                    observaciones = observaciones.Substring(0, 256);
                }

                // Verificar si ya existe un registro para este remito
                var existente = await _context.ExtrasRemitos
                    .FirstOrDefaultAsync(x => x.Doprilar == doprilar);

                ExtrasRemito entidad;
                if (existente != null)
                {
                    existente.NumeClie = extrasRemito.NumeClie;
                    existente.Fechmov = extrasRemito.Fechmov;
                    existente.Observaciones = observaciones;
                    existente.Bultos = extrasRemito.Bultos;
                    existente.ValorDeclarado = extrasRemito.ValorDeclarado;
                    existente.LastUpdate = ahora;
                    existente.NumUsuar = userId;

                    entidad = existente;
                    _logger.LogInformation("Actualizando ExtrasRemito para remito {Doprilar}", doprilar);
                }
                else
                {
                    entidad = new ExtrasRemito
                    {
                        NumeClie = extrasRemito.NumeClie,
                        Fechmov = extrasRemito.Fechmov,
                        Doprilar = doprilar,
                        Observaciones = observaciones,
                        Bultos = extrasRemito.Bultos,
                        ValorDeclarado = extrasRemito.ValorDeclarado,
                        AddRecord = ahora,
                        LastUpdate = ahora,
                        CodiEmprNet = 0,
                        NumUsuar = userId
                    };

                    _context.ExtrasRemitos.Add(entidad);
                    _logger.LogInformation("Creando ExtrasRemito para remito {Doprilar}", doprilar);
                }

                await _context.SaveChangesAsync();
                return Ok(entidad.Id);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al guardar ExtrasRemito");
                return StatusCode(500, $"Error al guardar datos adicionales del remito: {ex.Message}");
            }
        }
    }
}
