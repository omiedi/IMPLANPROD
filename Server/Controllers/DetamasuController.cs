using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DetamasuController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<DetamasuController> _logger;

        public DetamasuController(DataContext context, ILogger<DetamasuController> logger)
        {
            _context = context;
            _logger = logger;
        }

        [HttpGet("resumen/{talon:int}")]
        public async Task<ActionResult<List<DetamasuResumenDTO>>> GetResumenByTalon([FromRoute] int talon)
        {
            try
            {
                // Nota: algunas versiones de EF Core no traducen correctamente ciertos GroupBy + proyecciones.
                // Para garantizar compatibilidad, traemos los movimientos del talón y agrupamos en memoria.
                var movimientos = await _context.Detamasus
                    .AsNoTracking()
                    .Where(d => d.Talon == talon)
                    .ToListAsync();

                var data = movimientos
                    .GroupBy(d => new { d.Cod_Inte, d.Cod_Exte })
                    .Select(g => new
                    {
                        CodInte = g.Key.Cod_Inte,
                        CodExte = g.Key.Cod_Exte,
                        Br = g.Where(x => (x.CodiMovi ?? string.Empty).Trim().ToUpperInvariant() == "BR")
                            .Sum(x => x.CantIngre ?? 0m),
                        Cf = g.Where(x => (x.CodiMovi ?? string.Empty).Trim().ToUpperInvariant() == "CF")
                            .Sum(x => x.CantSalid ?? 0m),
                        Dv = g.Where(x => (x.CodiMovi ?? string.Empty).Trim().ToUpperInvariant() == "DV")
                            .Sum(x => x.CantSalid ?? 0m),
                        Aj = g.Where(x => (x.CodiMovi ?? string.Empty).Trim().ToUpperInvariant() == "AJ")
                            .Sum(x => (x.CantIngre ?? 0m) - (x.CantSalid ?? 0m)),
                        FechaUltMov = g.Max(x => x.FeRemovi ?? x.FechMov)
                    })
                    .Select(x => new
                    {
                        x.CodInte,
                        x.CodExte,
                        x.Br,
                        x.Cf,
                        x.Dv,
                        x.Aj,
                        Resul = (x.Br - (x.Cf + x.Dv) + x.Aj),
                        x.FechaUltMov
                    })
                    .Where(x => x.Resul >= 0m)
                    .ToList();

                var codintes = data
                    .Where(x => x.CodInte.HasValue)
                    .Select(x => x.CodInte!.Value)
                    .Distinct()
                    .ToList();

                var master = await _context.masters
                    .Where(m => codintes.Contains(m.Codint))
                    .Select(m => new { m.Codint, m.Codigo, m.Descripcion })
                    .ToListAsync();

                var masterDict = master.ToDictionary(x => x.Codint, x => x);

                var result = data
                    .Where(x => x.CodInte.HasValue)
                    .Select(x =>
                    {
                        masterDict.TryGetValue(x.CodInte!.Value, out var m);
                        return new DetamasuResumenDTO
                        {
                            Cod_Inte = x.CodInte!.Value,
                            Cod_Exte = x.CodExte,
                            Codigo = m?.Codigo,
                            Descripcion = m?.Descripcion,
                            Br = Math.Round(x.Br, 4),
                            Cf = Math.Round(x.Cf, 4),
                            Dv = Math.Round(x.Dv, 4),
                            Aj = Math.Round(x.Aj, 4),
                            Resul = Math.Round(x.Resul, 4),
                            Cantidad = Math.Round(x.Resul, 4),
                            FechaUltMov = x.FechaUltMov
                        };
                    })
                    .OrderBy(x => x.Descripcion ?? string.Empty)
                    .ThenBy(x => x.Cod_Inte)
                    .ToList();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener resumen DETAMASU Talon={Talon}", talon);
                return StatusCode(500, $"Error al obtener resumen. Detalle: {ex.Message}");
            }
        }

        [HttpPost("movimiento")]
        public async Task<ActionResult> CrearMovimiento([FromBody] DetamasuMovimientoCreateDTO dto)
        {
            try
            {
                if (dto.Talon.GetValueOrDefault() <= 0)
                {
                    return BadRequest("Talon inválido");
                }

                if (!dto.Cod_Inte.HasValue || dto.Cod_Inte.Value <= 0)
                {
                    return BadRequest("Cod_Inte inválido");
                }

                if (string.IsNullOrWhiteSpace(dto.CodiMovi))
                {
                    return BadRequest("CodiMovi requerido");
                }

                var codInte = dto.Cod_Inte.Value;
                var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == codInte);

                var mov = new Detamasu
                {
                    CodiMovi = dto.CodiMovi.Trim(),
                    Cod_Inte = codInte,
                    Cod_Exte = master?.Codigo,
                    Talon = dto.Talon,
                    NReparación = 0,
                    Nume_Clie = dto.Nume_Clie,
                    NumRemito = null,
                    Observaciones = string.Empty,
                    Referencia = dto.Referencia,
                    FechMov = dto.FechMov ?? DateTime.Now,
                    FeRemovi = DateTime.Now,
                    CantIngre = dto.Cantidad >= 0 ? Math.Round(dto.Cantidad, 4) : 0m,
                    CantSalid = dto.Cantidad < 0 ? Math.Round(Math.Abs(dto.Cantidad), 4) : 0m,
                };

                _context.Detamasus.Add(mov);
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear movimiento DETAMASU");
                return StatusCode(500, $"Error al crear movimiento. Detalle: {ex.Message}");
            }
        }

        [HttpPut("ajustar-cantidad")]
        public async Task<ActionResult> AjustarCantidad([FromBody] DetamasuAjusteCantidadDTO dto)
        {
            try
            {
                if (dto.Talon <= 0)
                {
                    return BadRequest("Talon inválido");
                }

                if (dto.Cod_Inte <= 0)
                {
                    return BadRequest("Cod_Inte inválido");
                }

                var actual = await _context.Detamasus
                    .Where(x => x.Talon == dto.Talon && x.Cod_Inte == dto.Cod_Inte)
                    .GroupBy(x => 1)
                    .Select(g => (g.Sum(x => (decimal?)(x.CantIngre ?? 0m)) ?? 0m) - (g.Sum(x => (decimal?)(x.CantSalid ?? 0m)) ?? 0m))
                    .FirstOrDefaultAsync();

                var diff = Math.Round(dto.NuevaCantidad - actual, 4);
                if (Math.Abs(diff) < 0.00005m)
                {
                    return Ok();
                }

                var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == dto.Cod_Inte);

                var mov = new Detamasu
                {
                    CodiMovi = string.IsNullOrWhiteSpace(dto.CodiMovi) ? "AJ" : dto.CodiMovi!.Trim(),
                    Cod_Inte = dto.Cod_Inte,
                    Cod_Exte = master?.Codigo,
                    Talon = dto.Talon,
                    NReparación = 0,
                    Nume_Clie = dto.Nume_Clie,
                    Observaciones = string.Empty,
                    Referencia = dto.Referencia,
                    FechMov = dto.FechMov ?? DateTime.Now,
                    FeRemovi = DateTime.Now,
                    CantIngre = diff >= 0 ? Math.Round(diff, 4) : 0m,
                    CantSalid = diff < 0 ? Math.Round(Math.Abs(diff), 4) : 0m,
                };

                _context.Detamasus.Add(mov);
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al ajustar cantidad DETAMASU Talon={Talon} CodInte={CodInte}", dto.Talon, dto.Cod_Inte);
                return StatusCode(500, $"Error al ajustar cantidad. Detalle: {ex.Message}");
            }
        }

        [HttpPost("salida")]
        public async Task<ActionResult> RegistrarSalida([FromBody] DetamasuSalidaCreateDTO dto)
        {
            try
            {
                if (dto.Talon <= 0)
                {
                    return BadRequest("Talon inválido");
                }

                if (dto.Cod_Inte <= 0)
                {
                    return BadRequest("Cod_Inte inválido");
                }

                if (dto.Cantidad <= 0)
                {
                    return BadRequest("La cantidad debe ser mayor a 0");
                }

                var movi = (dto.CodiMovi ?? string.Empty).Trim().ToUpperInvariant();
                if (movi != "CF" && movi != "DV")
                {
                    return BadRequest("CodiMovi inválido. Debe ser CF o DV");
                }

                var disponible = await _context.Detamasus
                    .Where(x => x.Talon == dto.Talon && x.Cod_Inte == dto.Cod_Inte)
                    .GroupBy(x => 1)
                    .Select(g => (g.Sum(x => (decimal?)(x.CantIngre ?? 0m)) ?? 0m) - (g.Sum(x => (decimal?)(x.CantSalid ?? 0m)) ?? 0m))
                    .FirstOrDefaultAsync();

                if (dto.Cantidad > disponible)
                {
                    return BadRequest($"Cantidad inválida. Disponible: {disponible:0.####}");
                }

                var master = await _context.masters.FirstOrDefaultAsync(m => m.Codint == dto.Cod_Inte);

                var mov = new Detamasu
                {
                    CodiMovi = movi,
                    Cod_Inte = dto.Cod_Inte,
                    Cod_Exte = master?.Codigo,
                    Talon = dto.Talon,
                    NReparación = 0,
                    Nume_Clie = dto.Nume_Clie,
                    Observaciones = string.Empty,
                    Referencia = dto.Referencia,
                    FechMov = dto.FechMov ?? DateTime.Now,
                    FeRemovi = DateTime.Now,
                    CantIngre = 0m,
                    CantSalid = Math.Round(dto.Cantidad, 4),
                };

                _context.Detamasus.Add(mov);
                await _context.SaveChangesAsync();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al registrar salida DETAMASU Talon={Talon} CodInte={CodInte}", dto.Talon, dto.Cod_Inte);
                return StatusCode(500, $"Error al registrar salida. Detalle: {ex.Message}");
            }
        }
    }
}
