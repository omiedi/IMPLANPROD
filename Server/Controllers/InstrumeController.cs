using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class InstrumeController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<InstrumeController> _logger;

        public InstrumeController(DataContext context, ILogger<InstrumeController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Búsqueda rápida para autocomplete: filtra por CodIntMedIdent, Codigo o Descripcion.
        /// Si filter está vacío, devuelve todos los registros activos hasta pageSize.
        /// Devuelve hasta pageSize registros activos.
        /// </summary>
        [HttpGet("search")]
        public async Task<ActionResult<List<Instrume>>> Search(
            [FromQuery] string? filter = null,
            [FromQuery] int pageSize = 15)
        {
            try
            {
                if (pageSize > 200) pageSize = 200;

                var query = _context.Instrumes
                    .AsNoTracking()
                    .Where(x => x.Status >= 0);

                if (!string.IsNullOrWhiteSpace(filter))
                {
                    var f = filter.Trim().ToLower();
                    query = query.Where(x =>
                        EF.Functions.Like((x.Codigo      ?? "").ToLower(), $"%{f}%") ||
                        EF.Functions.Like((x.Descripcion ?? "").ToLower(), $"%{f}%") ||
                        x.CodIntMedIdent.ToString()!.Contains(f)
                    );
                }

                var result = await query
                    .OrderBy(x => x.Codigo)
                    .Take(pageSize)
                    .Select(x => new Instrume
                    {
                        Id             = x.Id,
                        CodIntMedIdent = x.CodIntMedIdent,
                        Codigo         = x.Codigo,
                        Descripcion    = x.Descripcion
                    })
                    .ToListAsync();

                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error en búsqueda de instrumentos");
                return StatusCode(500, $"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Devuelve el detalle completo de un instrumento por su Id.
        /// </summary>
        [HttpGet("{id:int}")]
        public async Task<ActionResult<Instrume>> GetById(int id)
        {
            try
            {
                var instrume = await _context.Instrumes
                    .AsNoTracking()
                    .FirstOrDefaultAsync(x => x.Id == id);

                if (instrume == null)
                    return NotFound($"No se encontró el instrumento con Id={id}.");

                return Ok(instrume);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener instrumento Id={Id}", id);
                return StatusCode(500, $"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Actualiza un instrumento existente.
        /// PUT /api/instrume/{id}
        /// </summary>
        [HttpPut("{id:int}")]
        public async Task<ActionResult> Update(int id, [FromBody] Instrume dto)
        {
            try
            {
                var entity = await _context.Instrumes.FirstOrDefaultAsync(x => x.Id == id);
                if (entity == null)
                    return NotFound($"No se encontró el instrumento con Id={id}.");

                // Actualizar campos editables
                entity.Tipo = dto.Tipo;
                entity.Marca = dto.Marca;
                entity.Modelo = dto.Modelo;
                entity.NroSerie = dto.NroSerie;
                entity.Rango = dto.Rango;
                entity.Posicion = dto.Posicion;
                entity.Ubicacion = dto.Ubicacion;
                entity.NroGrupo = dto.NroGrupo;
                entity.IntervCalibra = dto.IntervCalibra;
                entity.CodIntMedCalib = dto.CodIntMedCalib;
                entity.CodIntMedIdent = dto.CodIntMedIdent;
                entity.Observaciones = dto.Observaciones;
                entity.LastUpdate = DateTime.Now;

                await _context.SaveChangesAsync();
                
                _logger.LogInformation("Instrumento Id={Id} actualizado correctamente", id);
                return NoContent();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar instrumento Id={Id}", id);
                return StatusCode(500, $"Error: {ex.Message}");
            }
        }
    }
}
