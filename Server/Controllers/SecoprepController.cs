using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme)]
    [Route("api/[controller]")]
    public class SecoprepController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<SecoprepController> _logger;

        public SecoprepController(DataContext context, ILogger<SecoprepController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene las operaciones de una reparación ordenadas por número de operación
        /// </summary>
        [HttpGet("reparacion/{codigoReparacion}")]
        public async Task<ActionResult<List<Secoprep>>> GetOperacionesPorReparacion(string codigoReparacion)
        {
            try
            {
                var operaciones = await _context.Secopreps
                    .Where(s => s.CodigoReparacion == codigoReparacion && s.Status >= 0)
                    .OrderBy(s => s.NumeroOperacion)
                    .ToListAsync();

                return Ok(operaciones);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener operaciones de reparación {CodigoReparacion}", codigoReparacion);
                return StatusCode(500, "Error al obtener las operaciones");
            }
        }

        /// <summary>
        /// Obtiene una operación específica por código de reparación y número de operación
        /// </summary>
        [HttpGet("{codigoReparacion}/{numeroOperacion}")]
        public async Task<ActionResult<Secoprep>> GetOperacion(string codigoReparacion, short numeroOperacion)
        {
            try
            {
                var operacion = await _context.Secopreps
                    .Where(s => s.CodigoReparacion == codigoReparacion && 
                                s.NumeroOperacion == numeroOperacion && 
                                s.Status >= 0)
                    .FirstOrDefaultAsync();

                if (operacion == null)
                {
                    return NotFound("Operación no encontrada");
                }

                return Ok(operacion);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener operación {CodigoReparacion}/{NumeroOperacion}", 
                    codigoReparacion, numeroOperacion);
                return StatusCode(500, "Error al obtener la operación");
            }
        }

        /// <summary>
        /// Crea una nueva operación para una reparación
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<Secoprep>> CrearOperacion([FromBody] Secoprep operacion)
        {
            try
            {
                // Verificar si ya existe una operación con el mismo código y número
                var existe = await _context.Secopreps
                    .AnyAsync(s => s.CodigoReparacion == operacion.CodigoReparacion && 
                                   s.NumeroOperacion == operacion.NumeroOperacion);

                if (existe)
                {
                    return BadRequest("Ya existe una operación con este número para esta reparación");
                }

                operacion.Status = 0;
                _context.Secopreps.Add(operacion);
                await _context.SaveChangesAsync();

                _logger.LogInformation("Operación creada: {CodigoReparacion}/{NumeroOperacion}", 
                    operacion.CodigoReparacion, operacion.NumeroOperacion);

                return CreatedAtAction(nameof(GetOperacion), 
                    new { codigoReparacion = operacion.CodigoReparacion, numeroOperacion = operacion.NumeroOperacion }, 
                    operacion);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear operación");
                return StatusCode(500, "Error al crear la operación");
            }
        }

        /// <summary>
        /// Actualiza una operación existente
        /// </summary>
        [HttpPut("{id}")]
        public async Task<ActionResult> ActualizarOperacion(int id, [FromBody] Secoprep operacion)
        {
            try
            {
                var operacionExistente = await _context.Secopreps.FindAsync(id);

                if (operacionExistente == null)
                {
                    return NotFound("Operación no encontrada");
                }

                operacionExistente.Descripcion = operacion.Descripcion;
                operacionExistente.Observacion = operacion.Observacion;

                await _context.SaveChangesAsync();

                _logger.LogInformation("Operación actualizada: {Id}", id);

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar operación {Id}", id);
                return StatusCode(500, "Error al actualizar la operación");
            }
        }

        /// <summary>
        /// Elimina una operación (marca como inactiva)
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<ActionResult> EliminarOperacion(int id)
        {
            try
            {
                var operacion = await _context.Secopreps.FindAsync(id);

                if (operacion == null)
                {
                    return NotFound("Operación no encontrada");
                }

                operacion.Status = -1;
                await _context.SaveChangesAsync();

                _logger.LogInformation("Operación eliminada: {Id}", id);

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al eliminar operación {Id}", id);
                return StatusCode(500, "Error al eliminar la operación");
            }
        }

        /// <summary>
        /// Obtiene el siguiente número de operación disponible para una reparación
        /// </summary>
        [HttpGet("siguiente-numero/{codigoReparacion}")]
        public async Task<ActionResult<short>> GetSiguienteNumero(string codigoReparacion)
        {
            try
            {
                var maxNumero = await _context.Secopreps
                    .Where(s => s.CodigoReparacion == codigoReparacion)
                    .MaxAsync(s => (short?)s.NumeroOperacion);

                short siguienteNumero = (short)((maxNumero ?? 0) + 10);

                return Ok(siguienteNumero);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener siguiente número para {CodigoReparacion}", codigoReparacion);
                return StatusCode(500, "Error al obtener el siguiente número");
            }
        }

        /// <summary>
        /// Calcula las horas totales de mano de obra para una reparación
        /// Suma: CantidadOperarios * Horas de cada operación
        /// </summary>
        [HttpGet("horas-mano-obra/{codigoReparacion}")]
        public async Task<ActionResult<double>> GetHorasManoObra(string codigoReparacion)
        {
            try
            {
                var operaciones = await _context.Secopreps
                    .Where(s => s.CodigoReparacion == codigoReparacion && s.Status >= 0)
                    .ToListAsync();

                double totalHoras = 0;

                foreach (var operacion in operaciones)
                {
                    int cantOperarios = operacion.CantidadOperarios ?? 1;
                    if (cantOperarios < 1) cantOperarios = 1;

                    double horas = operacion.Horas ?? 0;
                    totalHoras += cantOperarios * horas;
                }

                return Ok(totalHoras);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al calcular horas de mano de obra {CodigoReparacion}", codigoReparacion);
                return StatusCode(500, "Error al calcular las horas de mano de obra");
            }
        }
    }
}
