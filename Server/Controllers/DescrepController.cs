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
    public class DescrepController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<DescrepController> _logger;

        public DescrepController(DataContext context, ILogger<DescrepController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene la especificación de una reparación
        /// </summary>
        [HttpGet("{codigoReparacion}")]
        public async Task<ActionResult<Descrep>> GetEspecificacion(string codigoReparacion)
        {
            try
            {
                var descrep = await _context.Descreps
                    .Where(d => d.CodigoReparacion == codigoReparacion && d.Status >= 0)
                    .FirstOrDefaultAsync();

                if (descrep == null)
                {
                    return NotFound("Especificación no encontrada");
                }

                return Ok(descrep);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener especificación {CodigoReparacion}", codigoReparacion);
                return StatusCode(500, "Error al obtener la especificación");
            }
        }

        /// <summary>
        /// Actualiza la especificación de una reparación
        /// </summary>
        [HttpPut("{codigoReparacion}")]
        public async Task<ActionResult> ActualizarEspecificacion(string codigoReparacion, [FromBody] string especificacion)
        {
            try
            {
                var descrep = await _context.Descreps
                    .Where(d => d.CodigoReparacion == codigoReparacion)
                    .FirstOrDefaultAsync();

                if (descrep == null)
                {
                    return NotFound("Especificación no encontrada");
                }

                descrep.Especificacion = especificacion;
                await _context.SaveChangesAsync();

                _logger.LogInformation("Especificación actualizada: {CodigoReparacion}", codigoReparacion);

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar especificación {CodigoReparacion}", codigoReparacion);
                return StatusCode(500, "Error al actualizar la especificación");
            }
        }

    }
}
