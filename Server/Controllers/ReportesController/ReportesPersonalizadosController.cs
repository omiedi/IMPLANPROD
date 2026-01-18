using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.Entities;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Shared.Entities;

namespace IMPLANPROD.Server.Controllers.ReportesController
{
    /// <summary>
    /// Controlador para la gestión de reportes personalizados.
    /// Permite crear, listar y obtener reportes definidos por el usuario.
    /// </summary>
    [ApiController]
    [Route("/api/reportes-personalizados")]
    public class ReportesPersonalizadosController : ControllerBase
    {
        private readonly DataContext _context;

        /// <summary>
        /// Constructor con inyección del contexto de datos.
        /// </summary>
        public ReportesPersonalizadosController(DataContext context)
        {
            _context = context;
        }

        /// <summary>
        /// Obtiene la lista de todos los reportes personalizados guardados.
        /// </summary>
        [HttpGet]
        public async Task<ActionResult<IEnumerable<ReportePersonalizado>>> GetReportes()
        {
            try
            {
                return await _context.ReportesPersonalizados.OrderBy(r => r.Nombre).ToListAsync();
            }
            catch (Exception ex)
            {
                // Retorna error 500 con mensaje en español
                return StatusCode(500, $"Error al obtener la lista de reportes personalizados: {ex.Message}");
            }
        }

        /// <summary>
        /// Obtiene un reporte personalizado por su Id.
        /// </summary>
        [HttpGet("{id}")]
        public async Task<ActionResult<ReportePersonalizado>> GetReporte(int id)
        {
            try
            {
                var reporte = await _context.ReportesPersonalizados.FindAsync(id);
                if (reporte == null)
                    return NotFound();
                return reporte;
            }
            catch (Exception ex)
            {
                // Retorna error 500 con mensaje en español
                return StatusCode(500, $"Error al obtener el reporte: {ex.Message}");
            }
        }

        /// <summary>
        /// Crea o actualiza un reporte personalizado.
        /// Si el Id es 0, crea uno nuevo. Si existe, lo actualiza.
        /// </summary>
        [HttpPost]
        public async Task<ActionResult<ReportePersonalizado>> GuardarReporte([FromBody] ReportePersonalizado reporte)
        {
            try
            {
                if (reporte.Id == 0)
                {
                    reporte.FechaCreacion = DateTime.Now;
                    reporte.FechaModificacion = DateTime.Now;
                    _context.ReportesPersonalizados.Add(reporte);
                    await _context.SaveChangesAsync();
                    return CreatedAtAction(nameof(GetReporte), new { id = reporte.Id }, reporte);
                }
                else
                {
                    var existente = await _context.ReportesPersonalizados.FindAsync(reporte.Id);
                    if (existente == null)
                        return NotFound();
                    existente.Nombre = reporte.Nombre;
                    existente.Usuario = reporte.Usuario;
                    existente.ConfiguracionJson = reporte.ConfiguracionJson;
                    existente.FechaModificacion = DateTime.Now;
                    await _context.SaveChangesAsync();
                    return Ok(existente);
                }
            }
            catch (Exception ex)
            {
                // Retorna error 500 con mensaje en español
                return StatusCode(500, $"Error al guardar el reporte personalizado: {ex.Message}");
            }
        }

        /// <summary>
        /// Elimina un reporte personalizado por Id.
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<IActionResult> EliminarReporte(int id)
        {
            try
            {
                var reporte = await _context.ReportesPersonalizados.FindAsync(id);
                if (reporte == null)
                    return NotFound();
                _context.ReportesPersonalizados.Remove(reporte);
                await _context.SaveChangesAsync();
                return NoContent();
            }
            catch (Exception ex)
            {
                // Retorna error 500 con mensaje en español
                return StatusCode(500, $"Error al eliminar el reporte personalizado: {ex.Message}");
            }
        }
    }
}
