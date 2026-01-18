using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using IMPLANPROD.Server.Data;
using Microsoft.AspNetCore.Authorization;
using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize]
    public class IndirepController : ControllerBase
    {
        private readonly DataContext _context;
        private readonly ILogger<IndirepController> _logger;

        public IndirepController(DataContext context, ILogger<IndirepController> logger)
        {
            _context = context;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene lista de códigos de reparación para combo desde INDIREP
        /// </summary>
        [HttpGet("combo")]
        public async Task<ActionResult<List<Indirep>>> GetCombo()
        {
            try
            {
                var codigos = await _context.Indireps
                    .Where(i => i.Status >= 0)
                    .OrderBy(i => i.Codigo)
                    .Select(i => new
                    {
                        Codigo = i.Codigo,
                        Descripcion = i.Descripcion
                    })
                    .ToListAsync();

                return Ok(codigos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener combo de códigos de reparación desde INDIREP");
                return StatusCode(500, "Error al obtener los códigos");
            }
        }
    }
}
