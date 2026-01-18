using IMPLANPROD.Server.Services;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [Microsoft.AspNetCore.Authorization.Authorize]
    public class MigrationController : ControllerBase
    {
        private readonly MigrationLegacyService _migrationService;

        public MigrationController(MigrationLegacyService migrationService)
        {
            _migrationService = migrationService;
        }

        [HttpPost("indirep")]
        public async Task<ActionResult<int>> MigrateIndirep([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateIndirepAsync(basePath);
                if (count == 0)
                {
                    return Ok("La tabla ya contenía datos o no se encontraron registros para migrar.");
                }
                return Ok($"Se migraron {count} registros exitosamente.");
            }
            catch (FileNotFoundException ex)
            {
                return NotFound(ex.Message);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }

        [HttpPost("repuesre")]
        public async Task<ActionResult<int>> MigrateRepuesre([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateRepuesreAsync(basePath);
                if (count == 0)
                {
                    return Ok("La tabla ya contenía datos o no se encontraron registros para migrar.");
                }
                return Ok($"Se migraron {count} registros exitosamente.");
            }
            catch (FileNotFoundException ex)
            {
                return NotFound(ex.Message);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }

        [HttpPost("secoprep")]
        public async Task<ActionResult<int>> MigrateSecoprep([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateSecoprepAsync(basePath);
                if (count == 0)
                {
                    return Ok("La tabla ya contenía datos o no se encontraron registros para migrar.");
                }
                return Ok($"Se migraron {count} registros exitosamente.");
            }
            catch (FileNotFoundException ex)
            {
                return NotFound(ex.Message);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }

        [HttpPost("obsoprep")]
        public async Task<ActionResult<int>> MigrateObsoprep([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateObsoprepAsync(basePath);
                if (count == 0)
                {
                    return Ok("No se encontraron registros para actualizar o el archivo no existe.");
                }
                return Ok($"Se actualizaron {count} registros exitosamente con observaciones.");
            }
            catch (FileNotFoundException ex)
            {
                return NotFound(ex.Message);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }

        [HttpPost("descrep")]
        public async Task<ActionResult<int>> MigrateDescrep([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateDescrepAsync(basePath);
                if (count == 0)
                {
                    return Ok("La tabla ya contenía datos o no se encontraron registros para migrar.");
                }
                return Ok($"Se migraron {count} registros exitosamente.");
            }
            catch (FileNotFoundException ex)
            {
                return NotFound(ex.Message);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error interno: {ex.Message}");
            }
        }
    }
}
