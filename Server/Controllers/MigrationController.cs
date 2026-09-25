using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [Microsoft.AspNetCore.Authorization.Authorize]
    public class MigrationController : ControllerBase
    {
        private readonly MigrationLegacyService _migrationService;

        public MigrationController(
            MigrationLegacyService migrationService)
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

        [HttpPost("plainsp")]
        public async Task<ActionResult<int>> MigratePlainsp([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigratePlainspsAsync(basePath);
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

        [HttpPost("codiproc")]
        public async Task<ActionResult<int>> MigrateCodiproc([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateCodiprocAsync(basePath);
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

        [HttpPost("instruci")]
        public async Task<ActionResult<int>> MigrateInstruci([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateInstruciAsync(basePath);
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

        [HttpPost("obinstru")]
        public async Task<ActionResult<int>> MigrateObInstru([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateObInstruAsync(basePath);
                if (count == 0)
                {
                    return Ok("Ya existen registros con Observaciones en INSTRUME o no se encontraron coincidencias.");
                }
                return Ok($"Se actualizaron {count} registros de INSTRUME con Observaciones.");
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

        [HttpPost("instrume")]
        public async Task<ActionResult<int>> MigrateInstrume([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateInstrumeAsync(basePath);
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

        [HttpPost("instrucd")]
        public async Task<ActionResult<int>> MigrateInstrucd([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateInstrucdAsync(basePath);
                if (count == 0)
                {
                    return Ok("Ya existen registros con Observaciones en INSTRUCI o no se encontraron coincidencias.");
                }
                return Ok($"Se actualizaron {count} registros de INSTRUCI con Observaciones.");
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

        [HttpPost("brecepen")]
        public async Task<ActionResult<int>> MigrateBrecepen([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateBrecepenAsync(basePath);
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

        [HttpGet("brecepen/preview")]
        public async Task<ActionResult<List<BrecepenPreviewDTO>>> PreviewBrecepen([FromQuery] string? basePath = null)
        {
            try
            {
                var data = await _migrationService.PreviewBrecepenAsync(basePath);
                return Ok(data);
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

        /// <summary>
        /// Previsualiza los registros del archivo legacy REPLACOP.DAT sin migrarlos.
        /// GET /api/Migration/replacop/preview?basePath=...
        /// </summary>
        [HttpGet("replacop/preview")]
        public async Task<ActionResult<List<ReplacopPreviewDTO>>> PreviewReplacop([FromQuery] string? basePath = null)
        {
            try
            {
                var data = await _migrationService.PreviewReplacopAsync(basePath);
                return Ok(data);
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

        /// <summary>
        /// Migra los datos del archivo legacy REPLACOP.DAT a la tabla SQL REPLACO.
        /// Solo migra si la tabla destino está vacía.
        /// POST /api/Migration/replacop?basePath=...
        /// </summary>
        [HttpPost("replacop")]
        public async Task<ActionResult<int>> MigrateReplacop([FromQuery] string? basePath = null)
        {
            try
            {
                int count = await _migrationService.MigrateReplacopAsync(basePath);
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
