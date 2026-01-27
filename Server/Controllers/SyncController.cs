using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para sincronización entre Central y Sucursal
    /// </summary>
    [ApiController]
    [Route("api/sync")]
    public class SyncController : ControllerBase
    {
        private readonly ISyncService _syncService;
        private readonly ILogger<SyncController> _logger;

        public SyncController(ISyncService syncService, ILogger<SyncController> logger)
        {
            _syncService = syncService;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene la configuración actual de sincronización
        /// </summary>
        [HttpGet("config")]
        [Authorize(Roles = "Admin")]
        public ActionResult<SyncConfigDTO> GetConfig()
        {
            try
            {
                var config = _syncService.ObtenerConfiguracion();
                return Ok(config);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener configuración de sincronización");
                return StatusCode(500, $"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Guarda la configuración de sincronización
        /// </summary>
        [HttpPost("config")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> SaveConfig([FromBody] SyncConfigDTO config)
        {
            try
            {
                var resultado = await _syncService.GuardarConfiguracionAsync(config);
                if (resultado)
                {
                    return Ok(new { mensaje = "Configuración guardada exitosamente" });
                }
                else
                {
                    return BadRequest("No se pudo guardar la configuración");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al guardar configuración de sincronización");
                return StatusCode(500, $"Error: {ex.Message}");
            }
        }

        /// <summary>
        /// Procesa archivos de sincronización pendientes (Sucursal)
        /// </summary>
        [HttpPost("procesar")]
        public async Task<ActionResult<SyncResultDTO>> ProcesarSync([FromBody] ValidarRutaDTO request)
        {
            try
            {
                var resultado = await _syncService.ProcesarArchivosSyncAsync(request.Ruta);
                return Ok(resultado);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al procesar archivos de sincronización");
                return StatusCode(500, new SyncResultDTO
                {
                    Exitoso = false,
                    Mensaje = $"Error: {ex.Message}"
                });
            }
        }

        /// <summary>
        /// Valida que una ruta de intercambio sea accesible
        /// </summary>
        [HttpPost("validar-ruta")]
        [Authorize(Roles = "Admin")]
        public ActionResult<ValidarRutaResponseDTO> ValidarRuta([FromBody] ValidarRutaDTO request)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(request.Ruta))
                {
                    return BadRequest("La ruta no puede estar vacía");
                }

                if (Directory.Exists(request.Ruta))
                {
                    return Ok(new ValidarRutaResponseDTO { Valida = true, Mensaje = "Ruta accesible" });
                }
                else
                {
                    return Ok(new ValidarRutaResponseDTO { Valida = false, Mensaje = "La ruta no existe o no es accesible" });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al validar ruta: {request.Ruta}");
                return Ok(new ValidarRutaResponseDTO { Valida = false, Mensaje = $"Error: {ex.Message}" });
            }
        }
    }
}
