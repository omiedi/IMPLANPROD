using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs.Notificaciones;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/notificaciones/reglas-sql")]
    [Authorize]
    public class NotificacionReglasSqlController : ControllerBase
    {
        private readonly NotificacionReglasSqlService _service;
        private readonly ILogger<NotificacionReglasSqlController> _logger;

        public NotificacionReglasSqlController(NotificacionReglasSqlService service, ILogger<NotificacionReglasSqlController> logger)
        {
            _service = service;
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> GetAllAsync(CancellationToken cancellationToken)
        {
            try
            {
                return Ok(await _service.GetAllAsync(cancellationToken));
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al listar reglas SQL");
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetByIdAsync([FromRoute] int id, CancellationToken cancellationToken)
        {
            try
            {
                var dto = await _service.GetByIdAsync(id, cancellationToken);
                return dto == null ? NotFound() : Ok(dto);
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener regla SQL {Id}", id);
                return BadRequest(ex.Message);
            }
        }

        [HttpPost]
        public async Task<IActionResult> CreateAsync([FromBody] NotificacionReglaSqlUpsertDTO dto, CancellationToken cancellationToken)
        {
            try
            {
                var id = await _service.CreateAsync(dto, cancellationToken);
                return Ok(id);
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear regla SQL");
                return BadRequest(ex.Message);
            }
        }

        [HttpPut("{id:int}")]
        public async Task<IActionResult> UpdateAsync([FromRoute] int id, [FromBody] NotificacionReglaSqlUpsertDTO dto, CancellationToken cancellationToken)
        {
            try
            {
                await _service.UpdateAsync(id, dto, cancellationToken);
                return NoContent();
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar regla SQL {Id}", id);
                return BadRequest(ex.Message);
            }
        }

        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync([FromRoute] int id, CancellationToken cancellationToken)
        {
            try
            {
                await _service.DeleteAsync(id, cancellationToken);
                return NoContent();
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al eliminar regla SQL {Id}", id);
                return BadRequest(ex.Message);
            }
        }
    }
}
