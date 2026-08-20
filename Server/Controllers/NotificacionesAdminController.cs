using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.DTOs.Notificaciones;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/notificaciones/admin")]
    [Authorize]
    public class NotificacionesAdminController : ControllerBase
    {
        private readonly NotificacionesService _service;
        private readonly ILogger<NotificacionesAdminController> _logger;

        public NotificacionesAdminController(NotificacionesService service, ILogger<NotificacionesAdminController> logger)
        {
            _service = service;
            _logger = logger;
        }

        [HttpGet]
        public async Task<IActionResult> ListAsync([FromQuery] PaginationDTO pagination, CancellationToken cancellationToken)
        {
            try
            {
                return Ok(await _service.AdminListAsync(pagination, cancellationToken));
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al listar notificaciones (admin)");
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("totalPages")]
        public async Task<IActionResult> TotalPagesAsync([FromQuery] PaginationDTO pagination, CancellationToken cancellationToken)
        {
            try
            {
                return Ok(await _service.AdminTotalPagesAsync(pagination, cancellationToken));
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al calcular totalPages notificaciones (admin)");
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("{id:int}")]
        public async Task<IActionResult> GetByIdAsync([FromRoute] int id, CancellationToken cancellationToken)
        {
            try
            {
                var dto = await _service.AdminGetByIdAsync(id, cancellationToken);
                return dto == null ? NotFound() : Ok(dto);
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener notificación (admin) {Id}", id);
                return BadRequest(ex.Message);
            }
        }

        [HttpPost]
        public async Task<IActionResult> CreateAsync([FromBody] NotificacionAdminUpsertDTO dto, CancellationToken cancellationToken)
        {
            try
            {
                var id = await _service.AdminCreateAsync(dto, cancellationToken);
                return Ok(id);
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear notificación (admin)");
                return BadRequest(ex.Message);
            }
        }

        [HttpPut("{id:int}")]
        public async Task<IActionResult> UpdateAsync([FromRoute] int id, [FromBody] NotificacionAdminUpsertDTO dto, CancellationToken cancellationToken)
        {
            try
            {
                await _service.AdminUpdateAsync(id, dto, cancellationToken);
                return NoContent();
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al actualizar notificación (admin) {Id}", id);
                return BadRequest(ex.Message);
            }
        }

        [HttpDelete("{id:int}")]
        public async Task<IActionResult> DeleteAsync([FromRoute] int id, CancellationToken cancellationToken)
        {
            try
            {
                await _service.AdminDeleteAsync(id, cancellationToken);
                return NoContent();
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al borrar notificación (admin) {Id}", id);
                return BadRequest(ex.Message);
            }
        }
    }
}
