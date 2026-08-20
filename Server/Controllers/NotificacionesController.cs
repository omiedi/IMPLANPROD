using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.DTOs.Notificaciones;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/notificaciones")]
    [Authorize]
    public class NotificacionesController : ControllerBase
    {
        private readonly NotificacionesService _service;
        private readonly ILogger<NotificacionesController> _logger;

        public NotificacionesController(NotificacionesService service, ILogger<NotificacionesController> logger)
        {
            _service = service;
            _logger = logger;
        }

        [HttpPost]
        public async Task<IActionResult> CrearAsync([FromBody] CrearNotificacionRequestDTO request, CancellationToken cancellationToken)
        {
            try
            {
                var id = await _service.CrearAsync(request, cancellationToken);
                return Ok(id);
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al crear notificación");
                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        public async Task<IActionResult> ListarAsync([FromQuery] PaginationDTO pagination, CancellationToken cancellationToken)
        {
            try
            {
                var items = await _service.ListarParaUsuarioActualAsync(pagination, cancellationToken);
                return Ok(items);
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al listar notificaciones");
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("totalPages")]
        public async Task<IActionResult> TotalPagesAsync([FromQuery] PaginationDTO pagination, CancellationToken cancellationToken)
        {
            try
            {
                var totalPages = await _service.TotalPagesParaUsuarioActualAsync(pagination, cancellationToken);
                return Ok(totalPages);
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al calcular totalPages de notificaciones");
                return BadRequest(ex.Message);
            }
        }

        [HttpGet("unreadCount")]
        public async Task<IActionResult> ContarNoLeidasAsync(CancellationToken cancellationToken)
        {
            try
            {
                var count = await _service.ContarNoLeidasAsync(cancellationToken);
                return Ok(count);
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al contar no leídas");
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("markRead")]
        public async Task<IActionResult> MarcarLeidaAsync([FromBody] MarcarLeidaRequestDTO request, CancellationToken cancellationToken)
        {
            try
            {
                await _service.MarcarLeidaAsync(request.NotificacionId, cancellationToken);
                return NoContent();
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (DbUpdateException ex)
            {
                _logger.LogError(ex, "Error de base de datos al marcar leída");
                return BadRequest(GetInnermostMessage(ex));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al marcar leída");
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("hide")]
        public async Task<IActionResult> OcultarAsync([FromBody] OcultarNotificacionRequestDTO request, CancellationToken cancellationToken)
        {
            try
            {
                await _service.OcultarParaUsuarioActualAsync(request.NotificacionId, cancellationToken);
                return NoContent();
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (DbUpdateException ex)
            {
                _logger.LogError(ex, "Error de base de datos al ocultar notificación");
                return BadRequest(GetInnermostMessage(ex));
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al ocultar notificación");
                return BadRequest(ex.Message);
            }
        }

        private static string GetInnermostMessage(Exception ex)
        {
            var current = ex;
            while (current.InnerException != null)
            {
                current = current.InnerException;
            }

            return current.Message;
        }

        [HttpGet("{id:int}/detail")]
        public async Task<IActionResult> ObtenerDetalleAsync([FromRoute] int id, CancellationToken cancellationToken)
        {
            try
            {
                var dto = await _service.ObtenerDetalleAsync(id, cancellationToken);
                return Ok(dto);
            }
            catch (UnauthorizedAccessException)
            {
                return Unauthorized();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error al obtener detalle de notificación {Id}", id);
                return BadRequest(ex.Message);
            }
        }
    }
}
