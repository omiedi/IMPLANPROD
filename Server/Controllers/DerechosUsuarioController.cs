using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    /// <summary>
    /// Controlador para gestión de derechos de usuarios
    /// </summary>
    [ApiController]
    [Route("api/[controller]")]
    [Authorize(AuthenticationSchemes = JwtBearerDefaults.AuthenticationScheme)]
    public class DerechosUsuarioController : ControllerBase
    {
        private readonly IDerechosUsuarioService _derechosService;
        private readonly ILogger<DerechosUsuarioController> _logger;

        public DerechosUsuarioController(
            IDerechosUsuarioService derechosService,
            ILogger<DerechosUsuarioController> logger)
        {
            _derechosService = derechosService;
            _logger = logger;
        }

        /// <summary>
        /// Obtiene todos los derechos de un usuario específico
        /// GET: api/DerechosUsuario/usuario/{numUser}
        /// </summary>
        [HttpGet("usuario/{numUser}")]
        public async Task<ActionResult<List<DerechoUsuarioDTO>>> GetDerechosUsuario(short numUser)
        {
            try
            {
                _logger.LogInformation($"Obteniendo derechos para usuario {numUser}");
                var derechos = await _derechosService.ObtenerDerechosUsuarioAsync(numUser);
                return Ok(derechos);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al obtener derechos del usuario {numUser}");
                return StatusCode(500, "Error al obtener los derechos del usuario");
            }
        }

        /// <summary>
        /// Guarda o actualiza los derechos de un usuario
        /// POST: api/DerechosUsuario/usuario/{numUser}
        /// </summary>
        [HttpPost("usuario/{numUser}")]
        public async Task<ActionResult> GuardarDerechosUsuario(short numUser, [FromBody] List<DerechoUsuarioDTO> derechos)
        {
            try
            {
                _logger.LogInformation($"Guardando derechos para usuario {numUser}");
                var resultado = await _derechosService.GuardarDerechosUsuarioAsync(numUser, derechos);

                if (resultado)
                {
                    return Ok(new { mensaje = "Derechos guardados correctamente" });
                }
                else
                {
                    return BadRequest("Error al guardar los derechos");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al guardar derechos del usuario {numUser}");
                return StatusCode(500, "Error al guardar los derechos del usuario");
            }
        }

        /// <summary>
        /// Verifica si un usuario tiene derecho de lectura sobre una aplicación
        /// GET: api/DerechosUsuario/verificar-lectura/{numUser}/{codApli}
        /// </summary>
        [HttpGet("verificar-lectura/{numUser}/{codApli}")]
        public async Task<ActionResult<bool>> VerificarDerechoLectura(short numUser, string codApli)
        {
            try
            {
                var tieneDerecho = await _derechosService.TieneDerechoLecturaAsync(numUser, codApli);
                return Ok(tieneDerecho);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al verificar derecho de lectura para usuario {numUser}, aplicación {codApli}");
                return StatusCode(500, "Error al verificar el derecho");
            }
        }

        /// <summary>
        /// Verifica si un usuario tiene derecho de escritura sobre una aplicación
        /// GET: api/DerechosUsuario/verificar-escritura/{numUser}/{codApli}
        /// </summary>
        [HttpGet("verificar-escritura/{numUser}/{codApli}")]
        public async Task<ActionResult<bool>> VerificarDerechoEscritura(short numUser, string codApli)
        {
            try
            {
                var tieneDerecho = await _derechosService.TieneDerechoEscrituraAsync(numUser, codApli);
                return Ok(tieneDerecho);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al verificar derecho de escritura para usuario {numUser}, aplicación {codApli}");
                return StatusCode(500, "Error al verificar el derecho");
            }
        }

        /// <summary>
        /// Verifica los derechos de un usuario sobre una aplicación (funcionalidad DERACCE del VB6)
        /// GET: api/DerechosUsuario/verificar/{numUser}/{codApli}/{desApli}
        /// Retorna:
        /// 0 = Sin acceso
        /// 1 = Solo lectura
        /// 2 = Lectura y escritura
        /// </summary>
        [HttpGet("verificar/{numUser}/{codApli}/{desApli}")]
        public async Task<ActionResult<int>> VerificarDerecho(short numUser, string codApli, string desApli)
        {
            try
            {
                _logger.LogInformation($"Verificando derecho DERACCE para usuario {numUser}, aplicación {codApli}");
                var nivelAcceso = await _derechosService.VerificarDerechoAsync(numUser, codApli, desApli);
                return Ok(nivelAcceso);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error al verificar derecho para usuario {numUser}, aplicación {codApli}");
                return StatusCode(500, "Error al verificar el derecho");
            }
        }
    }
}
