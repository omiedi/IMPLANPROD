using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/email-empresa")]
    [Authorize]
    public class EmailConfigEmpresaController : ControllerBase
    {
        private readonly IEmailEmpresaService _service;

        public EmailConfigEmpresaController(IEmailEmpresaService service)
        {
            _service = service;
        }

        [HttpGet]
        public async Task<ActionResult<EmailConfigEmpresaDto>> Get()
        {
            try
            {
                var config = await _service.GetConfiguracionAsync();
                return Ok(config);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener la configuración: {ex.Message}");
            }
        }

        [HttpPost]
        public async Task<ActionResult> Post([FromBody] EmailConfigEmpresaDto dto)
        {
            if (!ModelState.IsValid)
                return BadRequest(ModelState);

            var esSuperAdmin   = User.IsInRole("SUPERADMIN");
            var esAdministrador = User.IsInRole("ADMINISTRADOR") || User.IsInRole("ADMIN");

            if (!esSuperAdmin && !esAdministrador)
                return Forbid();

            try
            {
                if (string.IsNullOrWhiteSpace(dto.Contrasena) && !dto.TieneContrasenaGuardada)
                    return BadRequest("Debe ingresar una contraseña la primera vez que configura la cuenta.");

                await _service.GuardarConfiguracionAsync(dto);
                return Ok("Configuración guardada correctamente");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al guardar la configuración: {ex.Message}");
            }
        }

        [HttpPost("sendorder")]
        public async Task<ActionResult> SendOrder([FromBody] SendOrdenEmailDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.To))
                return BadRequest("Debe indicar el email del destinatario.");

            if (dto.AttachmentContent == null || dto.AttachmentContent.Length == 0)
                return BadRequest("No se proporcionó el archivo adjunto.");

            try
            {
                await _service.EnviarEmailAsync(
                    dto.To,
                    dto.Subject,
                    dto.Body,
                    isHtml: true,
                    replyTo: dto.ReplyTo,
                    replyToName: dto.ReplyToName,
                    attachmentContent: dto.AttachmentContent,
                    attachmentName: dto.AttachmentName
                );
                return Ok("Correo enviado correctamente");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al enviar el correo: {ex.Message}");
            }
        }

        [HttpPost("test")]
        public async Task<ActionResult> Test([FromBody] EmailConfigEmpresaDto dto)
        {
            var esSuperAdmin    = User.IsInRole("SUPERADMIN");
            var esAdministrador = User.IsInRole("ADMINISTRADOR") || User.IsInRole("ADMIN");

            if (!esSuperAdmin && !esAdministrador)
                return Forbid();

            if (string.IsNullOrWhiteSpace(dto.DestinatarioPrueba))
                return BadRequest("Debe ingresar un destinatario para el correo de prueba.");

            try
            {
                if (!string.IsNullOrWhiteSpace(dto.Contrasena) || dto.TieneContrasenaGuardada)
                    await _service.GuardarConfiguracionAsync(dto);

                await _service.EnviarEmailPruebaAsync(dto.DestinatarioPrueba);
                return Ok("Correo de prueba enviado correctamente");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al enviar el correo de prueba: {ex.Message}");
            }
        }
    }
}
