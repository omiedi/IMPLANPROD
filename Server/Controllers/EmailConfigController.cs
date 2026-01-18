using IMPLANPROD.Server.Services;
using IMPLANPROD.Shared.DTOs;
using Microsoft.AspNetCore.Mvc;

namespace IMPLANPROD.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class EmailConfigController : ControllerBase
    {
        private readonly IEmailService _emailService;

        public EmailConfigController(IEmailService emailService)
        {
            _emailService = emailService;
        }

        [HttpGet]
        public async Task<ActionResult<EmailConfigDTO>> Get()
        {
            try
            {
                var config = await _emailService.GetConfigurationAsync();
                return Ok(config);
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al obtener la configuración: {ex.Message}");
            }
        }

        [HttpPost]
        public async Task<ActionResult> Post([FromBody] EmailConfigDTO config)
        {
            try
            {
                await _emailService.SaveConfigurationAsync(config);
                return Ok("Configuración guardada correctamente");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al guardar la configuración: {ex.Message}");
            }
        }

        [HttpPost("test")]
        public async Task<ActionResult> TestEmail(EmailTestDTO emailTest)
        {
            try
            {
                // Guardamos temporalmente la configuración para la prueba
                await _emailService.SaveConfigurationAsync(emailTest.Config);

                // Enviamos el correo de prueba con el archivo adjunto si existe
                if (emailTest.AttachmentContent != null && emailTest.AttachmentName != null)
                {
                    await _emailService.SendTestEmailAsync(emailTest.AttachmentContent, emailTest.AttachmentName);
                }
                else
                {
                    await _emailService.SendTestEmailAsync();
                }

                return Ok("Correo de prueba enviado correctamente");
            }
            catch (Exception ex)
            {
                return BadRequest($"Error al enviar el correo de prueba: {ex.Message}");
            }
        }

        [HttpPost("sendorder")]
        public async Task<ActionResult> SendOrderEmail([FromBody] EmailTestDTO emailTest)
        {
            try
            {
                Console.WriteLine($"[EmailConfigController] 📧 Recibida solicitud de envío de orden por email");
                Console.WriteLine($"[EmailConfigController] 📬 Destinatario: {emailTest.To}");
                Console.WriteLine($"[EmailConfigController] 📋 Asunto: {emailTest.Subject}");
                Console.WriteLine($"[EmailConfigController] 📎 Adjunto: {emailTest.AttachmentName} ({emailTest.AttachmentContent?.Length ?? 0} bytes)");

                // Validar que se haya proporcionado el adjunto
                if (emailTest.AttachmentContent == null || emailTest.AttachmentContent.Length == 0)
                {
                    Console.WriteLine($"[EmailConfigController] ⚠️ ADVERTENCIA: No se proporcionó contenido del adjunto");
                    return BadRequest("No se proporcionó el archivo adjunto");
                }

                // Validar email del destinatario
                if (string.IsNullOrEmpty(emailTest.To))
                {
                    Console.WriteLine($"[EmailConfigController] ⚠️ ADVERTENCIA: No se proporcionó email del destinatario");
                    return BadRequest("No se proporcionó el email del destinatario");
                }

                // Guardamos temporalmente la configuración
                Console.WriteLine($"[EmailConfigController] 💾 Guardando configuración de email...");
                await _emailService.SaveConfigurationAsync(emailTest.Config);
                Console.WriteLine($"[EmailConfigController] ✅ Configuración guardada");

                // Enviamos el correo con la orden de compra adjunta
                Console.WriteLine($"[EmailConfigController] 🚀 Enviando correo...");
                await _emailService.SendEmailAsync(
                    emailTest.To,
                    emailTest.Subject,
                    emailTest.Body,
                    true,
                    emailTest.AttachmentContent,
                    emailTest.AttachmentName
                );

                Console.WriteLine($"[EmailConfigController] ✅ Correo enviado exitosamente");
                return Ok("Correo enviado correctamente");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[EmailConfigController] ❌ ERROR al enviar correo: {ex.Message}");
                Console.WriteLine($"[EmailConfigController] 🔍 Tipo: {ex.GetType().Name}");
                Console.WriteLine($"[EmailConfigController] 🔍 StackTrace: {ex.StackTrace}");
                
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"[EmailConfigController] 🔍 InnerException: {ex.InnerException.Message}");
                }
                
                return BadRequest($"Error al enviar el correo: {ex.Message}");
            }
        }
    }
}
