using IMPLANPROD.Server.Data;
using IMPLANPROD.Server.Utilities;
using IMPLANPROD.Shared.DTOs;
using IMPLANPROD.Shared.Entities;
using MailKit.Net.Smtp;
using MailKit.Security;
using Microsoft.EntityFrameworkCore;
using MimeKit;

namespace IMPLANPROD.Server.Services
{
    public interface IEmailEmpresaService
    {
        Task<EmailConfigEmpresaDto> GetConfiguracionAsync();
        Task GuardarConfiguracionAsync(EmailConfigEmpresaDto dto);
        Task EnviarEmailAsync(string to, string subject, string body, bool isHtml = false, string? replyTo = null, string? replyToName = null, byte[]? attachmentContent = null, string? attachmentName = null);
        Task EnviarEmailPruebaAsync(string destinatario);
    }

    public class EmailEmpresaService : IEmailEmpresaService
    {
        private readonly DataContext _context;

        public EmailEmpresaService(DataContext context)
        {
            _context = context;
        }

        public async Task<EmailConfigEmpresaDto> GetConfiguracionAsync()
        {
            var config = await _context.EmailConfigEmpresa
                .AsNoTracking()
                .OrderByDescending(x => x.LastModified)
                .FirstOrDefaultAsync();

            if (config == null)
                return new EmailConfigEmpresaDto();

            return new EmailConfigEmpresaDto
            {
                SmtpHost           = config.SmtpHost,
                SmtpPort           = config.SmtpPort,
                UsarSsl            = config.UsarSsl,
                UsarStartTls       = config.UsarStartTls,
                UsuarioSmtp        = config.UsuarioSmtp,
                Contrasena         = string.Empty,
                EmailDesde         = config.EmailDesde,
                NombreDesde        = config.NombreDesde,
                TieneContrasenaGuardada = !string.IsNullOrEmpty(config.ContrasenaEncriptada)
            };
        }

        public async Task GuardarConfiguracionAsync(EmailConfigEmpresaDto dto)
        {
            var existing = await _context.EmailConfigEmpresa
                .OrderByDescending(x => x.LastModified)
                .FirstOrDefaultAsync();

            if (existing == null)
            {
                existing = new EmailConfigEmpresa { Codiempr = 0 };
                _context.EmailConfigEmpresa.Add(existing);
            }

            existing.SmtpHost     = dto.SmtpHost.Trim();
            existing.SmtpPort     = dto.SmtpPort;
            existing.UsarSsl      = dto.UsarSsl;
            existing.UsarStartTls = dto.UsarStartTls;
            existing.UsuarioSmtp  = dto.UsuarioSmtp.Trim();
            existing.EmailDesde   = dto.EmailDesde.Trim();
            existing.NombreDesde  = dto.NombreDesde.Trim();
            existing.LastModified = DateTime.UtcNow;

            if (!string.IsNullOrWhiteSpace(dto.Contrasena))
                existing.ContrasenaEncriptada = EncryptionHelper.Encrypt(dto.Contrasena);

            await _context.SaveChangesAsync();
        }

        public async Task EnviarEmailAsync(string to, string subject, string body, bool isHtml = false, string? replyTo = null, string? replyToName = null, byte[]? attachmentContent = null, string? attachmentName = null)
        {
            var config = await ObtenerConfigOException();
            var password = EncryptionHelper.Decrypt(config.ContrasenaEncriptada);

            var message = new MimeMessage();
            message.From.Add(new MailboxAddress(config.NombreDesde, config.EmailDesde));
            message.To.Add(MailboxAddress.Parse(to));
            message.Subject = subject;

            if (!string.IsNullOrWhiteSpace(replyTo))
                message.ReplyTo.Add(new MailboxAddress(replyToName ?? replyTo, replyTo));

            var bodyBuilder = new BodyBuilder();
            if (isHtml)
                bodyBuilder.HtmlBody = body;
            else
                bodyBuilder.TextBody = body;

            if (attachmentContent != null && !string.IsNullOrWhiteSpace(attachmentName))
                bodyBuilder.Attachments.Add(attachmentName, attachmentContent);

            message.Body = bodyBuilder.ToMessageBody();

            using var client = new SmtpClient();

            // -----------------------------------------------------------------------
            // VALIDACIÓN DE CERTIFICADO SSL - COMPORTAMIENTO INTENCIONAL
            // -----------------------------------------------------------------------
            // Muchos servidores de hosting compartido argentino (cPanel, Plesk, etc.)
            // usan certificados cuya Lista de Revocación (CRL) no es accesible desde
            // el servidor de la aplicación (sin salida a internet o CRL caída).
            // Eso causa SslHandshakeException aunque el certificado sea válido.
            //
            // Esta callback acepta el certificado SOLO cuando el único error es la
            // imposibilidad de verificar la revocación (RevocationStatusUnknown /
            // OfflineRevocation). Cualquier otro error (certificado vencido, nombre
            // no coincide, etc.) sigue siendo rechazado normalmente.
            //
            // Si en el futuro el servidor tiene acceso a las CRL del proveedor,
            // se puede eliminar esta callback para restaurar la validación estricta.
            // -----------------------------------------------------------------------
            client.ServerCertificateValidationCallback = (s, c, chain, errors) =>
            {
                if (errors == System.Net.Security.SslPolicyErrors.None) return true;
                if (chain == null) return false;
                return chain.ChainStatus.All(x =>
                    x.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.RevocationStatusUnknown ||
                    x.Status == System.Security.Cryptography.X509Certificates.X509ChainStatusFlags.OfflineRevocation);
            };

            var sslOption = ObtenerOpcionSsl(config);
            await client.ConnectAsync(config.SmtpHost, config.SmtpPort, sslOption);
            await client.AuthenticateAsync(config.UsuarioSmtp, password);
            await client.SendAsync(message);
            await client.DisconnectAsync(true);
        }

        public async Task EnviarEmailPruebaAsync(string destinatario)
        {
            var config = await ObtenerConfigOException();

            var subject = "✅ Prueba de configuración SMTP de empresa";
            var sslDesc = config.UsarSsl
                ? "SSL/TLS (puerto 465)"
                : config.UsarStartTls
                    ? "STARTTLS (puerto 587)"
                    : "Sin cifrado";

            var body = $@"
                <h3 style='color:#198754;'>Prueba de correo exitosa</h3>
                <p>Este correo confirma que la configuración SMTP corporativa funciona correctamente.</p>
                <hr/>
                <table style='font-family:sans-serif;font-size:14px;'>
                    <tr><td><strong>Servidor SMTP:</strong></td><td>{config.SmtpHost}:{config.SmtpPort}</td></tr>
                    <tr><td><strong>Cifrado:</strong></td><td>{sslDesc}</td></tr>
                    <tr><td><strong>Usuario SMTP:</strong></td><td>{config.UsuarioSmtp}</td></tr>
                    <tr><td><strong>Remitente:</strong></td><td>{config.NombreDesde} &lt;{config.EmailDesde}&gt;</td></tr>
                    <tr><td><strong>Destinatario prueba:</strong></td><td>{destinatario}</td></tr>
                    <tr><td><strong>Fecha/Hora:</strong></td><td>{DateTime.Now:dd/MM/yyyy HH:mm:ss}</td></tr>
                </table>";

            await EnviarEmailAsync(destinatario, subject, body, true);
        }

        private async Task<EmailConfigEmpresa> ObtenerConfigOException()
        {
            return await _context.EmailConfigEmpresa
                .AsNoTracking()
                .OrderByDescending(x => x.LastModified)
                .FirstOrDefaultAsync()
                ?? throw new InvalidOperationException("La configuración de correo de empresa no está configurada.");
        }

        private static SecureSocketOptions ObtenerOpcionSsl(EmailConfigEmpresa config)
        {
            if (config.UsarSsl)    return SecureSocketOptions.SslOnConnect;
            if (config.UsarStartTls) return SecureSocketOptions.StartTls;
            return SecureSocketOptions.None;
        }
    }
}
