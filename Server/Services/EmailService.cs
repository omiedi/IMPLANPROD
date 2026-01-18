using IMPLANPROD.Server.Data;
using IMPLANPROD.Shared.DTOs;
using Microsoft.EntityFrameworkCore;
using System.Net;
using System.Net.Mail;
using IMPLANPROD.Shared.Entities;
using IMPLANPROD.Server.Utilities;


namespace IMPLANPROD.Server.Services
{
    public interface IEmailService
    {
        Task SendEmailAsync(string to, string subject, string body, bool isHtml = false, byte[]? attachment = null, string? attachmentName = null);
        Task SendTestEmailAsync();
        Task SendTestEmailAsync(byte[] attachment, string attachmentName);
        Task<EmailConfigDTO> GetConfigurationAsync();
        Task SaveConfigurationAsync(EmailConfigDTO config);
    }

    public class EmailService : IEmailService
    {
        private readonly DataContext _context;
        private EmailConfigDTO? _currentConfig;

        public EmailService(DataContext context)
        {
            _context = context;
        }

        private async Task LoadConfigurationAsync()
        {
            if (_currentConfig != null) return;

            var dbConfig = await _context.EmailConfiguration.OrderByDescending(x => x.LastModified).FirstOrDefaultAsync();
            if (dbConfig != null)
            {
                _currentConfig = new EmailConfigDTO
                {
                    EmailAddress = dbConfig.EmailAddress,
                    AppPassword = EncryptionHelper.Decrypt(dbConfig.EncryptedPassword),
                    DisplayName = dbConfig.DisplayName
                };
            }
        }

        public async Task SendEmailAsync(string to, string subject, string body, bool isHtml = false, byte[]? attachment = null, string? attachmentName = null)
        {
            Console.WriteLine($"[EmailService] 📧 Iniciando envío de correo a: {to}");
            Console.WriteLine($"[EmailService] 📋 Asunto: {subject}");
            
            await LoadConfigurationAsync();
            if (_currentConfig == null)
            {
                Console.WriteLine($"[EmailService] ❌ ERROR: La configuración de correo no está establecida");
                throw new InvalidOperationException("La configuración de correo no está establecida");
            }

            Console.WriteLine($"[EmailService] ⚙️ Configuración cargada - Email: {_currentConfig.EmailAddress}");
            Console.WriteLine($"[EmailService] 📎 Adjunto: {(attachment != null ? $"Sí ({attachmentName}, {attachment.Length} bytes)" : "No")}");

            var smtpClient = new SmtpClient("smtp.gmail.com")
            {
                Port = 587,
                Credentials = new NetworkCredential(_currentConfig.EmailAddress, _currentConfig.AppPassword),
                EnableSsl = true,
            };

            Console.WriteLine($"[EmailService] 🔧 Cliente SMTP configurado - Servidor: smtp.gmail.com:587, SSL: true");

            var mailMessage = new MailMessage
            {
                From = new MailAddress(_currentConfig.EmailAddress, _currentConfig.DisplayName),
                Subject = subject,
                Body = body,
                IsBodyHtml = isHtml
            };

            mailMessage.To.Add(to);
            Console.WriteLine($"[EmailService] 📬 Destinatario agregado: {to}");

            if (attachment != null && attachmentName != null)
            {
                var attachmentStream = new MemoryStream(attachment);
                mailMessage.Attachments.Add(new Attachment(attachmentStream, attachmentName));
                Console.WriteLine($"[EmailService] 📎 Adjunto agregado: {attachmentName}");
            }

            try
            {
                Console.WriteLine($"[EmailService] 🚀 Enviando correo...");
                await smtpClient.SendMailAsync(mailMessage);
                Console.WriteLine($"[EmailService] ✅ Correo enviado exitosamente a {to}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[EmailService] ❌ ERROR al enviar correo: {ex.Message}");
                Console.WriteLine($"[EmailService] 📋 Tipo de excepción: {ex.GetType().Name}");
                Console.WriteLine($"[EmailService] 🔍 StackTrace: {ex.StackTrace}");
                
                if (ex.InnerException != null)
                {
                    Console.WriteLine($"[EmailService] 🔍 InnerException: {ex.InnerException.Message}");
                }
                
                throw new Exception($"Error al enviar el correo: {ex.Message}", ex);
            }
            finally
            {
                mailMessage.Dispose();
                Console.WriteLine($"[EmailService] 🧹 Recursos liberados");
            }
        }

        public async Task SendTestEmailAsync()
        {
            await LoadConfigurationAsync();
            if (_currentConfig == null)
                throw new InvalidOperationException("La configuración de correo no está establecida");

            var subject = "Prueba de Configuración de Correo";
            var body = @"
                <h3>Prueba de Configuración de Correo Exitosa</h3>
                <p>Este es un correo de prueba para verificar la configuración de su cuenta de correo en el sistema de Mantenimiento.</p>
                <p>Si recibió este correo, la configuración es correcta.</p>
                <br>
                <p><strong>Configuración utilizada:</strong></p>
                <ul>
                    <li>Correo: " + _currentConfig.EmailAddress + @"</li>
                    <li>Nombre para mostrar: " + _currentConfig.DisplayName + @"</li>
                </ul>";

            await SendEmailAsync(_currentConfig.EmailAddress, subject, body, true);
        }

        public async Task SendTestEmailAsync(byte[] attachment, string attachmentName)
        {
            await LoadConfigurationAsync();
            if (_currentConfig == null)
                throw new InvalidOperationException("La configuración de correo no está establecida");

            var subject = "Prueba de Configuración de Correo con Archivo Adjunto";
            var body = @"
                <h3>Prueba de Configuración de Correo Exitosa</h3>
                <p>Este es un correo de prueba para verificar la configuración de su cuenta de correo en el sistema de Mantenimiento.</p>
                <p>Si recibió este correo, la configuración es correcta.</p>
                <p>Se ha adjuntado el archivo: " + attachmentName + @"</p>
                <br>
                <p><strong>Configuración utilizada:</strong></p>
                <ul>
                    <li>Correo: " + _currentConfig.EmailAddress + @"</li>
                    <li>Nombre para mostrar: " + _currentConfig.DisplayName + @"</li>
                </ul>";

            await SendEmailAsync(_currentConfig.EmailAddress, subject, body, true, attachment, attachmentName);
        }

        public async Task<EmailConfigDTO> GetConfigurationAsync()
        {
            await LoadConfigurationAsync();
            return _currentConfig ?? new EmailConfigDTO();
        }

        public async Task SaveConfigurationAsync(EmailConfigDTO config)
        {
            var emailConfig = new EmailConfiguration
            {
                EmailAddress = config.EmailAddress,
                EncryptedPassword = EncryptionHelper.Encrypt(config.AppPassword),
                DisplayName = config.DisplayName,
                LastModified = DateTime.UtcNow
            }; 

            _context.EmailConfiguration.Add(emailConfig);
            await _context.SaveChangesAsync();

            _currentConfig = config;
        }
    }
}

