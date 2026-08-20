using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    public class EmailConfigEmpresaDto
    {
        [Required(ErrorMessage = "El servidor SMTP es obligatorio")]
        [MaxLength(150, ErrorMessage = "El servidor no puede superar 150 caracteres")]
        public string SmtpHost { get; set; } = string.Empty;

        [Required(ErrorMessage = "El puerto es obligatorio")]
        [Range(1, 65535, ErrorMessage = "El puerto debe estar entre 1 y 65535")]
        public int SmtpPort { get; set; } = 587;

        public bool UsarSsl { get; set; } = false;

        public bool UsarStartTls { get; set; } = true;

        [Required(ErrorMessage = "El usuario SMTP es obligatorio")]
        [MaxLength(200)]
        public string UsuarioSmtp { get; set; } = string.Empty;

        public string Contrasena { get; set; } = string.Empty;

        [Required(ErrorMessage = "El correo de origen es obligatorio")]
        [EmailAddress(ErrorMessage = "El formato del correo no es válido")]
        [MaxLength(200)]
        public string EmailDesde { get; set; } = string.Empty;

        [Required(ErrorMessage = "El nombre para mostrar es obligatorio")]
        [MaxLength(200)]
        public string NombreDesde { get; set; } = string.Empty;

        [EmailAddress(ErrorMessage = "El correo de prueba no es válido")]
        public string DestinatarioPrueba { get; set; } = string.Empty;

        public bool TieneContrasenaGuardada { get; set; } = false;
    }
}
