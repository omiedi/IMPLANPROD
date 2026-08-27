using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para crear una nueva solicitud
    /// </summary>
    public class CrearSolicitudDTO
    {
        [Required(ErrorMessage = "El modulo es obligatorio.")]
        [MaxLength(80)]
        public string Modulo { get; set; } = string.Empty;

        [Required(ErrorMessage = "El tipo de solicitud es obligatorio.")]
        [MaxLength(50)]
        public string TipoSolicitud { get; set; } = string.Empty;

        [Required(ErrorMessage = "El nivel es obligatorio.")]
        [MaxLength(50)]
        public string Nivel { get; set; } = string.Empty;

        [Required(ErrorMessage = "El titulo es obligatorio.")]
        [MaxLength(255)]
        public string Titulo { get; set; } = string.Empty;

        [Required(ErrorMessage = "El detalle es obligatorio.")]
        public string Detalle { get; set; } = string.Empty;

        /// <summary>
        /// Solo para uso administrativo. Si es null, se asocia al usuario activo.
        /// </summary>
        public int? UsuarioSolicitante { get; set; }
    }
}
