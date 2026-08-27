using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para agregar un avance a una solicitud
    /// </summary>
    public class CrearSolicitudAvanceDTO
    {
        public int SolicitudId { get; set; }

        [Required(ErrorMessage = "El estado es obligatorio.")]
        [MaxLength(50)]
        public string Estado { get; set; } = string.Empty;

        [Required(ErrorMessage = "El comentario es obligatorio.")]
        public string Comentario { get; set; } = string.Empty;
    }
}
