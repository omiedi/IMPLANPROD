using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs.Notificaciones
{
    public class MarcarLeidaRequestDTO
    {
        [Required]
        public int NotificacionId { get; set; }
    }
}
