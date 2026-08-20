using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs.Notificaciones
{
    public class OcultarNotificacionRequestDTO
    {
        [Required]
        public int NotificacionId { get; set; }
    }
}
