using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs.Notificaciones
{
    public class NotificacionAdminUpsertDTO
    {
        public int Id { get; set; }

        [Required]
        [MaxLength(200)]
        public string Titulo { get; set; } = string.Empty;

        [MaxLength(500)]
        public string? Resumen { get; set; }

        public string? DetalleHtml { get; set; }

        public bool EsParaTodos { get; set; }

        public List<int> UsuarioIds { get; set; } = new();

        public DateTime? FechaExpiracion { get; set; }
    }
}
