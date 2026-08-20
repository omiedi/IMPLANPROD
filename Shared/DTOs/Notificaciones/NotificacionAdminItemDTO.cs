using System;
using System.Collections.Generic;

namespace IMPLANPROD.Shared.DTOs.Notificaciones
{
    public class NotificacionAdminItemDTO
    {
        public int Id { get; set; }

        public string Titulo { get; set; } = string.Empty;

        public string? Resumen { get; set; }

        public bool EsParaTodos { get; set; }

        public int CreadorUsuarioId { get; set; }

        public DateTime? AddRecord { get; set; }

        public DateTime? FechaExpiracion { get; set; }

        public bool EsGeneradaPorRegla { get; set; }

        public int? ReglaSqlId { get; set; }

        public List<int> UsuarioIds { get; set; } = new();

        public string DestinatariosResumen { get; set; } = string.Empty;
    }
}
