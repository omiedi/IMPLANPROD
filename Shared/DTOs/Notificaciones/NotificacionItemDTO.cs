using System;

namespace IMPLANPROD.Shared.DTOs.Notificaciones
{
    public class NotificacionItemDTO
    {
        public int Id { get; set; }

        public string Titulo { get; set; } = string.Empty;

        public string? Resumen { get; set; }

        public DateTime? AddRecord { get; set; }

        public bool Leida { get; set; }

        public bool TieneDetalle { get; set; }

        public bool EsParaTodos { get; set; }

        public int CreadorUsuarioId { get; set; }

        public bool EsGeneradaPorRegla { get; set; }

        public int? ReglaSqlId { get; set; }

        public string DestinatariosResumen { get; set; } = string.Empty;
    }
}
