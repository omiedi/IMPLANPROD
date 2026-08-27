using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para un avance/historial de una solicitud
    /// </summary>
    public class SolicitudAvanceDTO
    {
        public int Id { get; set; }
        public int SolicitudId { get; set; }
        public DateTime Fecha { get; set; }
        public string Estado { get; set; } = string.Empty;
        public string Comentario { get; set; } = string.Empty;
        public int Usuario { get; set; }
        public string? NombreUsuario { get; set; }
    }
}
