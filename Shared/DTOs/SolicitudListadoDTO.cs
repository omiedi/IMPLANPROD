using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para un item del listado de solicitudes
    /// </summary>
    public class SolicitudListadoDTO
    {
        public int Id { get; set; }
        public DateTime FechaIngreso { get; set; }
        public string Titulo { get; set; } = string.Empty;
        public string Modulo { get; set; } = string.Empty;
        public string TipoSolicitud { get; set; } = string.Empty;
        public string Nivel { get; set; } = string.Empty;
        public string Estado { get; set; } = string.Empty;
        public DateTime FechaUltimaActualizacion { get; set; }
        public int? UsuarioUltimaActualizacion { get; set; }
        public int UsuarioSolicitante { get; set; }
        public string? NombreSolicitante { get; set; }
    }
}
