using System;
using System.Collections.Generic;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO con el detalle completo de una solicitud, incluyendo su historial
    /// </summary>
    public class SolicitudDetalleDTO
    {
        public int Id { get; set; }
        public DateTime FechaIngreso { get; set; }
        public int UsuarioSolicitante { get; set; }
        public string? NombreSolicitante { get; set; }
        public string Modulo { get; set; } = string.Empty;
        public string TipoSolicitud { get; set; } = string.Empty;
        public string Nivel { get; set; } = string.Empty;
        public string Titulo { get; set; } = string.Empty;
        public string Detalle { get; set; } = string.Empty;
        public string Estado { get; set; } = string.Empty;
        public DateTime FechaUltimaActualizacion { get; set; }
        public int? UsuarioUltimaActualizacion { get; set; }
        public DateTime? FechaCierre { get; set; }
        public int? UsuarioCierre { get; set; }
        public List<SolicitudAvanceDTO> Avances { get; set; } = new();
    }
}
