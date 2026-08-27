using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para filtrar y paginar el listado de solicitudes
    /// </summary>
    public class SolicitudFiltroDTO : PaginationDTO
    {
        public DateTime? FechaDesde { get; set; }
        public DateTime? FechaHasta { get; set; }
        public string? Modulo { get; set; }
        public string? TipoSolicitud { get; set; }
        public string? Nivel { get; set; }
        public string? Estado { get; set; }
        public int? Numero { get; set; }
        public int? UsuarioFiltro { get; set; }
        public bool VerTodos { get; set; }
    }
}
