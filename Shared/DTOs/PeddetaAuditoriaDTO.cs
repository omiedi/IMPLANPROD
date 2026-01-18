using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para mostrar los datos de auditoría de PEDDETA parseados desde ValorNuevo
    /// </summary>
    public class PeddetaAuditoriaDTO
    {
        public int IdAuditoria { get; set; }
        public int NroPed { get; set; }
        public string Accion { get; set; } = string.Empty;
        public DateTime FechaCambio { get; set; }
        public int? NumUsuario { get; set; }
        
        // Campos parseados desde ValorNuevo
        public decimal? CanPed { get; set; }
        public decimal? PreUnid { get; set; }
        public decimal? CantEnt { get; set; }
        public string? NroOcom { get; set; }
        public int? Vended { get; set; }
        public int? NuOrItem { get; set; }
        public int? CodiInt { get; set; }
        public string? CodiExt { get; set; }
        public string? Descrip { get; set; }
        public int? MoneEmis { get; set; }
        public decimal? TipoCam { get; set; }
    }
}
