namespace IMPLANPROD.Shared.DTOs
{
    public sealed class PedidoVentaPendienteDetalleDTO
    {
        public int NroPed { get; set; }
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public decimal CantSolicitada { get; set; }
        public string UMedida { get; set; } = string.Empty;
        public decimal CantEntregada { get; set; }
        public decimal Saldo { get; set; }
        public DateTime? FechaEntrega { get; set; }
    }
}
