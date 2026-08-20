namespace IMPLANPROD.Shared.DTOs
{
    public sealed class StockCompraPendienteDTO
    {
        public int NumeOcom { get; set; }
        public int NproOcom { get; set; }
        public string Proveedor { get; set; } = string.Empty;
        public decimal CantPendiente { get; set; }
        public DateTime? FechaEntregaSolicitada { get; set; }
    }
}
