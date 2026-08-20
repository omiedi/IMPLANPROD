namespace IMPLANPROD.Shared.DTOs
{
    public sealed class StockReservaPendienteDTO
    {
        public string IdSubOr { get; set; } = string.Empty;
        public string IdenLote { get; set; } = string.Empty;
        public DateTime? FechaReserva { get; set; }
        public DateTime? FechaEntregaSolicitada { get; set; }
        public string CodExteDestino { get; set; } = string.Empty;
        public string DescripDestino { get; set; } = string.Empty;
        public decimal CantReservada { get; set; }
        public decimal CantConsumida { get; set; }
        public decimal CantPendiente { get; set; }
    }
}
