namespace IMPLANPROD.Shared.DTOs
{
    public sealed class StockFabricacionPendienteDTO
    {
        public int NumeOrfa { get; set; }
        public string Referencia { get; set; } = string.Empty;
        public decimal CantPendiente { get; set; }
        public DateTime? FechaEntregaSolicitada { get; set; }
        public DateTime? FechaProyectadaEntrega { get; set; }
        public DateTime? FechaRecalculo { get; set; }
    }
}
