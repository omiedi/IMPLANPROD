namespace IMPLANPROD.Shared.DTOs
{
    public sealed class StockPedidoPendienteDTO
    {
        public int NroPed { get; set; }
        public int NroClie { get; set; }
        public string Cliente { get; set; } = string.Empty;
        public decimal CantPendiente { get; set; }
        public decimal SalPreparado { get; set; }
        public DateTime? FechaSolEnt { get; set; }
    }
}
