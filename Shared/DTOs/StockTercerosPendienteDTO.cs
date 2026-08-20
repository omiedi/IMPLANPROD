namespace IMPLANPROD.Shared.DTOs
{
    public sealed class StockTercerosPendienteDTO
    {
        public string Ocom { get; set; } = string.Empty; // OC-000001
        public int NproOcom { get; set; }
        public string Proveedor { get; set; } = string.Empty;
        public decimal Entregado { get; set; }
        public decimal Rendido { get; set; }
        public decimal Devuelto { get; set; }
        public decimal Asignado { get; set; }
        public decimal CapenRendir { get; set; }
        public decimal CapenEntregar { get; set; }
        public short? StatOcom { get; set; }
        public DateTime? FechaEmisionOcom { get; set; }
    }
}
