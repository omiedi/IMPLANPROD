namespace IMPLANPROD.Server.DTOs
{
    /// <summary>
    /// DTO para exportar detalle de orden a Excel
    /// </summary>
    public class DetalleOrdenExportDTO
    {
        public int NuOrdItem { get; set; }
        public string CodExte { get; set; } = string.Empty;
        public string Descritem { get; set; } = string.Empty;
        public decimal CantOcom { get; set; }
        public decimal CantRec { get; set; }
        public decimal Pendiente { get; set; }
        public decimal PreUnit { get; set; }
        public decimal INetItem { get; set; }
    }
}
