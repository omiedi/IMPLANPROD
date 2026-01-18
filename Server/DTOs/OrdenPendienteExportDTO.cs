namespace IMPLANPROD.Server.DTOs
{
    /// <summary>
    /// DTO para exportar órdenes pendientes a Excel
    /// </summary>
    public class OrdenPendienteExportDTO
    {
        public int NumeOcom { get; set; }
        public int NuOrdItem { get; set; }
        public string CodExte { get; set; } = string.Empty;
        public string Descritem { get; set; } = string.Empty;
        public decimal CantOcom { get; set; }
        public decimal CantRec { get; set; }
        public decimal Pendiente { get; set; }
        public string FemiOcom { get; set; } = string.Empty;
    }
}
