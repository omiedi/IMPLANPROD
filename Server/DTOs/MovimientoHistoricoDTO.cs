namespace IMPLANPROD.Server.DTOs
{
    /// <summary>
    /// DTO para exportar movimientos históricos a Excel
    /// </summary>
    public class MovimientoHistoricoDTO
    {
        public DateTime? Fecha { get; set; }
        public string TipoMovimiento { get; set; } = string.Empty;
        public string NumeroDocumento { get; set; } = string.Empty;
        public string CodigoProducto { get; set; } = string.Empty;
        public string DescripcionProducto { get; set; } = string.Empty;
        public decimal Cantidad { get; set; }
        public decimal PrecioUnitario { get; set; }
        public decimal Total { get; set; }
        public string Observaciones { get; set; } = string.Empty;
    }
}
