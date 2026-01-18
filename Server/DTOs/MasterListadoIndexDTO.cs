namespace IMPLANPROD.Server.DTOs
{
    /// <summary>
    /// DTO para exportar listado de productos master a Excel
    /// </summary>
    public class MasterListadoIndexDTO
    {
        public int Id { get; set; }
        public int Codint { get; set; }
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public string Umedida { get; set; } = string.Empty;
        public string TipoMaterialDescripcion { get; set; } = string.Empty;
        public decimal Stock { get; set; }
        public decimal StockMinimo { get; set; }
        public decimal PrecioUnitario { get; set; }
        public string Categoria { get; set; } = string.Empty;
        public string Estado { get; set; } = string.Empty;
        public DateTime? FechaCreacion { get; set; }
        public DateTime? FechaActualizacion { get; set; }
    }
}
