namespace IMPLANPROD.Shared.DTOs
{
    public class ProductoConSaldoDTO
    {
        public int Cod_inte { get; set; }
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public string UnidadMedida { get; set; } = string.Empty;
        public decimal Saldo { get; set; }
        public string? IdenLote { get; set; }
    }
}
