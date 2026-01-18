namespace IMPLANPROD.Shared.DTOs
{
    public class MovimientoHistoricoDTO
    {
        public int Id { get; set; }
        public DateTime Fecha { get; set; }
        public string? TipoMovimiento { get; set; }
        public string? DocumentoNumero { get; set; }
        public string? Transaccion { get; set; }
        public int? Deposito { get; set; }
        public string? IdBulto { get; set; }
        public string? Referencia { get; set; }
        public decimal? Entrada { get; set; }
        public decimal? Salida { get; set; }
        public decimal Saldo { get; set; }
        public int CodInte { get; set; }
    }

    public class FiltroMovimientosDTO
    {
        public int CodigoRepuesto { get; set; }
        public int? Deposito { get; set; }
        public DateTime? FechaDesde { get; set; }
        public DateTime? FechaHasta { get; set; }
    }
}
