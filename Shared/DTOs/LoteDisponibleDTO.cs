namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para representar un lote con saldo disponible
    /// </summary>
    public class LoteDisponibleDTO
    {
        public string Lote { get; set; } = string.Empty;
        public decimal Saldo { get; set; }
    }
}
