namespace IMPLANPROD.Shared.DTOs
{
    public class DepositosPosiblesConSaldoDTO
    {
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public decimal SaldoStock { get; set; }
        public int Deposito { get; set; }
    }
}
