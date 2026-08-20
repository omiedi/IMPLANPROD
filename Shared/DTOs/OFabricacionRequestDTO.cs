namespace IMPLANPROD.Shared.DTOs
{
    public class OFabricacionRequestDTO
    {
        public int? IdFlexoftUsuario { get; set; }
        public List<int>? Depositos { get; set; } = new List<int>();

        public int Page { get; set; } = 1;

        public int RecordsNumber { get; set; } = 10;

        public string? Filter { get; set; }
    }
}
