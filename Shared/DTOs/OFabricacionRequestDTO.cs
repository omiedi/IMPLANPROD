namespace IMPLANPROD.Shared.DTOs
{
    public class OFabricacionRequestDTO
    {
        public int IdFlexoftUsuario { get; set; }
        public List<int> Depositos { get; set; } = new List<int>();
    }
}
