namespace IMPLANPROD.Shared.DTOs
{
    public sealed class MasterStockListItemDTO
    {
        public int Codint { get; set; }
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public string Umedida { get; set; } = string.Empty;
    }
}
