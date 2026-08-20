namespace IMPLANPROD.Shared.DTOs
{
    public class TrazabilidadOFSerieDetalleDTO
    {
        public int CodInte { get; set; }
        public string CodExte { get; set; } = string.Empty;
        public string Descrip { get; set; } = string.Empty;
        public string Unidad { get; set; } = string.Empty;
        public decimal UsoBruto { get; set; }
        public decimal CantidadAsignada { get; set; }
        public int ControserieId { get; set; }
        public int Orden { get; set; }
        public int NivelIndentado { get; set; }
        public Dictionary<int, decimal> StockPorDeposito { get; set; } = new();
    }

    public class TrazabilidadOFSerieResultadoDTO
    {
        public bool TieneDatosControserie { get; set; }
        public List<TrazabilidadOFSerieDetalleDTO> Items { get; set; } = new();
    }
}
