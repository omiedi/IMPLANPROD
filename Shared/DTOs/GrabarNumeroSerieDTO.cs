namespace IMPLANPROD.Shared.DTOs
{
    public class GrabarNumeroSerieRequestDTO
    {
        public string IdSubOr { get; set; } = string.Empty;
        public string NumeroSerie { get; set; } = string.Empty;
        public int CodInte { get; set; }
        public int? DepositoMovi { get; set; }
        public List<TrazabilidadOFSerieDetalleDTO> Items { get; set; } = new();
    }

    public class GrabarNumeroSerieResultadoDTO
    {
        public bool Ok { get; set; }
        public string Mensaje { get; set; } = string.Empty;
        public bool YaExistia { get; set; }
        public int RegistrosGrabados { get; set; }
        public int DepositoMovi { get; set; }
    }
}
