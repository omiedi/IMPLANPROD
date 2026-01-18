namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para listar reparaciones en el selector de Material Suministrado
    /// </summary>
    public class ReparacionListaDTO
    {
        public int NumRepa { get; set; }
        public string? Referen { get; set; }
        public int? Nume_Cli { get; set; }
        public string? Raso_Cli { get; set; }
        public int? NuOrdRep { get; set; }
        public string? CodRepa { get; set; }
        public string? RefOrdRep { get; set; }
    }
}
