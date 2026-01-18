namespace IMPLANPROD.Shared.DTOs
{
    public class MasterDTO
    {
        public int Codint { get; set; } = 0;        // Para el Id/codint
        public string Codigo { get; set; } = string.Empty; // Para el código externo
        public string Descripcion { get; set; } = string.Empty;
        public int? Umedida { get; set; }  // Número de unidad de medida
    }
}
