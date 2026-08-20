namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para imprimir la especificación de un producto
    /// </summary>
    public class EspecificacionPrintDTO
    {
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public string Especificacion { get; set; } = string.Empty;
    }
}
