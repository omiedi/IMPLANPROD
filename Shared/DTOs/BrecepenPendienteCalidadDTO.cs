namespace IMPLANPROD.Shared.DTOs
{
    // DTO para la grilla de "Aprobación de Calidad".
    // Devuelve recepciones pendientes agrupadas desde BRECEPEN + datos descriptivos desde MASTER.
    public class BrecepenPendienteCalidadDTO
    {
        // Lote de la recepción: junto con CodInte identifica el grupo pendiente.
        public string? Idenlote { get; set; }
        public string? Doprilar { get; set; }
        public int CodInte { get; set; }
        public string? Codigo { get; set; }
        public string? Descripcion { get; set; }
        public decimal CantidadPendiente { get; set; }
        // Número de recepción para validación de tolerancia
        public int? NumeroRecepcion { get; set; }
        // Fecha de recepción para validación de tolerancia
        public DateTime? FechaRecepcion { get; set; }
    }
}
