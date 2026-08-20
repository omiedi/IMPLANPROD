namespace IMPLANPROD.Shared.DTOs
{
    // Request para aprobar (consumir) una cantidad pendiente desde BRECEPEN.
    // Se identifica el grupo por IDENLOTE + COD_INTE, y se genera un registro de aprobación (CODIMOVI='AP').
    public class BrecepenAprobarRequestDTO
    {
        // Lote de la recepción a aprobar (identifica el grupo junto con CodInte).
        public string? Idenlote { get; set; }
        public string? Doprilar { get; set; }
        public int CodInte { get; set; }
        public decimal CantidadAprobar { get; set; }
    }
}
