namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Resultado de la validación de tolerancia para aprobar recepciones
    /// </summary>
    public class ValidacionToleranciaDTO
    {
        /// <summary>
        /// Indica si el usuario puede aprobar la recepción
        /// </summary>
        public bool PuedeAprobar { get; set; }

        /// <summary>
        /// Mensaje de error si no puede aprobar (null si puede aprobar)
        /// </summary>
        public string? Mensaje { get; set; }
    }
}
