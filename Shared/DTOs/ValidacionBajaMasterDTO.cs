namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para la respuesta de validación de baja de producto Master
    /// </summary>
    public class ValidacionBajaMasterDTO
    {
        /// <summary>
        /// Indica si el producto puede ser dado de baja
        /// </summary>
        public bool PuedeDarDeBaja { get; set; }

        /// <summary>
        /// Lista de códigos de productos padres donde se usa este producto como componente
        /// Solo se llena si PuedeDarDeBaja es false
        /// </summary>
        public List<string> CodigosPadres { get; set; } = new();
    }
}
