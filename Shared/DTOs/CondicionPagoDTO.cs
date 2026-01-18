namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para condición de pago
    /// Mapea campos de la tabla CONDPAG
    /// </summary>
    public class CondicionPagoDTO
    {
        /// <summary>
        /// Número de condición de pago
        /// </summary>
        public short NumCondPag { get; set; }

        /// <summary>
        /// Descripción de la condición de pago
        /// </summary>
        public string? DescCondPag { get; set; }
    }
}
