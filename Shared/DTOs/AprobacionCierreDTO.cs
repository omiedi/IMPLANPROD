namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para procesar la aprobación de una orden de fabricación con cierre existente
    /// </summary>
    public class AprobacionCierreDTO
    {
        /// <summary>
        /// ID de la suborden de fabricación (NORFA$)
        /// </summary>
        public string IdSubOr { get; set; } = string.Empty;

        /// <summary>
        /// ID de texto del cierre (IDTX$)
        /// </summary>
        public string IdText { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad aprobada por el usuario
        /// </summary>
        public decimal CantidadAprobada { get; set; }

        /// <summary>
        /// Depósito de entrada (DEPOEN%)
        /// </summary>
        public int DepositoEntrada { get; set; }

        /// <summary>
        /// Depósito de consumo (DP2%)
        /// </summary>
        public int DepositoConsumo { get; set; }

        /// <summary>
        /// Observaciones adicionales
        /// </summary>
        public string? Observaciones { get; set; }
    }
}
