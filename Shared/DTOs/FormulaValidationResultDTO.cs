namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Resultado de la validación de modificación de fórmula
    /// </summary>
    public class FormulaValidationResultDTO
    {
        /// <summary>
        /// Indica si se debe bloquear la modificación de la fórmula
        /// </summary>
        public bool DebeBloquear { get; set; }

        /// <summary>
        /// Mensaje explicativo del motivo del bloqueo (si aplica)
        /// </summary>
        public string? Mensaje { get; set; }

        /// <summary>
        /// Código del producto padre que causa el bloqueo (si aplica)
        /// </summary>
        public string? CodigoPadre { get; set; }

        /// <summary>
        /// Descripción del producto padre que causa el bloqueo (si aplica)
        /// </summary>
        public string? DescripcionPadre { get; set; }

        /// <summary>
        /// Fecha del último pedido del producto padre (si aplica)
        /// </summary>
        public DateTime? FechaUltimoPedido { get; set; }

        /// <summary>
        /// Días desde el último pedido (si aplica)
        /// </summary>
        public int? DiasDesdeUltimoPedido { get; set; }

        /// <summary>
        /// Número del pedido que origina el bloqueo (PEDDETA.NroPed)
        /// </summary>
        public int? NroPedido { get; set; }

        /// <summary>
        /// Número de cliente del pedido que origina el bloqueo (PEDDETA.NroClie)
        /// </summary>
        public int? NroCliente { get; set; }

        /// <summary>
        /// Razón social del cliente del pedido que origina el bloqueo (PEDDETA.RaSoClie)
        /// </summary>
        public string? RazonSocial { get; set; }

        /// <summary>
        /// Cadena de estructura de fórmula desde el padre vencido hasta el código activo
        /// Formato: "PADRE > HIJO1 > HIJO2 > HIJO3 (CÓDIGO ACTIVO)"
        /// </summary>
        public string? EstructuraFormula { get; set; }
    }
}
