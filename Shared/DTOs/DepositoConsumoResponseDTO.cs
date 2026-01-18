namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para la respuesta del depósito de consumo
    /// Incluye información sobre si el depósito está activo y permite manejar casos especiales
    /// </summary>
    public class DepositoConsumoResponseDTO
    {
        /// <summary>
        /// Código del depósito de consumo encontrado
        /// </summary>
        public int CodigoDeposito { get; set; }

        /// <summary>
        /// Indica si el depósito está activo en TADEPOSI (Status >= 0)
        /// </summary>
        public bool EstaActivo { get; set; }

        /// <summary>
        /// Descripción del depósito
        /// </summary>
        public string Descripcion { get; set; } = string.Empty;

        /// <summary>
        /// Mensaje informativo sobre el estado del depósito
        /// </summary>
        public string Mensaje { get; set; } = string.Empty;

        /// <summary>
        /// Indica si se requiere selección manual de depósito
        /// </summary>
        public bool RequiereSeleccion { get; set; }
    }
}
