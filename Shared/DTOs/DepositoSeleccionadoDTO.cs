namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO que representa el resultado de la selección de depósito
    /// Utilizado por el componente DepositoSelectorModal
    /// </summary>
    public class DepositoSeleccionadoDTO
    {
        /// <summary>
        /// Código del depósito seleccionado
        /// </summary>
        public int CodigoDepo { get; set; }

        /// <summary>
        /// Descripción completa del depósito
        /// </summary>
        public string Descripcion { get; set; } = string.Empty;

        /// <summary>
        /// Descripción abreviada del depósito
        /// </summary>
        public string DescripcionAbreviada { get; set; } = string.Empty;

        /// <summary>
        /// Código de planta asociado al depósito (opcional)
        /// </summary>
        public int? CodPlanta { get; set; }
    }
}
