namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para crear o actualizar registros en FLEXCRL
    /// </summary>
    public class FlexcrlRequestDTO
    {
        /// <summary>
        /// Clasificación del proceso de negocio
        /// </summary>
        public string? Clabus { get; set; }

        /// <summary>
        /// Proceso de negocio
        /// </summary>
        public string? Probus { get; set; }

        /// <summary>
        /// Valor de control del proceso
        /// </summary>
        public string? Valcontrol { get; set; }
    }
}
