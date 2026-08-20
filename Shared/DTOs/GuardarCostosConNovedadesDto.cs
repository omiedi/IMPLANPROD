namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para guardar costos de orden de compra junto con observación para NOVEDADES
    /// Se usa cuando la configuración ACOSREP es diferente de "NUNCA"
    /// </summary>
    public class GuardarCostosConNovedadesDto
    {
        /// <summary>
        /// Lista de costos a guardar en OCOMDETACosto
        /// </summary>
        public List<OCOMDETACostoDto> Costos { get; set; } = new List<OCOMDETACostoDto>();

        /// <summary>
        /// Observación sobre el cambio de precio
        /// Se guardará en NOVEDADES.Observacion (máximo 512 caracteres)
        /// </summary>
        public string? Observacion { get; set; }

        /// <summary>
        /// Configuración ACOSREP actual (SIEMPRE, PROHABIT, NUNCA)
        /// Determina si se actualiza el costo de reposición y se registra en NOVEDADES
        /// </summary>
        public string ConfiguracionACOSREP { get; set; } = "NUNCA";
    }
}
