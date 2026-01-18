namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para guardar una nueva observación en cualquier tabla
    /// </summary>
    public class GuardarObservacionDTO
    {
        /// <summary>
        /// Nombre de la tabla donde se guardará la observación
        /// Ejemplo: "BADAMODE_NOTA", "MASTER_NOTA"
        /// </summary>
        public string NombreTabla { get; set; } = string.Empty;

        /// <summary>
        /// Nombre del campo identificador en la tabla
        /// Ejemplo: "Cod_Mod", "Codint"
        /// </summary>
        public string CampoIdentificador { get; set; } = string.Empty;

        /// <summary>
        /// Valor del identificador
        /// </summary>
        public string ValorIdentificador { get; set; } = string.Empty;

        /// <summary>
        /// Indica si el campo identificador es de tipo texto (true) o numérico (false)
        /// </summary>
        public bool EsTexto { get; set; } = true;

        /// <summary>
        /// Contenido de la observación a guardar
        /// </summary>
        public string Observaciones { get; set; } = string.Empty;
    }
}
