namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para configuración de sincronización
    /// </summary>
    public class SyncConfigDTO
    {
        /// <summary>
        /// Indica si está habilitada la generación de scripts de sincronización
        /// </summary>
        public bool GeneraScriptSync { get; set; }

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos
        /// </summary>
        public string? RutaIntercambio { get; set; }

        /// <summary>
        /// Indica si es Central (genera archivos) o Sucursal (procesa archivos)
        /// </summary>
        public bool EsCentral { get; set; }
    }
}
