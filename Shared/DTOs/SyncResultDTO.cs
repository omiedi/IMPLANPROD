namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para resultado de procesamiento de sincronización
    /// </summary>
    public class SyncResultDTO
    {
        /// <summary>
        /// Indica si el procesamiento fue exitoso
        /// </summary>
        public bool Exitoso { get; set; }

        /// <summary>
        /// Cantidad de archivos procesados
        /// </summary>
        public int ArchivosProcessados { get; set; }

        /// <summary>
        /// Cantidad de registros sincronizados
        /// </summary>
        public int RegistrosSincronizados { get; set; }

        /// <summary>
        /// Cantidad de errores encontrados
        /// </summary>
        public int Errores { get; set; }

        /// <summary>
        /// Mensaje descriptivo del resultado
        /// </summary>
        public string Mensaje { get; set; } = string.Empty;

        /// <summary>
        /// Lista de errores detallados
        /// </summary>
        public List<string> DetalleErrores { get; set; } = new List<string>();
    }
}
