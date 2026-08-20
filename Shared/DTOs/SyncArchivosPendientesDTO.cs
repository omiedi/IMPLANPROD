namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para consultar la cantidad de archivos pendientes de sincronización
    /// </summary>
    public class SyncArchivosPendientesDTO
    {
        /// <summary>
        /// Cantidad de archivos de productos (MASTER_*.json, FORMULAS_*.json, SECOPER_*.json) pendientes
        /// </summary>
        public int ArchivosProductos { get; set; }

        /// <summary>
        /// Cantidad de archivos de remitos (REMITO_*.json) pendientes
        /// </summary>
        public int ArchivosRemitos { get; set; }

        /// <summary>
        /// Cantidad de archivos de órdenes de reparación pendientes
        /// </summary>
        public int ArchivosReparaciones { get; set; }

        /// <summary>
        /// Cantidad de archivos de ajustes de stock pendientes
        /// </summary>
        public int ArchivosAjustes { get; set; }

        /// <summary>
        /// Ruta de intercambio de productos
        /// </summary>
        public string? RutaProductos { get; set; }

        /// <summary>
        /// Ruta de intercambio de remitos
        /// </summary>
        public string? RutaRemitos { get; set; }

        /// <summary>
        /// Ruta de intercambio de órdenes de reparación
        /// </summary>
        public string? RutaReparaciones { get; set; }

        /// <summary>
        /// Ruta de intercambio de ajustes de stock
        /// </summary>
        public string? RutaAjustes { get; set; }
    }
}
