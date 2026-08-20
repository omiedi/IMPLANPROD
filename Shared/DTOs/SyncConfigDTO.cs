namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para configuración de sincronización Central-Sucursal
    /// Contiene rutas de intercambio para diferentes entidades y parámetros de comportamiento
    /// independientes por tipo de sincronización (productos y remitos)
    /// </summary>
    public class SyncConfigDTO
    {
        /// <summary>
        /// Indica si está habilitada la generación de scripts de sincronización
        /// </summary>
        public bool GeneraScriptSync { get; set; }

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos de Productos (Master, Formulas, SecOper)
        /// </summary>
        public string? RutaIntercambio { get; set; }

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos de Remitos
        /// Se usa para sincronizar remitos entre Central y Sucursal de Neuquén
        /// </summary>
        public string? RutaIntercambioRemitos { get; set; }

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos de Órdenes de Reparación
        /// Se usa para sincronizar consumos de reparación entre Sucursal y Central
        /// </summary>
        public string? RutaIntercambioReparaciones { get; set; }

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos de Ajustes de Stock
        /// Se usa para sincronizar ajustes de stock entre Sucursal y Central
        /// </summary>
        public string? RutaIntercambioAjustes { get; set; }

        /// <summary>
        /// Indica si este equipo genera archivos de sincronización de productos (Master, Formulas, SecOper)
        /// Solo debe estar en true para la Central
        /// </summary>
        public bool GeneraProductos { get; set; }

        /// <summary>
        /// Indica si este equipo procesa archivos de sincronización de productos (Master, Formulas, SecOper)
        /// Solo debe estar en true para las Sucursales
        /// </summary>
        public bool ProcesaProductos { get; set; }

        /// <summary>
        /// Indica si este equipo genera archivos de sincronización de remitos
        /// Solo debe estar en true para la Sucursal de Neuquén
        /// </summary>
        public bool GeneraRemitos { get; set; }

        /// <summary>
        /// Indica si este equipo procesa archivos de sincronización de remitos
        /// Solo debe estar en true para la Central
        /// </summary>
        public bool ProcesaRemitos { get; set; }

        /// <summary>
        /// Indica si este equipo genera archivos de sincronización de órdenes de reparación
        /// Solo debe estar en true para la Sucursal
        /// </summary>
        public bool GeneraReparaciones { get; set; }

        /// <summary>
        /// Indica si este equipo procesa archivos de sincronización de órdenes de reparación
        /// Solo debe estar en true para la Central
        /// </summary>
        public bool ProcesaReparaciones { get; set; }

        /// <summary>
        /// Indica si este equipo genera archivos de sincronización de ajustes de stock
        /// Solo debe estar en true para la Sucursal
        /// </summary>
        public bool GeneraAjustes { get; set; }

        /// <summary>
        /// Indica si este equipo procesa archivos de sincronización de ajustes de stock
        /// Solo debe estar en true para la Central
        /// </summary>
        public bool ProcesaAjustes { get; set; }

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos de Pedidos de Reposición
        /// Se usa para sincronizar pedidos de reposición entre Sucursal y Central
        /// </summary>
        public string? RutaIntercambioPedidosReposicion { get; set; }

        /// <summary>
        /// Indica si este equipo genera archivos de sincronización de pedidos de reposición
        /// Solo debe estar en true para la Sucursal
        /// </summary>
        public bool GeneraPedidosReposicion { get; set; }

        /// <summary>
        /// Indica si este equipo procesa archivos de sincronización de pedidos de reposición
        /// Solo debe estar en true para la Central
        /// </summary>
        public bool ProcesaPedidosReposicion { get; set; }
    }
}
