using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Configuración de sincronización Central-Sucursal
    /// Almacena rutas de intercambio para diferentes entidades (Productos, Remitos)
    /// y parámetros de comportamiento de sincronización
    /// </summary>
    [Table("SyncConfig")]
    public class SyncConfig
    {
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Indica si está habilitada la generación de scripts de sincronización
        /// </summary>
        [Column("GeneraScriptSync")]
        public bool GeneraScriptSync { get; set; }

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos de Productos (Master, Formulas, SecOper)
        /// </summary>
        [Column("RutaIntercambio")]
        [MaxLength(500)]
        public string RutaIntercambio { get; set; } = string.Empty;

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos de Remitos
        /// Se usa para sincronizar remitos entre Central y Sucursal de Neuquén
        /// </summary>
        [Column("RutaIntercambioRemitos")]
        [MaxLength(500)]
        public string RutaIntercambioRemitos { get; set; } = string.Empty;

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos de Órdenes de Reparación
        /// Se usa para sincronizar consumos de reparación entre Sucursal y Central
        /// </summary>
        [Column("RutaIntercambioReparaciones")]
        [MaxLength(500)]
        public string RutaIntercambioReparaciones { get; set; } = string.Empty;

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos de Ajustes de Stock
        /// Se usa para sincronizar ajustes de stock entre Sucursal y Central
        /// </summary>
        [Column("RutaIntercambioAjustes")]
        [MaxLength(500)]
        public string RutaIntercambioAjustes { get; set; } = string.Empty;

        /// <summary>
        /// Indica si este equipo genera archivos de sincronización de productos (Master, Formulas, SecOper)
        /// Solo debe estar en true para la Central
        /// </summary>
        [Column("GeneraProductos")]
        public bool GeneraProductos { get; set; }

        /// <summary>
        /// Indica si este equipo procesa archivos de sincronización de productos (Master, Formulas, SecOper)
        /// Solo debe estar en true para las Sucursales
        /// </summary>
        [Column("ProcesaProductos")]
        public bool ProcesaProductos { get; set; }

        /// <summary>
        /// Indica si este equipo genera archivos de sincronización de remitos
        /// Solo debe estar en true para la Sucursal de Neuquén
        /// </summary>
        [Column("GeneraRemitos")]
        public bool GeneraRemitos { get; set; }

        /// <summary>
        /// Indica si este equipo procesa archivos de sincronización de remitos
        /// Solo debe estar en true para la Central
        /// </summary>
        [Column("ProcesaRemitos")]
        public bool ProcesaRemitos { get; set; }

        /// <summary>
        /// Indica si este equipo genera archivos de sincronización de órdenes de reparación
        /// Solo debe estar en true para la Sucursal
        /// </summary>
        [Column("GeneraReparaciones")]
        public bool GeneraReparaciones { get; set; }

        /// <summary>
        /// Indica si este equipo procesa archivos de sincronización de órdenes de reparación
        /// Solo debe estar en true para la Central
        /// </summary>
        [Column("ProcesaReparaciones")]
        public bool ProcesaReparaciones { get; set; }

        /// <summary>
        /// Indica si este equipo genera archivos de sincronización de ajustes de stock
        /// Solo debe estar en true para la Sucursal
        /// </summary>
        [Column("GeneraAjustes")]
        public bool GeneraAjustes { get; set; }

        /// <summary>
        /// Indica si este equipo procesa archivos de sincronización de ajustes de stock
        /// Solo debe estar en true para la Central
        /// </summary>
        [Column("ProcesaAjustes")]
        public bool ProcesaAjustes { get; set; }

        /// <summary>
        /// Ruta de la carpeta compartida para intercambio de archivos de Pedidos de Reposición
        /// Se usa para sincronizar pedidos de reposición entre Sucursal y Central
        /// </summary>
        [Column("RutaIntercambioPedidosReposicion")]
        [MaxLength(500)]
        public string RutaIntercambioPedidosReposicion { get; set; } = string.Empty;

        /// <summary>
        /// Indica si este equipo genera archivos de sincronización de pedidos de reposición
        /// Solo debe estar en true para la Sucursal
        /// </summary>
        [Column("GeneraPedidosReposicion")]
        public bool GeneraPedidosReposicion { get; set; }

        /// <summary>
        /// Indica si este equipo procesa archivos de sincronización de pedidos de reposición
        /// Solo debe estar en true para la Central
        /// </summary>
        [Column("ProcesaPedidosReposicion")]
        public bool ProcesaPedidosReposicion { get; set; }

        /// <summary>
        /// Fecha de última actualización de la configuración
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }
    }
}

/*
 * MIGRACIÓN DE BASE DE DATOS
 * Ejecutar este script SQL para agregar los nuevos campos en la tabla SyncConfig
 *
 * -- Paso 1: Agregar campo RutaIntercambioRemitos
 * ALTER TABLE SyncConfig
 * ADD RutaIntercambioRemitos NVARCHAR(500) NOT NULL DEFAULT '';
 *
 * -- Paso 2: Agregar campos específicos por tipo de sincronización
 * ALTER TABLE SyncConfig
 * ADD GeneraProductos BIT NOT NULL DEFAULT 0;
 *
 * ALTER TABLE SyncConfig
 * ADD ProcesaProductos BIT NOT NULL DEFAULT 0;
 *
 * ALTER TABLE SyncConfig
 * ADD GeneraRemitos BIT NOT NULL DEFAULT 0;
 *
 * ALTER TABLE SyncConfig
 * ADD ProcesaRemitos BIT NOT NULL DEFAULT 0;
 *
 * ALTER TABLE SyncConfig
 * ADD RutaIntercambioReparaciones NVARCHAR(500) NOT NULL DEFAULT '';
 *
 * ALTER TABLE SyncConfig
 * ADD GeneraReparaciones BIT NOT NULL DEFAULT 0;
 *
 * ALTER TABLE SyncConfig
 * ADD ProcesaReparaciones BIT NOT NULL DEFAULT 0;
 *
 * ALTER TABLE SyncConfig
 * ADD RutaIntercambioAjustes NVARCHAR(500) NOT NULL DEFAULT '';
 *
 * ALTER TABLE SyncConfig
 * ADD GeneraAjustes BIT NOT NULL DEFAULT 0;
 *
 * ALTER TABLE SyncConfig
 * ADD ProcesaAjustes BIT NOT NULL DEFAULT 0;
 *
 * -- Paso 3: Agregar campos para Pedidos de Reposición
 * ALTER TABLE SyncConfig
 * ADD RutaIntercambioPedidosReposicion NVARCHAR(500) NOT NULL DEFAULT '';
 *
 * ALTER TABLE SyncConfig
 * ADD GeneraPedidosReposicion BIT NOT NULL DEFAULT 0;
 *
 * ALTER TABLE SyncConfig
 * ADD ProcesaPedidosReposicion BIT NOT NULL DEFAULT 0;
 *
 * -- Paso 4: Migrar datos existentes (si hay registros)
 * UPDATE SyncConfig
 * SET GeneraProductos = CASE WHEN EsCentral = 1 THEN 1 ELSE 0 END,
 *     ProcesaProductos = CASE WHEN EsCentral = 0 THEN 1 ELSE 0 END,
 *     GeneraRemitos = 0,
 *     ProcesaRemitos = CASE WHEN EsCentral = 1 THEN 1 ELSE 0 END
 * WHERE EsCentral IS NOT NULL;
 *
 * -- Paso 5: Eliminar campo obsoleto (opcional, después de verificar migración)
 * ALTER TABLE SyncConfig
 * DROP COLUMN EsCentral;
 *
 * Nota: Los nuevos campos permiten configurar el comportamiento de sincronización
 * de forma independiente para productos, remitos, reparaciones, ajustes y pedidos de reposición,
 * eliminando la necesidad de un rol global Central/Sucursal.
 */
