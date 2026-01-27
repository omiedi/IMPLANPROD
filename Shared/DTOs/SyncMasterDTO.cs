using IMPLANPROD.Shared.Entities;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para sincronización de datos de Master entre Central y Sucursal
    /// </summary>
    public class SyncMasterDTO
    {
        /// <summary>
        /// Nombre de la tabla que se sincroniza
        /// </summary>
        public string Tabla { get; set; } = "MASTER";

        /// <summary>
        /// Operación a realizar: UPSERT
        /// </summary>
        public string Operacion { get; set; } = "UPSERT";

        /// <summary>
        /// Fecha y hora de generación del archivo de sincronización
        /// </summary>
        public DateTime Fecha { get; set; }

        /// <summary>
        /// Datos completos del registro Master a sincronizar
        /// </summary>
        public Master Datos { get; set; } = new Master();
    }
}
