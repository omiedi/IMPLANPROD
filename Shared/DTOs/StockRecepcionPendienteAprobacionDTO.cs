using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para mostrar recepciones pendientes de aprobación en Stock Proyectado
    /// Origen: tabla BRECEPEN (recepciones pendientes de aprobación)
    /// Filtro: CantIngre > CantSalid
    /// </summary>
    public sealed class StockRecepcionPendienteAprobacionDTO
    {
        /// <summary>
        /// Número de boleta de recepción
        /// </summary>
        public int Numebore { get; set; }

        /// <summary>
        /// Número de proveedor
        /// </summary>
        public int? NumeProv { get; set; }

        /// <summary>
        /// Nombre del proveedor
        /// </summary>
        public string Proveedor { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad pendiente de aprobar (CantIngre - CantSalid)
        /// </summary>
        public decimal CantPendiente { get; set; }

        /// <summary>
        /// Remito del proveedor
        /// </summary>
        public string RemitoProveedor { get; set; } = string.Empty;

        /// <summary>
        /// Fecha de recepción
        /// </summary>
        public DateTime FechaRecepcion { get; set; }

        /// <summary>
        /// Lote de la recepción
        /// </summary>
        public string Lote { get; set; } = string.Empty;
    }
}
