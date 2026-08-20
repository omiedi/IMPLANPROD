using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para mostrar detalle de pedido pendiente
    /// Origen: tabla PEDDETA
    /// Usado en modal "Asignado a Pedidos" de Stock
    /// </summary>
    public sealed class DetallePedidoPendienteDTO
    {
        /// <summary>
        /// Código interno del producto
        /// </summary>
        public int CodiInt { get; set; }

        /// <summary>
        /// Código externo del producto
        /// </summary>
        public string CodiExt { get; set; } = string.Empty;

        /// <summary>
        /// Descripción del producto
        /// </summary>
        public string Descrip { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad pedida
        /// </summary>
        public decimal CanPed { get; set; }

        /// <summary>
        /// Cantidad entregada
        /// </summary>
        public decimal CantEnt { get; set; }

        /// <summary>
        /// Número de pedido
        /// </summary>
        public int NroPed { get; set; }

        /// <summary>
        /// Fecha solicitada de entrega
        /// </summary>
        public DateTime? FeSolEnt { get; set; }

        /// <summary>
        /// Número de orden/item
        /// </summary>
        public int NuOrItem { get; set; }

        /// <summary>
        /// ID del detalle de pedido
        /// </summary>
        public int PeddetaId { get; set; }

        /// <summary>
        /// Días de atraso (calculado)
        /// </summary>
        public int Atraso { get; set; }

        /// <summary>
        /// Cantidad pendiente de remitir (calculado)
        /// </summary>
        public decimal PendienteRemitir { get; set; }

        /// <summary>
        /// Cantidad a remitir (editable por el usuario)
        /// Por defecto toma el valor de PendienteRemitir
        /// </summary>
        public decimal CantidadARemitir { get; set; }

        /// <summary>
        /// Nivel de trazabilidad del producto
        /// 0 = No trazable (verde)
        /// 1 = Trazable con lote (amarillo)
        /// 2 = Trazable especial - no permite transferencia (rojo)
        /// </summary>
        public short Trazable { get; set; }

        /// <summary>
        /// Lote seleccionado para productos trazables (Trazable = 1)
        /// </summary>
        public string? LoteSeleccionado { get; set; }
    }
}
