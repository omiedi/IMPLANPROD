using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para mostrar pedidos pendientes por cliente
    /// Origen: tabla PEDENCA con join a PEDDETA
    /// Usado en modal "Asignado a Pedidos" de Stock
    /// </summary>
    public sealed class PedidoPendienteClienteDTO
    {
        /// <summary>
        /// Número de pedido
        /// </summary>
        public int NroPed { get; set; }

        /// <summary>
        /// Referencia del pedido
        /// </summary>
        public string Referencia { get; set; } = string.Empty;
    }
}
