using System.Collections.Generic;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Request para marcar como cumplidos (Status = 8) los ítems seleccionados
    /// de un pedido de venta. Se usa en la utilidad "Anular Items de Pedido"
    /// de /pedidos-venta (equivale al "Marcar Cumplidos" del VB6).
    /// </summary>
    public class MarcarItemsCumplidosDTO
    {
        /// <summary>Número de pedido (PEDDETA.NroPed)</summary>
        public int NroPed { get; set; }

        /// <summary>Números de ítem dentro del pedido (PEDDETA.NuOrItem)</summary>
        public List<int> NuOrItems { get; set; } = new();
    }
}
