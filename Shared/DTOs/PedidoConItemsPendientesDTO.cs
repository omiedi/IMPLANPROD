using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Pedido de venta que tiene al menos un ítem de PEDDETA pendiente de entrega
    /// (Status &lt; 8 y CantEnt &lt; CanPed - 0.005), usado en la utilidad
    /// "Anular Items de Pedido" de /pedidos-venta.
    /// Equivale al recordset Pedianuittemp del VB6.
    /// </summary>
    public class PedidoConItemsPendientesDTO
    {
        /// <summary>Número de pedido (PEDENCA.NroPed)</summary>
        public int NroPed { get; set; }

        /// <summary>Número de cliente (PEDENCA.NroClie)</summary>
        public int NroClie { get; set; }

        /// <summary>Razón social del cliente (PEDENCA.RasoClie)</summary>
        public string RasoClie { get; set; } = string.Empty;

        /// <summary>Fecha de emisión del pedido (PEDENCA.FechEmis)</summary>
        public DateTime? FechEmis { get; set; }

        /// <summary>Número de orden de compra del cliente (PEDENCA.NroOcom)</summary>
        public string NroOcom { get; set; } = string.Empty;
    }
}
