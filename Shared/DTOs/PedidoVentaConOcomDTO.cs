namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Fila del reporte "PEDIDOS C/O.Compra x Fechas" de /pedidos-venta.
    /// Lee directamente el encabezado de PEDENCA (sin ir al detalle) para listar,
    /// dentro de un rango de fechas de emisión, los pedidos junto con el número
    /// de Orden de Compra del cliente (Pedenca.NroOcom).
    /// Excluye siempre el cliente "998" (cliente genérico / venta mostrador).
    /// </summary>
    public sealed class PedidoVentaConOcomDTO
    {
        /// <summary>Fecha de emisión del pedido (Pedenca.FechEmis).</summary>
        public DateTime? FechEmis { get; set; }

        public int? NroPed { get; set; }

        public int? NroClie { get; set; }

        /// <summary>Razón social del cliente (Pedenca.RasoClie).</summary>
        public string RasoClie { get; set; } = string.Empty;

        /// <summary>Número de Orden de Compra del cliente (Pedenca.NroOcom).</summary>
        public string NroOcom { get; set; } = string.Empty;
    }
}
