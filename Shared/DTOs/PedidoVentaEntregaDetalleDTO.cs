namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// Detalle de entregas de pedidos de venta (PEDDETA.CantEnt), agrupado por
    /// Pedido + Producto. A diferencia de <see cref="PedidoVentaPendienteDetalleDTO"/>
    /// (que solo muestra ítems con saldo pendiente), este reporte muestra ítems que
    /// tuvieron alguna entrega registrada (CantEnt > 0), estén o no totalmente
    /// entregados. Se usa en el reporte "Entregas de Pedidos" de /pedidos-venta.
    /// </summary>
    public sealed class PedidoVentaEntregaDetalleDTO
    {
        public int NroPed { get; set; }

        /// <summary>Fecha de entrega solicitada del ítem (Peddeta.FeSolEnt).</summary>
        public DateTime? Fecha { get; set; }

        public int? NroClie { get; set; }
        public string Cliente { get; set; } = string.Empty;
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;

        /// <summary>Cantidad pedida (Peddeta.CanPed).</summary>
        public decimal CanPed { get; set; }

        /// <summary>Cantidad entregada (Peddeta.CantEnt).</summary>
        public decimal CantEnt { get; set; }

        /// <summary>Saldo pendiente (CanPed - CantEnt, nunca negativo).</summary>
        public decimal Saldo { get; set; }

        /// <summary>
        /// true cuando la cabecera del pedido (PEDENCA.Status >= 9) está anulada;
        /// el reporte muestra la etiqueta ANULADO en la fila.
        /// </summary>
        public bool Anulado { get; set; }

        /// <summary>Texto de estado para impresión/Excel: ANULADO o VIGENTE.</summary>
        public string Estado => Anulado ? "ANULADO" : "VIGENTE";

        /// <summary>
        /// Código interno del producto (PEDDETA.CodiInt). Se usa internamente
        /// para asociar los remitos de MOVISTO (NuPedCli + Cod_Inte).
        /// </summary>
        public int CodiInt { get; set; }

        /// <summary>
        /// Remito(s) con los que se entregó el ítem (MOVISTO.DoPriLar,
        /// CodiMovi = "RT"). Si el ítem tuvo entregas parciales con distintos
        /// remitos, se listan todos separados por ", ".
        /// </summary>
        public string Remito { get; set; } = string.Empty;

        /// <summary>
        /// Detalle de cada remito asociado a la fila (número + fecha del
        /// movimiento), usado para abrir la vista impresa del remito.
        /// </summary>
        public List<RemitoEntregaDTO> RemitosDetalle { get; set; } = new();
    }

    /// <summary>
    /// Remito que entregó (total o parcialmente) un ítem del pedido.
    /// </summary>
    public sealed class RemitoEntregaDTO
    {
        /// <summary>Número de remito (MOVISTO.DoPriLar).</summary>
        public string Remito { get; set; } = string.Empty;

        /// <summary>Fecha del movimiento de entrega (MOVISTO.FechMov).</summary>
        public DateTime? FechaRemito { get; set; }
    }
}
