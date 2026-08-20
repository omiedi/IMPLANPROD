namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para el detalle de productos de una orden de compra de sucursal
    /// Usado en el modal de detalle de OC Sucursal
    /// </summary>
    public class OcSucursalDetalleDTO
    {
        /// <summary>
        /// Código externo del producto
        /// </summary>
        public string? CodExt { get; set; }

        /// <summary>
        /// Descripción del producto
        /// </summary>
        public string? Descripcion { get; set; }

        /// <summary>
        /// Cantidad comprada
        /// </summary>
        public decimal? CantidadCompra { get; set; }

        /// <summary>
        /// Fecha de solicitud de compra
        /// </summary>
        public DateTime? FechaSolicitudCompra { get; set; }

        /// <summary>
        /// Número de pedido generado en Avellaneda
        /// </summary>
        public int? NpedGeneradoEnAvellaneda { get; set; }

        /// <summary>
        /// Cantidad pedida (del pedido en Avellaneda)
        /// </summary>
        public decimal? CantidadPedida { get; set; }

        /// <summary>
        /// Cantidad entregada
        /// </summary>
        public decimal? CantidadEntregada { get; set; }

        /// <summary>
        /// Saldo pendiente (Cantidad Pedida - Cantidad Entregada)
        /// </summary>
        public decimal? Saldo { get; set; }
    }
}
