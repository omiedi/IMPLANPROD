namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para consulta de órdenes de compra de sucursal
    /// </summary>
    public class OcSucursalConsultaDTO
    {
        /// <summary>
        /// Código de sucursal
        /// </summary>
        public short? Sucursal { get; set; }

        /// <summary>
        /// Razón social de la sucursal (se obtiene de tabla de clientes)
        /// </summary>
        public string? RazonSocialSucursal { get; set; }

        /// <summary>
        /// Orden de compra de sucursal
        /// </summary>
        public int? OCompraSuc { get; set; }

        /// <summary>
        /// Fecha de emisión
        /// </summary>
        public DateTime? Fechemisi { get; set; }

        /// <summary>
        /// Número de cliente del pedido
        /// </summary>
        public int? NclienPedido { get; set; }

        /// <summary>
        /// Razón social del cliente de sucursal
        /// </summary>
        public string? RazCteSuc { get; set; }

        /// <summary>
        /// Número de pedido generado en Avellaneda
        /// </summary>
        public int? NpedGeneradoEnAvellaneda { get; set; }

        /// <summary>
        /// Orden de compra del cliente
        /// </summary>
        public string? OComCliente { get; set; }

        /// <summary>
        /// Versión de la compra de sucursal
        /// </summary>
        public short? VersionCompraSuc { get; set; }
    }
}
