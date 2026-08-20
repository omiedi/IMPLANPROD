namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para generar un pedido desde una orden de compra de sucursal
    /// </summary>
    public class GenerarPedidoDTO
    {
        /// <summary>
        /// Código de sucursal
        /// </summary>
        public int Sucursal { get; set; }

        /// <summary>
        /// Orden de compra de sucursal
        /// </summary>
        public int OCompraSuc { get; set; }

        /// <summary>
        /// Número de cliente del pedido (opcional, se calcula según empresa)
        /// </summary>
        public int? NcPedido { get; set; }

        /// <summary>
        /// Número de usuario que genera el pedido
        /// </summary>
        public int? NumUsuar { get; set; }
    }
}
