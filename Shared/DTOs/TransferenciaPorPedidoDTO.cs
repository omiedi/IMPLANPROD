namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para transferencia de productos asignados a un pedido
    /// </summary>
    public class TransferenciaPorPedidoDTO
    {
        /// <summary>
        /// Número de pedido
        /// </summary>
        public int NumeroPedido { get; set; }

        /// <summary>
        /// Código del cliente
        /// </summary>
        public int CodigoCliente { get; set; }

        /// <summary>
        /// Depósito origen
        /// </summary>
        public int DepositoOrigen { get; set; }

        /// <summary>
        /// Depósito destino
        /// </summary>
        public int DepositoDestino { get; set; }

        /// <summary>
        /// Referencia de la transferencia
        /// </summary>
        public string Referencia { get; set; } = string.Empty;

        /// <summary>
        /// Lista de items a transferir
        /// </summary>
        public List<ItemTransferenciaPedidoDTO> Items { get; set; } = new();
    }

    /// <summary>
    /// DTO para cada item de la transferencia por pedido
    /// </summary>
    public class ItemTransferenciaPedidoDTO
    {
        /// <summary>
        /// ID del detalle del pedido (PeddetaId)
        /// </summary>
        public int PeddetaId { get; set; }

        /// <summary>
        /// Código interno del producto
        /// </summary>
        public int CodigoInterno { get; set; }

        /// <summary>
        /// Código externo del producto
        /// </summary>
        public string CodigoExterno { get; set; } = string.Empty;

        /// <summary>
        /// Descripción del producto
        /// </summary>
        public string Descripcion { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad a transferir
        /// </summary>
        public decimal Cantidad { get; set; }

        /// <summary>
        /// Lote (obligatorio para productos trazables = 1)
        /// </summary>
        public string? Lote { get; set; }

        /// <summary>
        /// Nivel de trazabilidad (0, 1, 2)
        /// </summary>
        public short Trazable { get; set; }
    }
}
