namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para el archivo ARCHIVO_GENERO_PEDIDO
    /// Este archivo se genera en la Central después de procesar exitosamente un archivo Ocompra_Sucursal.txt
    /// Contiene el detalle de los pedidos generados (PEDDETA) para enviar a la Sucursal
    /// Formato basado en código VB6 original
    /// </summary>
    public class ArchivoGeneroPedidoDTO
    {
        /// <summary>
        /// Número de cliente (Sucursal)
        /// Ancho: 8 caracteres
        /// </summary>
        public string NCliente { get; set; } = string.Empty;

        /// <summary>
        /// Número de pedido generado
        /// Ancho: 8 caracteres
        /// </summary>
        public string NPedido { get; set; } = string.Empty;

        /// <summary>
        /// Número de orden de compra (O_COMPRA_SUC)
        /// Ancho: 8 caracteres
        /// </summary>
        public string NCompra { get; set; } = string.Empty;

        /// <summary>
        /// Fecha de emisión del pedido
        /// Ancho: 8 caracteres
        /// </summary>
        public string Emision { get; set; } = string.Empty;

        /// <summary>
        /// Código interno del producto (CodiInt)
        /// Ancho: 8 caracteres
        /// </summary>
        public string CodigoInt { get; set; } = string.Empty;

        /// <summary>
        /// Código externo del producto (CODIEXT)
        /// Ancho: 24 caracteres
        /// </summary>
        public string CodigoExt { get; set; } = string.Empty;

        /// <summary>
        /// Descripción del producto
        /// Ancho: 32 caracteres
        /// </summary>
        public string Descripcion { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad pedida
        /// Ancho: 10 caracteres (formato F10.1)
        /// </summary>
        public string CantidadPedida { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad entregada
        /// Ancho: 10 caracteres (formato F10.1)
        /// </summary>
        public string CantidadEntregada { get; set; } = string.Empty;

        /// <summary>
        /// Acción realizada
        /// Ancho: 10 caracteres
        /// </summary>
        public string Accion { get; set; } = string.Empty;

        /// <summary>
        /// Usuario que generó el pedido
        /// Ancho: 8 caracteres
        /// </summary>
        public string Usuario { get; set; } = string.Empty;
    }
}
