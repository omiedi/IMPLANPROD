using System.ComponentModel.DataAnnotations;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para la creación de un remito de traslado a sucursal
    /// </summary>
    public class RemitoTrasladoDTO
    {
        /// <summary>
        /// Prefijo del remito (NUMPVTA)
        /// </summary>
        [Required(ErrorMessage = "El prefijo es obligatorio")]
        public short Prefijo { get; set; }

        /// <summary>
        /// Número del remito (DOPRINUM)
        /// </summary>
        [Required(ErrorMessage = "El número de remito es obligatorio")]
        public int Numero { get; set; }

        /// <summary>
        /// Fecha del remito
        /// </summary>
        [Required(ErrorMessage = "La fecha es obligatoria")]
        public DateTime Fecha { get; set; }

        /// <summary>
        /// Número de cliente
        /// </summary>
        public int NumeroCliente { get; set; }

        /// <summary>
        /// Depósito de origen
        /// </summary>
        [Required(ErrorMessage = "El depósito de origen es obligatorio")]
        public int DepositoOrigen { get; set; }

        /// <summary>
        /// Depósito de destino
        /// </summary>
        [Required(ErrorMessage = "El depósito de destino es obligatorio")]
        public int DepositoDestino { get; set; }

        /// <summary>
        /// Referencia del traslado
        /// </summary>
        public string? Referencia { get; set; }

        /// <summary>
        /// Observaciones
        /// </summary>
        public string? Observaciones { get; set; }

        /// <summary>
        /// Items del remito de traslado
        /// </summary>
        [Required(ErrorMessage = "Debe haber al menos un item")]
        public List<RemitoTrasladoItemDTO> Items { get; set; } = new();
    }

    /// <summary>
    /// DTO para un item de remito de traslado
    /// </summary>
    public class RemitoTrasladoItemDTO
    {
        /// <summary>
        /// Código interno del producto
        /// </summary>
        [Required(ErrorMessage = "El código interno es obligatorio")]
        public int CodInt { get; set; }

        /// <summary>
        /// Código externo del producto
        /// </summary>
        public string? Codigo { get; set; }

        /// <summary>
        /// Descripción del producto
        /// </summary>
        public string? Descripcion { get; set; }

        /// <summary>
        /// Indicador de producto trazable
        /// </summary>
        public short Trazable { get; set; }

        /// <summary>
        /// Cantidad
        /// </summary>
        [Required(ErrorMessage = "La cantidad es obligatoria")]
        public decimal Cantidad { get; set; }

        /// <summary>
        /// Unidad de medida
        /// </summary>
        public string? Unidad { get; set; }

        /// <summary>
        /// Número de pedido original (NuPedCli)
        /// </summary>
        public long? NroPedido { get; set; }

        /// <summary>
        /// Número de orden del artículo dentro del pedido (NUMEORDPED)
        /// </summary>
        public int? NuOrItemPedido { get; set; }

        /// <summary>
        /// Lotes (para productos trazables)
        /// </summary>
        public List<LoteDTO> Lotes { get; set; } = new();
    }

    /// <summary>
    /// DTO para un lote
    /// </summary>
    public class LoteDTO
    {
        /// <summary>
        /// Número de lote
        /// </summary>
        public string? Lote { get; set; }

        /// <summary>
        /// Cantidad del lote
        /// </summary>
        public decimal Cantidad { get; set; }
    }
}
