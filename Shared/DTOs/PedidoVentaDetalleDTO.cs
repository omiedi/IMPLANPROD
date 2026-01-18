using System;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO liviano para devolver el detalle de un pedido desde el backend
    /// evitando problemas de mapeo de tipos (smallint/int) y exponiendo solo
    /// los campos necesarios para la edición en el formulario.
    /// </summary>
    public class PedidoVentaDetalleDTO
    {
        /// <summary>
        /// Código interno del producto (MASTER.CODINT)
        /// </summary>
        public int? CodiInt { get; set; }

        /// <summary>
        /// Código externo del producto
        /// </summary>
        public string? CodiExt { get; set; }

        /// <summary>
        /// Descripción del producto
        /// </summary>
        public string? Descrip { get; set; }

        /// <summary>
        /// Cantidad pedida
        /// </summary>
        public decimal? CanPed { get; set; }

        /// <summary>
        /// Precio unitario
        /// </summary>
        public decimal? PreUnid { get; set; }
    }
}
