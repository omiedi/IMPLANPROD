using System;
using System.Collections.Generic;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para representar una anotación de pedido parseada desde el campo ntext
    /// </summary>
    public class AnotacionPedidoDTO
    {
        /// <summary>
        /// Fecha de la anotación (posiciones 1-8 del formato VB6)
        /// </summary>
        public string Fecha { get; set; } = string.Empty;

        /// <summary>
        /// Usuario que realizó la anotación (posición 69, 24 caracteres)
        /// </summary>
        public string Usuario { get; set; } = string.Empty;

        /// <summary>
        /// Descripción principal de la anotación (posiciones 11-56)
        /// </summary>
        public string Descripcion { get; set; } = string.Empty;

        /// <summary>
        /// Lista de productos incluidos en la anotación
        /// </summary>
        public List<ProductoAnotacionDTO> Productos { get; set; } = new();

        /// <summary>
        /// Observaciones adicionales
        /// </summary>
        public string Observaciones { get; set; } = string.Empty;

        /// <summary>
        /// Texto original completo de la anotación
        /// </summary>
        public string TextoOriginal { get; set; } = string.Empty;
    }

    /// <summary>
    /// DTO para representar un producto dentro de una anotación
    /// </summary>
    public class ProductoAnotacionDTO
    {
        /// <summary>
        /// Código del producto
        /// </summary>
        public string Codigo { get; set; } = string.Empty;

        /// <summary>
        /// Descripción del producto
        /// </summary>
        public string Descripcion { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad del producto
        /// </summary>
        public decimal Cantidad { get; set; }
    }
}
