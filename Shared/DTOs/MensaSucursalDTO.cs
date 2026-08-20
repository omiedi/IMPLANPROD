namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para mensajes de sucursal
    /// Usado en el modal de detalle de OC Sucursal
    /// </summary>
    public class MensaSucursalDTO
    {
        /// <summary>
        /// Orden de compra de sucursal
        /// </summary>
        public int? OCompraSuc { get; set; }

        /// <summary>
        /// Mensaje de la sucursal
        /// </summary>
        public string? Mensaje { get; set; }

        /// <summary>
        /// Número de pedido generado en Avellaneda
        /// </summary>
        public int? Nroped { get; set; }
    }
}
