namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para datos del cliente necesarios en la impresión del pedido
    /// Mapea campos de la tabla DEUDOR12
    /// </summary>
    public class DatosClienteImpresionDTO
    {
        /// <summary>
        /// Número de cliente
        /// </summary>
        public int Nume_Cli { get; set; }

        /// <summary>
        /// Domicilio del cliente
        /// </summary>
        public string? Domi_Cli { get; set; }

        /// <summary>
        /// Localidad del cliente
        /// </summary>
        public string? Loca_Cli { get; set; }

        /// <summary>
        /// Teléfono del cliente
        /// </summary>
        public string? Tele_Cli { get; set; }

        /// <summary>
        /// CUIT del cliente
        /// </summary>
        public string? Cuit_Cli { get; set; }
    }
}
