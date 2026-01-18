namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para los datos de ORFAPEN necesarios para la aprobación
    /// </summary>
    public class DatosOrfapenDTO
    {
        /// <summary>
        /// Referencia correlativa
        /// </summary>
        public string ReferCor { get; set; } = string.Empty;

        /// <summary>
        /// Fecha de movimiento
        /// </summary>
        public DateTime FechMov { get; set; }

        /// <summary>
        /// Código interno
        /// </summary>
        public int CodInte { get; set; }

        /// <summary>
        /// Código externo
        /// </summary>
        public string CodExte { get; set; } = string.Empty;

        /// <summary>
        /// Código de movimiento
        /// </summary>
        public string CodiMovi { get; set; } = string.Empty;

        /// <summary>
        /// Documento primario correlativo
        /// </summary>
        public string DoPriCor { get; set; } = string.Empty;

        /// <summary>
        /// Documento secundario correlativo
        /// </summary>
        public string DoSecCor { get; set; } = string.Empty;

        /// <summary>
        /// Depósito de movimiento
        /// </summary>
        public int DepoMovi { get; set; }

        /// <summary>
        /// Depósito de consumo
        /// </summary>
        public int DepoConsu { get; set; }

        /// <summary>
        /// Cantidad pendiente/saldo
        /// </summary>
        public decimal Cantidad { get; set; }

        /// <summary>
        /// Código externo del producto (para mostrar al usuario)
        /// </summary>
        public string CodigoExterno { get; set; } = string.Empty;

        /// <summary>
        /// Descripción del producto
        /// </summary>
        public string Descripcion { get; set; } = string.Empty;
    }
}
