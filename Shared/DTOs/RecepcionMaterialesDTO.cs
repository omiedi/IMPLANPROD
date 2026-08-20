namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para el resultado de la recepción de materiales
    /// </summary>
    public class RecepcionMaterialesResult
    {
        /// <summary>
        /// Indica si el proceso fue exitoso
        /// </summary>
        public bool Exitoso { get; set; }

        /// <summary>
        /// Mensaje de resultado
        /// </summary>
        public string Mensaje { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad de archivos procesados
        /// </summary>
        public int ArchivosProcesados { get; set; }

        /// <summary>
        /// Cantidad de registros procesados
        /// </summary>
        public int RegistrosProcesados { get; set; }

        /// <summary>
        /// Lista de errores encontrados
        /// </summary>
        public List<string> Errores { get; set; } = new();
    }

    /// <summary>
    /// DTO para un item de recepción de materiales
    /// </summary>
    public class RecepcionMaterialesItem
    {
        /// <summary>
        /// Número de remito
        /// </summary>
        public string Remito { get; set; } = string.Empty;

        /// <summary>
        /// Código interno del producto
        /// </summary>
        public int CodInte { get; set; }

        /// <summary>
        /// Cantidad recibida
        /// </summary>
        public decimal Cantidad { get; set; }

        /// <summary>
        /// Número de pedido
        /// </summary>
        public long NroPed { get; set; }

        /// <summary>
        /// Número de cliente
        /// </summary>
        public string NumeClie { get; set; } = string.Empty;
    }
}
