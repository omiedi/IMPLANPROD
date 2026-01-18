namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para mostrar repuestos de una reparación con información del Master
    /// </summary>
    public class RepuestoReparacionDTO
    {
        /// <summary>
        /// ID del registro en REPUESRE
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Código de reparación
        /// </summary>
        public string CodigoReparacion { get; set; } = string.Empty;

        /// <summary>
        /// Código interno del producto (REPUESRE.CodigoInterno)
        /// </summary>
        public int CodigoInterno { get; set; }

        /// <summary>
        /// Código externo del producto (MASTER.Codigo)
        /// </summary>
        public string? CodigoExterno { get; set; }

        /// <summary>
        /// Unidad de medida (MASTER.Umedida)
        /// </summary>
        public string? UnidadMedida { get; set; }

        /// <summary>
        /// Descripción del producto (MASTER.Descripcion)
        /// </summary>
        public string? Descripcion { get; set; }

        /// <summary>
        /// Cantidad de uso (REPUESRE.Uso)
        /// </summary>
        public decimal Uso { get; set; }
    }

    /// <summary>
    /// DTO para agregar un nuevo repuesto a una reparación
    /// </summary>
    public class AgregarRepuestoRequest
    {
        /// <summary>
        /// Código de reparación
        /// </summary>
        public string CodigoReparacion { get; set; } = string.Empty;

        /// <summary>
        /// Código interno del producto
        /// </summary>
        public short CodigoInterno { get; set; }

        /// <summary>
        /// Cantidad de uso
        /// </summary>
        public decimal Uso { get; set; }
    }
}
