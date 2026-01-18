namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para representar productos/repuestos desde la tabla MASTER
    /// Mantiene el nombre RepuestoDTO por compatibilidad, pero trabaja con MASTER
    /// </summary>
    public class RepuestoDTO
    {
        /// <summary>
        /// Código interno del producto (Master.Codint)
        /// </summary>
        public int Codint { get; set; } = 0;
        
        /// <summary>
        /// Código externo del producto (Master.Codigo)
        /// </summary>
        public string CodigoExterno { get; set; } = string.Empty;
        
        /// <summary>
        /// Descripción del producto (Master.Descripcion)
        /// </summary>
        public string Descritem { get; set; } = string.Empty;
        
        /// <summary>
        /// Unidad de medida (Master.Umedida) - Código de 4 caracteres
        /// </summary>
        public string Uni_Com { get; set; } = string.Empty;
        
        /// <summary>
        /// Descripción completa de la unidad de medida (para visualización)
        /// </summary>
        public string UnidadMedidaDescripcion { get; set; } = string.Empty;
    }
}
