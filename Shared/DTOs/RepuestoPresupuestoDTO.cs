namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para mostrar el detalle de repuestos en un presupuesto de reparación
    /// </summary>
    public class RepuestoPresupuestoDTO
    {
        /// <summary>
        /// Código interno del producto (REPUESRE.CodigoInterno)
        /// </summary>
        public int CodigoInterno { get; set; }

        /// <summary>
        /// Código del producto (MASTER.Codigo)
        /// </summary>
        public string? Codigo { get; set; }

        /// <summary>
        /// Descripción del producto (MASTER.Descripcion)
        /// </summary>
        public string? Descripcion { get; set; }

        /// <summary>
        /// Uso unitario del repuesto (REPUESRE.Uso)
        /// </summary>
        public decimal Uso { get; set; }

        /// <summary>
        /// Costo unitario del producto (MASTER.CosUni)
        /// </summary>
        public decimal? CostoUnitarioBase { get; set; }

        /// <summary>
        /// Costo unitario con coeficiente aplicado (Coeficiente * MASTER.CosUni)
        /// </summary>
        public decimal CostoUnitario { get; set; }

        /// <summary>
        /// Importe total (CostoUnitario * Uso)
        /// </summary>
        public decimal Importe { get; set; }
    }
}
