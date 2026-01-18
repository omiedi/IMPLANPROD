namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para los datos de fin de proceso de ORDEFABR
    /// </summary>
    public class OrdefabrFinProcesoDTO
    {
        /// <summary>
        /// Cantidad de fin de proceso
        /// </summary>
        public decimal CantiFinProceso { get; set; }

        /// <summary>
        /// Cantidad aprobada
        /// </summary>
        public decimal CanApro { get; set; }

        /// <summary>
        /// Cantidad proyectada
        /// </summary>
        public decimal CanProy { get; set; }

        /// <summary>
        /// Estado de la orden de fabricación
        /// </summary>
        public int StatuOrf { get; set; }

        /// <summary>
        /// Estado del proceso
        /// </summary>
        public int StatusProceso { get; set; }
    }
}
