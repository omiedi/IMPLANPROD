namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para las fabricaciones pendientes de aprobación (tabla ORFAPEN, CodiMovi = 'OF').
    /// Representa órdenes de fabricación cerradas pero aún no aprobadas.
    /// </summary>
    public sealed class StockFabricacionPendienteAprobacionDTO
    {
        /// <summary>
        /// Documento primario del corte (DoPriCor) que identifica la orden.
        /// </summary>
        public string DoPriCor { get; set; } = string.Empty;

        /// <summary>
        /// Referencia del corte (ReferCor).
        /// </summary>
        public string Referencia { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad pendiente de aprobación = SUM(CantIngre) - SUM(CantSalid).
        /// </summary>
        public decimal CantPendiente { get; set; }

        /// <summary>
        /// Fecha del movimiento (FechMov).
        /// </summary>
        public DateTime? FechMov { get; set; }

        /// <summary>
        /// Identificador de texto del cierre (IDTEXT).
        /// </summary>
        public string IdText { get; set; } = string.Empty;
    }
}
