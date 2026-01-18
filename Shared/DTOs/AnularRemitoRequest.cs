namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para solicitud de anulación de remito
    /// </summary>
    public class AnularRemitoRequest
    {
        /// <summary>
        /// Número de remito a anular (DoPriCor)
        /// </summary>
        public string NumeroRemito { get; set; } = string.Empty;

        /// <summary>
        /// Número de cliente
        /// </summary>
        public int NumeroCliente { get; set; }

        /// <summary>
        /// Fecha del movimiento
        /// </summary>
        public DateTime FechaMovimiento { get; set; }
    }

    /// <summary>
    /// DTO para respuesta de anulación de remito
    /// </summary>
    public class AnularRemitoResponse
    {
        /// <summary>
        /// Indica si la anulación fue exitosa
        /// </summary>
        public bool Exitoso { get; set; }

        /// <summary>
        /// Mensaje de resultado
        /// </summary>
        public string Mensaje { get; set; } = string.Empty;

        /// <summary>
        /// Cantidad de registros afectados en movimientos
        /// </summary>
        public int RegistrosMovimientosAfectados { get; set; }

        /// <summary>
        /// Cantidad de pedidos actualizados
        /// </summary>
        public int PedidosActualizados { get; set; }

        /// <summary>
        /// Tabla donde se realizó la anulación (MOVISTO o MOVIREPA)
        /// </summary>
        public string TablaUtilizada { get; set; } = string.Empty;
    }
}
