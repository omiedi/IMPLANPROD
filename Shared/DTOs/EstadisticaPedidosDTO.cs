namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO principal para Estadística de Pedidos de Venta
    /// Contiene todos los datos para el dashboard sin importes monetarios
    /// </summary>
    public class EstadisticaPedidosDTO
    {
        /// <summary>
        /// Cantidad total de pedidos
        /// </summary>
        public int TotalPedidos { get; set; }

        /// <summary>
        /// Cantidad total de renglones de pedido
        /// </summary>
        public int TotalRenglones { get; set; }

        /// <summary>
        /// Suma de cantidades pedidas
        /// </summary>
        public decimal CantidadSolicitada { get; set; }

        /// <summary>
        /// Suma de cantidades remitidas
        /// </summary>
        public decimal CantidadRemitida { get; set; }

        /// <summary>
        /// Suma de cantidades pendientes
        /// </summary>
        public decimal CantidadPendiente { get; set; }

        /// <summary>
        /// Cantidad de pedidos atrasados
        /// </summary>
        public int PedidosAtrasados { get; set; }

        /// <summary>
        /// Evolución de pedidos por mes
        /// </summary>
        public List<DatoMensualPedidosDTO> EvolucionPedidos { get; set; } = new();

        /// <summary>
        /// Pedidos por cliente (top clientes)
        /// </summary>
        public List<DatoClientePedidosDTO> PedidosPorCliente { get; set; } = new();

        /// <summary>
        /// Estado de los pedidos
        /// </summary>
        public List<DatoEstadoPedidosDTO> EstadoPedidos { get; set; } = new();

        /// <summary>
        /// Reporte detallado de pedidos
        /// </summary>
        public List<DetallePedidoEstadisticaDTO> ReporteDetallado { get; set; } = new();
    }

    /// <summary>
    /// DTO para datos mensuales de pedidos
    /// </summary>
    public class DatoMensualPedidosDTO
    {
        public string Mes { get; set; } = "";
        public int CantidadPedidos { get; set; }
        public decimal CantidadSolicitada { get; set; }
    }

    /// <summary>
    /// DTO para datos de pedidos por cliente
    /// </summary>
    public class DatoClientePedidosDTO
    {
        public int NroClie { get; set; }
        public string NombreCliente { get; set; } = "";
        public int CantidadPedidos { get; set; }
        public decimal CantidadSolicitada { get; set; }
    }

    /// <summary>
    /// DTO para estado de pedidos
    /// </summary>
    public class DatoEstadoPedidosDTO
    {
        public string Estado { get; set; } = "";
        public int Cantidad { get; set; }
    }

    /// <summary>
    /// DTO para detalle de pedido en reporte
    /// </summary>
    public class DetallePedidoEstadisticaDTO
    {
        public int NroPed { get; set; }
        public DateTime? FechEmis { get; set; }
        public int? NroClie { get; set; }
        public string? RasoClie { get; set; }
        public int? CodiInt { get; set; }
        public string? CodiExt { get; set; }
        public string? Descrip { get; set; }
        public decimal CanPed { get; set; }
        public decimal CantidadEntregada { get; set; }
        public decimal CantidadPendiente { get; set; }
        public DateTime? FeSolEnt { get; set; }
        public string Estado { get; set; } = "";
    }

    /// <summary>
    /// DTO para detalle de remitos asociados a un pedido y producto
    /// </summary>
    public class DetalleRemitoEstadisticaDTO
    {
        public DateTime FechMov { get; set; }
        public string Documento { get; set; } = "";
        public decimal CantidadRemitida { get; set; }
    }

    /// <summary>
    /// DTO para filtros de estadística de pedidos
    /// </summary>
    public class FiltrosEstadisticaPedidosDTO
    {
        public DateTime? FechaDesde { get; set; }
        public DateTime? FechaHasta { get; set; }
        public int? NroClie { get; set; }
        public short? Status { get; set; }
    }
}
