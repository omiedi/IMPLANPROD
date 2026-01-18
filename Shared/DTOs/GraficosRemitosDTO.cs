namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para datos de remitos por mes
    /// </summary>
    public class DatoMensualRemitoDTO
    {
        public string Mes { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public int Cantidad { get; set; }
        public int Año { get; set; }
        public int NumeroMes { get; set; }
    }

    /// <summary>
    /// DTO para datos de clientes en gráficos de remitos
    /// </summary>
    public class DatoClienteRemitoDTO
    {
        public string Nombre { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public int CantidadRemitos { get; set; }
    }

    /// <summary>
    /// DTO para estados de remitos
    /// </summary>
    public class DatoEstadoRemitoDTO
    {
        public string Estado { get; set; } = string.Empty;
        public int Cantidad { get; set; }
        public decimal MontoTotal { get; set; }
    }

    /// <summary>
    /// DTO para remitos por depósito
    /// </summary>
    public class DatoDepositoRemitoDTO
    {
        public int CodigoDeposito { get; set; }
        public string Deposito { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public int CantidadRemitos { get; set; }
    }

    /// <summary>
    /// DTO para productos en gráficos de remitos
    /// </summary>
    public class DatoProductoRemitoDTO
    {
        public string Producto { get; set; } = string.Empty;
        public int CantidadUnidades { get; set; }
        public int CantidadRemitos { get; set; }
    }

    /// <summary>
    /// DTO principal que contiene todos los datos para los gráficos de remitos
    /// </summary>
    public class GraficosRemitosDTO
    {
        public List<DatoMensualRemitoDTO> RemitosPorMes { get; set; } = new();
        public List<DatoClienteRemitoDTO> TopClientes { get; set; } = new();
        public List<DatoEstadoRemitoDTO> EstadosRemitos { get; set; } = new();
        public List<DatoDepositoRemitoDTO> RemitosPorDeposito { get; set; } = new();
        public List<DatoProductoRemitoDTO> TopProductos { get; set; } = new();

        public int TotalRemitos { get; set; }
        public decimal MontoTotal { get; set; }
        public int TotalClientes { get; set; }
        public decimal PromedioMensual { get; set; }
        public int TotalProductosDistintos { get; set; }

        public DateTime FechaDesde { get; set; }
        public DateTime FechaHasta { get; set; }
    }
}
