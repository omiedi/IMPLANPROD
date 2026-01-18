namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para los datos de gráficos de pedidos de venta
    /// </summary>
    public class GraficosPedidosVentaDTO
    {
        public List<DatoMensualDTO> VentasPorMes { get; set; } = new();
        public List<DatoClienteDTO> TopClientes { get; set; } = new();
        public List<DatoEstadoDTO> EstadosPedidos { get; set; } = new();
        public List<DatoMonedaDTO> MontosPorMoneda { get; set; } = new();
        public List<DatoVendedorDTO> TopVendedores { get; set; } = new();
        public List<DatoProductoDTO> TopProductos { get; set; } = new();
        
        // Estadísticas generales
        public int TotalPedidos { get; set; }
        public decimal MontoTotal { get; set; }
        public int TotalClientes { get; set; }
        public decimal PromedioMensual { get; set; }
        public int TotalVendedores { get; set; }
        public int TotalProductos { get; set; }
    }

    /// <summary>
    /// DTO para datos de clientes en gráficos
    /// </summary>
    public class DatoClienteDTO
    {
        public string Nombre { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public int Cantidad { get; set; }
    }

    /// <summary>
    /// DTO para datos de vendedores en gráficos
    /// </summary>
    public class DatoVendedorDTO
    {
        public string Nombre { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public int Cantidad { get; set; }
    }

    /// <summary>
    /// DTO para datos de productos en gráficos
    /// </summary>
    public class DatoProductoDTO
    {
        public string Codigo { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public decimal Cantidad { get; set; }
        public decimal Monto { get; set; }
    }
}
