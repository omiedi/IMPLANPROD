namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para datos de compras por mes
    /// </summary>
    public class DatoMensualDTO
    {
        public string Mes { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public int Cantidad { get; set; }
        
        // Propiedades auxiliares para ordenamiento cronológico
        public int Año { get; set; }
        public int NumeroMes { get; set; }
    }

    /// <summary>
    /// DTO para datos de proveedores
    /// </summary>
    public class DatoProveedorDTO
    {
        public string Nombre { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public int CantidadOrdenes { get; set; }
    }

    /// <summary>
    /// DTO para datos de estados de órdenes
    /// </summary>
    public class DatoEstadoDTO
    {
        public string Estado { get; set; } = string.Empty;
        public int Cantidad { get; set; }
        public decimal MontoTotal { get; set; }
    }

    /// <summary>
    /// DTO para datos de monedas
    /// </summary>
    public class DatoMonedaDTO
    {
        public string Moneda { get; set; } = string.Empty;
        public string CodigoMoneda { get; set; } = string.Empty;
        public decimal Monto { get; set; }
        public int CantidadOrdenes { get; set; }
    }

    /// <summary>
    /// DTO principal que contiene todos los datos para los gráficos
    /// </summary>
    public class GraficosComprasDTO
    {
        public List<DatoMensualDTO> ComprasPorMes { get; set; } = new();
        public List<DatoProveedorDTO> TopProveedores { get; set; } = new();
        public List<DatoEstadoDTO> EstadosOrdenes { get; set; } = new();
        public List<DatoMonedaDTO> MontosPorMoneda { get; set; } = new();
        
        // Estadísticas generales
        public int TotalOrdenes { get; set; }
        public decimal MontoTotal { get; set; }
        public int TotalProveedores { get; set; }
        public decimal PromedioMensual { get; set; }
        
        // Filtros aplicados
        public DateTime FechaDesde { get; set; }
        public DateTime FechaHasta { get; set; }
    }
}
