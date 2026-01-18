namespace IMPLANPROD.Shared.DTOs
{
    public class EstadisticasResumenDTO
    {
        public int TotalOrdenes { get; set; }
        public decimal MontoTotal { get; set; }
        public int TotalProveedores { get; set; }
        public decimal PromedioMensual { get; set; }
    }

    public class ComprasPorMesDTO
    {
        public string Mes { get; set; } = string.Empty;
        public decimal Total { get; set; }
        public int Cantidad { get; set; }
    }

    public class TopProveedoresDTO
    {
        public string Proveedor { get; set; } = string.Empty;
        public string ProveedorCompleto { get; set; } = string.Empty;
        public decimal Total { get; set; }
        public int Cantidad { get; set; }
    }

    public class EstadosOrdenesDTO
    {
        public string Estado { get; set; } = string.Empty;
        public int Cantidad { get; set; }
        public decimal MontoTotal { get; set; }
    }

    public class MontosPorMonedaDTO
    {
        public int CodigoMoneda { get; set; }
        public string Moneda { get; set; } = string.Empty;
        public decimal Total { get; set; }
        public int Cantidad { get; set; }
    }

    public class ResumenCompletoDTO
    {
        public EstadisticasResumenDTO Estadisticas { get; set; } = new();
        public List<ComprasPorMesDTO> ComprasPorMes { get; set; } = new();
        public List<TopProveedoresDTO> TopProveedores { get; set; } = new();
        public List<EstadosOrdenesDTO> EstadosOrdenes { get; set; } = new();
        public List<MontosPorMonedaDTO> MontosPorMoneda { get; set; } = new();
    }
}
