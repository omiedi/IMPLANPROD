namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO principal que contiene todos los datos para los gráficos de recepciones
    /// </summary>
    public class GraficosRecepcionesDTO
    {
        public List<DatoMensualRecepcionDTO> RecepcionesPorMes { get; set; } = new();
        public List<CalificacionProveedorDTO> CalificacionProveedores { get; set; } = new();
        public List<DatoRechazoDTO> RechazosPorProveedor { get; set; } = new();
        public List<DatoEstadoRecepcionDTO> EstadosRecepciones { get; set; } = new();
        public List<DatoProductoRecepcionDTO> TopProductosRecibidos { get; set; } = new();
        
        // Estadísticas generales
        public int TotalRecepciones { get; set; }
        public decimal CantidadTotalRecibida { get; set; }
        public decimal CantidadTotalRechazada { get; set; }
        public decimal PorcentajeRechazo { get; set; }
        public int TotalProveedores { get; set; }
        public decimal PromedioMensual { get; set; }
    }

    /// <summary>
    /// Datos de recepciones agrupadas por mes
    /// </summary>
    public class DatoMensualRecepcionDTO
    {
        public string Mes { get; set; } = string.Empty;
        public decimal Cantidad { get; set; }
        public int NumeroRecepciones { get; set; }
    }

    /// <summary>
    /// Calificación de proveedores basada en múltiples métricas
    /// </summary>
    public class CalificacionProveedorDTO
    {
        public string Proveedor { get; set; } = string.Empty;
        public int NumeroProveedor { get; set; }
        public decimal PorcentajeAprobacion { get; set; }
        public decimal PorcentajeRechazo { get; set; }
        public decimal PorcentajeFaltante { get; set; }
        public int TotalRecepciones { get; set; }
        public decimal CalificacionGeneral { get; set; } // 0-100
        public string Categoria { get; set; } = string.Empty; // Excelente, Bueno, Regular, Deficiente
    }

    /// <summary>
    /// Datos de rechazos por proveedor
    /// </summary>
    public class DatoRechazoDTO
    {
        public string Proveedor { get; set; } = string.Empty;
        public decimal CantidadRechazada { get; set; }
        public decimal PorcentajeRechazo { get; set; }
    }

    /// <summary>
    /// Estados de recepciones (con/sin OC, con rechazos, etc.)
    /// </summary>
    public class DatoEstadoRecepcionDTO
    {
        public string Estado { get; set; } = string.Empty;
        public int Cantidad { get; set; }
    }

    /// <summary>
    /// Top productos más recibidos
    /// </summary>
    public class DatoProductoRecepcionDTO
    {
        public string Producto { get; set; } = string.Empty;
        public decimal CantidadTotal { get; set; }
        public int NumeroRecepciones { get; set; }
    }
}
