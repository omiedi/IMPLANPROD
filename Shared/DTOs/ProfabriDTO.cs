namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para el encabezado de un programa de fabricación
    /// </summary>
    public class ProfabriDTO
    {
        public int Id { get; set; }
        public int NumeroPrograma { get; set; }
        public string Descripcion { get; set; } = null!;
        public DateTime? FechaCreacion { get; set; }
        public int? CodiEmpr { get; set; }
        public int? Status { get; set; }
    }

    /// <summary>
    /// DTO para el detalle de un programa de fabricación
    /// </summary>
    public class DeprofabDTO
    {
        public int Id { get; set; }
        public int NumeroPrograma { get; set; }
        public int CodigoProducto { get; set; }
        public string? CodigoExterno { get; set; }  // Del Master
        public string? Descripcion { get; set; }
        public decimal Cantidad { get; set; }
        public string? Observaciones { get; set; }
    }

    /// <summary>
    /// DTO completo de un programa con sus detalles y cálculos de capacidad
    /// </summary>
    public class ProgramaCapacidadDTO
    {
        public ProfabriDTO? Programa { get; set; }
        public List<DeprofabDTO> Detalles { get; set; } = new();
        public List<CapacidadMaquinaDTO> CapacidadMaquinas { get; set; } = new();
        public List<CapacidadManoObraDTO> CapacidadManoObra { get; set; } = new();
        public TimeSpan TotalTiempoMaquinas { get; set; }
        public TimeSpan TotalTiempoManoObra { get; set; }
    }

    /// <summary>
    /// DTO para el tiempo requerido por máquina
    /// </summary>
    public class CapacidadMaquinaDTO
    {
        public string CodigoMaquina { get; set; } = null!;
        public string DescripcionMaquina { get; set; } = null!;
        public TimeSpan TiempoRequerido { get; set; }
        public string TiempoFormateado => FormatearTiempo(TiempoRequerido);

        private static string FormatearTiempo(TimeSpan tiempo)
        {
            int horas = (int)tiempo.TotalHours;
            int minutos = tiempo.Minutes;
            int segundos = tiempo.Seconds;
            return $"{horas:D2}:{minutos:D2}:{segundos:D2}";
        }
    }

    /// <summary>
    /// DTO para el tiempo requerido por categoría de mano de obra
    /// </summary>
    public class CapacidadManoObraDTO
    {
        public int Categoria { get; set; }
        public string DescripcionCategoria { get; set; } = null!;
        public TimeSpan TiempoRequerido { get; set; }
        public string TiempoFormateado => FormatearTiempo(TiempoRequerido);

        private static string FormatearTiempo(TimeSpan tiempo)
        {
            int horas = (int)tiempo.TotalHours;
            int minutos = tiempo.Minutes;
            int segundos = tiempo.Seconds;
            return $"{horas:D2}:{minutos:D2}:{segundos:D2}";
        }
    }

    /// <summary>
    /// DTO para filtros de generación de programas desde pedidos o remitos
    /// </summary>
    public class FiltroGeneracionDTO
    {
        public DateTime FechaDesde { get; set; }
        public DateTime FechaHasta { get; set; }
        public int? ClienteId { get; set; }  // Opcional, para filtrar por cliente específico
        public string TipoGeneracion { get; set; } = "PEDIDOS"; // "PEDIDOS" o "REMITOS"
        public string? DescripcionPrograma { get; set; }
    }

    /// <summary>
    /// DTO para solicitar el cálculo de capacidad de un programa
    /// </summary>
    public class CalculoCapacidadRequest
    {
        public int ProgramaId { get; set; }
    }
}
