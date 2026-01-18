using System;
using System.Collections.Generic;

namespace IMPLANPROD.Shared.DTOs
{
    /// <summary>
    /// DTO para representar un programa seleccionado
    /// </summary>
    public class ProgramaSeleccionDto
    {
        public int NumeroPrograma { get; set; }
        public string Descripcion { get; set; } = string.Empty;
        public DateTime? FechaCreacion { get; set; }
    }

    /// <summary>
    /// DTO para representar una pieza con cantidades por programa
    /// </summary>
    public class PiezaProgramaDto
    {
        public string CodigoProducto { get; set; } = string.Empty;
        public string DescripcionPieza { get; set; } = string.Empty;
        
        /// <summary>
        /// Cantidades por programa. Key: NumeroPrograma, Value: Cantidad
        /// </summary>
        public Dictionary<int, decimal> CantidadesPorPrograma { get; set; } = new();
        
        /// <summary>
        /// Promedio de cantidades entre todos los programas
        /// </summary>
        public decimal Promedio { get; set; }
    }

    /// <summary>
    /// DTO para representar la carga horaria por máquina/operación
    /// </summary>
    public class CargaHorariaMaquinaDto
    {
        public int CodigoMaquina { get; set; }
        public string DescripcionMaquina { get; set; } = string.Empty;
        public int CodigoOperacion { get; set; }
        public string DescripcionOperacion { get; set; } = string.Empty;
        public int CategoriaOperario { get; set; }
        public string NombreOperario { get; set; } = string.Empty;
        
        /// <summary>
        /// Tiempos por programa en segundos. Key: NumeroPrograma, Value: TiempoSegundos
        /// </summary>
        public Dictionary<int, decimal> TiemposPorPrograma { get; set; } = new();
        
        /// <summary>
        /// Promedio de tiempos en segundos
        /// </summary>
        public decimal Promedio { get; set; }
    }

    /// <summary>
    /// Response DTO que contiene toda la información procesada
    /// </summary>
    public class CargaHorariaResponseDto
    {
        /// <summary>
        /// Lista de programas seleccionados
        /// </summary>
        public List<ProgramaSeleccionDto> Programas { get; set; } = new();
        
        /// <summary>
        /// Detalles de piezas agrupadas con cantidades por programa
        /// </summary>
        public List<PiezaProgramaDto> Piezas { get; set; } = new();
        
        /// <summary>
        /// Carga horaria por máquina/operación
        /// </summary>
        public List<CargaHorariaMaquinaDto> CargaHoraria { get; set; } = new();
    }

    /// <summary>
    /// Request DTO con los números de programa seleccionados
    /// </summary>
    public class CargaHorariaRequestDto
    {
        public List<int> NumerosProgramas { get; set; } = new();

    }
}
