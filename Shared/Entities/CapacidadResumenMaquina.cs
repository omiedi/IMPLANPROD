using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Registro del resumen de máquinas/equipos generado al calcular la capacidad.
    /// Equivalente a la tabla VB6 RECALMAQ.
    /// </summary>
    [Table("RECALMAQ")]
    public class CapacidadResumenMaquina : IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        /// <summary>Equipo (código de máquina, 6 caracteres)</summary>
        [Required, MaxLength(6)]
        public string Equipo { get; set; } = string.Empty;

        /// <summary>Descripción del equipo (48 caracteres)</summary>
        [MaxLength(48)]
        public string DescripcionEquipo { get; set; } = string.Empty;

        /// <summary>Código de pieza/producto</summary>
        public int CodigoPieza { get; set; }

        /// <summary>Descripción de la pieza (64 caracteres)</summary>
        [MaxLength(64)]
        public string DescripcionPieza { get; set; } = string.Empty;

        /// <summary>Número de operación</summary>
        public int NumeroOperacion { get; set; }

        /// <summary>Descripción de la operación (48 caracteres)</summary>
        [MaxLength(48)]
        public string DescripcionOperacion { get; set; } = string.Empty;

        /// <summary>Cantidad programada</summary>
        public decimal CantidadProgramada { get; set; }

        /// <summary>Código del producto destino (puede ser mismo que CodigoPieza)</summary>
        public int CodigoProductoDestino { get; set; }

        /// <summary>Horas necesarias para la operación</summary>
        public decimal HorasNecesarias { get; set; }

        /// <summary>Total de horas por equipo (acumulado)</summary>
        public decimal TotalHorasPorEquipo { get; set; }

        /// <summary>Identificador del programa al que pertenece</summary>
        public int NumeroPrograma { get; set; }

        /// <summary>Marca de auditoría</summary>
        public DateTime FechaCreacion { get; set; } = DateTime.Now;

        /// <summary>
        /// Código de empresa
        /// </summary>
        public int? CodiEmprNet { get; set; } = 0;

        [Display(Name = "Fecha de Creación")]
        public DateTime? AddRecord { get; set; }

        [Display(Name = "Última Actualización")]
        public DateTime? LastUpdate { get; set; }
        /// <summary>
        /// Número de usuario
        /// </summary>
        public int? NumUsuar { get; set; }
    }
}
/*CREATE TABLE RECALMAQ (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Equipo VARCHAR(6) NOT NULL,
    DescripcionEquipo VARCHAR(48) NULL,
    CodigoPieza INT NULL,
    DescripcionPieza VARCHAR(64) NULL,
    NumeroOperacion INT NULL,
    DescripcionOperacion VARCHAR(48) NULL,
    CantidadProgramada DECIMAL(18,2) NULL,
    CodigoProductoDestino INT NULL,
    HorasNecesarias DECIMAL(18,4) NULL,
    TotalHorasPorEquipo DECIMAL(18,4) NULL,
    NumeroPrograma INT NOT NULL,
    FechaCreacion DATETIME NOT NULL DEFAULT GETDATE(),
        AddRecord DATETIME NULL,
    LastUpdate DATETIME NULL,
    CodiEmprNet INT NULL DEFAULT 0,
    NumUsuar INT NULL
    
);*/