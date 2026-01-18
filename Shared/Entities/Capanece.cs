using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa los registros de cálculo de capacidad necesaria.
    /// Esta tabla almacena el detalle completo del cálculo de capacidad por equipo, pieza y operación.
    /// Mapea a la tabla CAPANECE en la base de datos.
    /// Se genera al ejecutar el cálculo de capacidad desde PROFABRI.
    /// Implementa IAuditableEntity para auditoría automática.
    /// </summary>
    [Table("CAPANECE")]
    public class Capanece : IAuditableEntity
    {
        /// <summary>
        /// Identificador único del registro (campo IDENTITY)
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Código de equipo/máquina (de PADMAQ)
        /// </summary>
        [Required]
        [MaxLength(6)]
        public string Equipo { get; set; } = null!;

        /// <summary>
        /// Descripción del equipo/máquina
        /// </summary>
        [MaxLength(48)]
        public string? DescripcionEquipo { get; set; }

        /// <summary>
        /// Código interno de la pieza (de MASTER)
        /// </summary>
        public int? CodigoPieza { get; set; }

        /// <summary>
        /// Descripción de la pieza
        /// </summary>
        [MaxLength(64)]
        public string? DescripcionPieza { get; set; }

        /// <summary>
        /// Número de operación dentro del proceso
        /// </summary>
        public int? NumeroOperacion { get; set; }

        /// <summary>
        /// Descripción de la operación
        /// </summary>
        [MaxLength(48)]
        public string? DescripcionOperacion { get; set; }

        /// <summary>
        /// Cantidad programada del producto
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantidadProgramada { get; set; }

        /// <summary>
        /// Cantidad total de la pieza
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantidadTotalPieza { get; set; }

        /// <summary>
        /// Código de producto destino (de MASTER)
        /// </summary>
        public int? CodigoProductoDestino { get; set; }

        /// <summary>
        /// Descripción del producto destino
        /// </summary>
        [MaxLength(64)]
        public string? Descripcion { get; set; }

        /// <summary>
        /// Horas necesarias para esta operación
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? HorasNecesarias { get; set; }

        /// <summary>
        /// Total de horas por pieza/operación
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TotalHorasPorPieza { get; set; }

        /// <summary>
        /// Total de horas por equipo
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TotalHorasPorEquipo { get; set; }

        /// <summary>
        /// Número de programa de fabricación asociado (FK a PROFABRI)
        /// Permite relacionar estos cálculos con el programa que los generó
        /// </summary>
        public int? NumeroPrograma { get; set; }

        // Campos de auditoría - IAuditableEntity
        [Display(Name = "Fecha de Creación")]
        public DateTime? AddRecord { get; set; }

        [Display(Name = "Última Actualización")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa para auditoría .NET
        /// </summary>
        public int? CodiEmprNet { get; set; } = 0;

        /// <summary>
        /// Número de usuario que realizó la última modificación
        /// </summary>
        public int? NumUsuar { get; set; }
    }
}

/*
-- Script SQL para crear la tabla CAPANECE
-- Tabla de Cálculo de Capacidad Necesaria
-- Almacena el detalle del cálculo de capacidad por equipo, pieza y operación

CREATE TABLE CAPANECE (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Equipo NVARCHAR(6) NOT NULL,
    DescripcionEquipo NVARCHAR(48) NULL,
    CodigoPieza INT NULL,
    DescripcionPieza NVARCHAR(64) NULL,
    NumeroOperacion INT NULL,
    DescripcionOperacion NVARCHAR(48) NULL,
    CantidadProgramada DECIMAL(18,4) NULL,
    CantidadTotalPieza DECIMAL(18,4) NULL,
    CodigoProductoDestino INT NULL,
    Descripcion NVARCHAR(64) NULL,
    HorasNecesarias DECIMAL(18,4) NULL,
    TotalHorasPorPieza DECIMAL(18,4) NULL,
    TotalHorasPorEquipo DECIMAL(18,4) NULL,
    NumeroPrograma INT NULL,
    -- Campos de auditoría (se llenan automáticamente)
    AddRecord DATETIME NULL,
    LastUpdate DATETIME NULL,
    CodiEmprNet INT NULL DEFAULT 0,
    NumUsuar INT NULL,
    
    -- Foreign key a PROFABRI (opcional, permite NULL para cálculos temporales)
    CONSTRAINT FK_CAPANECE_PROFABRI 
        FOREIGN KEY (NumeroPrograma) 
        REFERENCES PROFABRI(NumeroPrograma)
        ON DELETE CASCADE
);

-- Índice para búsqueda por equipo
CREATE INDEX IX_CAPANECE_Equipo ON CAPANECE(Equipo);

-- Índice para búsqueda por programa
CREATE INDEX IX_CAPANECE_NumeroPrograma ON CAPANECE(NumeroPrograma);

-- Índice para búsqueda por pieza
CREATE INDEX IX_CAPANECE_CodigoPieza ON CAPANECE(CodigoPieza);

-- Índice compuesto para búsquedas combinadas equipo-pieza
CREATE INDEX IX_CAPANECE_Equipo_Pieza ON CAPANECE(Equipo, CodigoPieza);

-- Índice compuesto para búsquedas por programa y equipo
CREATE INDEX IX_CAPANECE_Programa_Equipo ON CAPANECE(NumeroPrograma, Equipo);

*/
