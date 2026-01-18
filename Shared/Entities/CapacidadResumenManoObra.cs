using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Registro del resumen de mano de obra generado al calcular la capacidad.
    /// Equivalente a la tabla VB6 RECALMOD.
    /// </summary>
    [Table("RECALMOD")]
    public class CapacidadResumenManoObra : IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        /// <summary>Categoría de operario (código entero)</summary>
        public int Categoria { get; set; }

        /// <summary>Descripción de la categoría (48 caracteres)</summary>
        [MaxLength(48)]
        public string DescripcionCategoria { get; set; } = string.Empty;

        /// <summary>Horas necesarias para la categoría</summary>
        public decimal HorasNecesarias { get; set; }

        /// <summary>Total de horas por categoría (acumulado)</summary>
        public decimal TotalHorasPorCategoria { get; set; }

        /// <summary>Identificador del programa al que pertenece</summary>
        public int NumeroPrograma { get; set; }

        /// <summary>Marca de auditoría</summary>
        public DateTime FechaCreacion { get; set; } = DateTime.Now;

        /// Número de usuario
        /// </summary>
        public int? NumUsuar { get; set; }
        /// <summary>
        /// Código de empresa
        /// </summary>
        public int? CodiEmprNet { get; set; } = 0;

        [Display(Name = "Fecha de Creación")]
        public DateTime? AddRecord { get; set; }

        [Display(Name = "Última Actualización")]
        public DateTime? LastUpdate { get; set; }
    }
}
/*CREATE TABLE RECALMOD (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Categoria INT NOT NULL,
    DescripcionCategoria VARCHAR(48) NULL,
    HorasNecesarias DECIMAL(18,4) NULL,
    TotalHorasPorCategoria DECIMAL(18,4) NULL,
    NumeroPrograma INT NOT NULL,
    FechaCreacion DATETIME NOT NULL DEFAULT GETDATE(),
    AddRecord DATETIME NULL,
    LastUpdate DATETIME NULL,
    CodiEmprNet INT NULL DEFAULT 0,
    NumUsuar INT NULL

);*/



