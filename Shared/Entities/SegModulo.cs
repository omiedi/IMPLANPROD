using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("SEG_MODULO")]
    public class SegModulo : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Required]
        [MaxLength(80)]
        public string Codigo { get; set; } = string.Empty;

        [Required]
        [MaxLength(150)]
        public string Nombre { get; set; } = string.Empty;

        [MaxLength(300)]
        public string? Descripcion { get; set; }

        public int Orden { get; set; }

        public bool Activo { get; set; } = true;

        [Column("Codiempr")]
        public int? Codiempr { get; set; }

        public DateTime? AddRecord { get; set; }

        public DateTime? LastUpdate { get; set; }

        public int? CodiEmprNet { get; set; }

        public int? NumUsuar { get; set; }

        public List<SegPagina> Paginas { get; set; } = new();
    }
}
/* 
 * IF OBJECT_ID('dbo.SEG_MODULO', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.SEG_MODULO
        (
            Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
            Codigo NVARCHAR(80) NOT NULL,
            Nombre NVARCHAR(150) NOT NULL,
            Descripcion NVARCHAR(300) NULL,
            Orden INT NOT NULL CONSTRAINT DF_SEG_MODULO_Orden DEFAULT(0),
            Activo BIT NOT NULL CONSTRAINT DF_SEG_MODULO_Activo DEFAULT(1),
            Codiempr INT NOT NULL CONSTRAINT DF_SEG_MODULO_Codiempr DEFAULT(0),
            AddRecord DATETIME2 NULL CONSTRAINT DF_SEG_MODULO_AddRecord DEFAULT(SYSDATETIME()),
            LastUpdate DATETIME2 NULL CONSTRAINT DF_SEG_MODULO_LastUpdate DEFAULT(SYSDATETIME()),
            CodiEmprNet INT NULL CONSTRAINT DF_SEG_MODULO_CodiEmprNet DEFAULT(0),
            NumUsuar INT NULL CONSTRAINT DF_SEG_MODULO_NumUsuar DEFAULT(0)
        );

        CREATE UNIQUE INDEX UX_SEG_MODULO_Codigo ON dbo.SEG_MODULO(Codigo);
        CREATE INDEX IX_SEG_MODULO_Orden ON dbo.SEG_MODULO(Orden);
    END;
*/
