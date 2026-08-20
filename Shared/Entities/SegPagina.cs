using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("SEG_PAGINA")]
    public class SegPagina : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public int ModuloId { get; set; }

        [ForeignKey(nameof(ModuloId))]
        public SegModulo? Modulo { get; set; }

        [Required]
        [MaxLength(150)]
        public string Nombre { get; set; } = string.Empty;

        [Required]
        [MaxLength(200)]
        public string Ruta { get; set; } = string.Empty;

        [MaxLength(160)]
        public string? Componente { get; set; }

        public bool EsConfiguracion { get; set; }

        public int Orden { get; set; }

        public bool Activo { get; set; } = true;

        [Column("Codiempr")]
        public int? Codiempr { get; set; }

        public DateTime? AddRecord { get; set; }

        public DateTime? LastUpdate { get; set; }

        public int? CodiEmprNet { get; set; }

        public int? NumUsuar { get; set; }

        public List<SegUsuarioPagina> UsuariosPermitidos { get; set; } = new();
    }
}

/*
 * IF OBJECT_ID('dbo.SEG_PAGINA', 'U') IS NULL
    BEGIN
        CREATE TABLE dbo.SEG_PAGINA
        (
            Id INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
            ModuloId INT NOT NULL,
            Nombre NVARCHAR(150) NOT NULL,
            Ruta NVARCHAR(200) NOT NULL,
            Componente NVARCHAR(160) NULL,
            EsConfiguracion BIT NOT NULL CONSTRAINT DF_SEG_PAGINA_EsConfiguracion DEFAULT(0),
            Orden INT NOT NULL CONSTRAINT DF_SEG_PAGINA_Orden DEFAULT(0),
            Activo BIT NOT NULL CONSTRAINT DF_SEG_PAGINA_Activo DEFAULT(1),
            Codiempr INT NOT NULL CONSTRAINT DF_SEG_PAGINA_Codiempr DEFAULT(0),
            AddRecord DATETIME2 NULL CONSTRAINT DF_SEG_PAGINA_AddRecord DEFAULT(SYSDATETIME()),
            LastUpdate DATETIME2 NULL CONSTRAINT DF_SEG_PAGINA_LastUpdate DEFAULT(SYSDATETIME()),
            CodiEmprNet INT NULL CONSTRAINT DF_SEG_PAGINA_CodiEmprNet DEFAULT(0),
            NumUsuar INT NULL CONSTRAINT DF_SEG_PAGINA_NumUsuar DEFAULT(0),
            CONSTRAINT FK_SEG_PAGINA_SEG_MODULO FOREIGN KEY (ModuloId) REFERENCES dbo.SEG_MODULO(Id)
        );

        CREATE UNIQUE INDEX UX_SEG_PAGINA_Ruta ON dbo.SEG_PAGINA(Ruta);
        CREATE INDEX IX_SEG_PAGINA_Modulo_Orden ON dbo.SEG_PAGINA(ModuloId, Orden);
    END;

*/