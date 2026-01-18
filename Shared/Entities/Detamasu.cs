using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("DETAMASU")]
    public class Detamasu : IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [MaxLength(4)]
        public string? CodiMovi { get; set; }

        public int? Cod_Inte { get; set; }

        [MaxLength(64)]
        public string? Cod_Exte { get; set; }

        public int? Talon { get; set; }

        public int? NReparación { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantIngre { get; set; }

        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantSalid { get; set; }

        public DateTime? FechMov { get; set; }

        public DateTime? FeRemovi { get; set; }

        public int? Nume_Clie { get; set; }

        [MaxLength(24)]
        public string? NumRemito { get; set; }

        [MaxLength(128)]
        public string? Observaciones { get; set; }

        [MaxLength(128)]
        public string? Referencia { get; set; }

        public DateTime? AddRecord { get; set; } = DateTime.Now;

        public DateTime? LastUpdate { get; set; } = DateTime.Now;

        public int? CodiEmprNet { get; set; } = 101;

        public int? NumUsuar { get; set; } = 10;
    }
}

/*

SCRIPT SQL - CREACIÓN DE TABLA DETAMASU (SQL Server)

CREATE TABLE [dbo].[DETAMASU] (
    [Id] INT IDENTITY(1,1) NOT NULL,

    [CodiMovi] VARCHAR(4) NULL,        -- Código de movimiento
    [Cod_Inte] INT NULL,
    [Cod_Exte] VARCHAR(64) NULL,
    [Talon] INT NULL,
    [NReparación] INT NULL,

    [CantIngre] DECIMAL(18,4) NULL,
    [CantSalid] DECIMAL(18,4) NULL,

    [FechMov] DATETIME NULL,
    [FeRemovi] DATETIME NULL,

    [Nume_Clie] INT NULL,
    [NumRemito] VARCHAR(24) NULL,
    [Observaciones] VARCHAR(128) NULL,
    [Referencia] VARCHAR(128) NULL,

    -- Auditoría
    [AddRecord] DATETIME NULL CONSTRAINT [DF_DETAMASU_AddRecord] DEFAULT (GETDATE()),
    [LastUpdate] DATETIME NULL CONSTRAINT [DF_DETAMASU_LastUpdate] DEFAULT (GETDATE()),
    [CodiEmprNet] INT NULL CONSTRAINT [DF_DETAMASU_CodiEmprNet] DEFAULT (101),
    [NumUsuar] INT NULL CONSTRAINT [DF_DETAMASU_NumUsuar] DEFAULT (10),

    CONSTRAINT [PK_DETAMASU] PRIMARY KEY CLUSTERED ([Id] ASC)
);

GO

*/
