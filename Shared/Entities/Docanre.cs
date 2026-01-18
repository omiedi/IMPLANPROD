using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("Docanre")]
    public class Docanre : IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        public int CodiEmpr { get; set; } = 0;

        public int NumRepa { get; set; }

        [MaxLength(256)]
        public string NomAsignado { get; set; } = string.Empty;

        [MaxLength(256)]
        public string? NomOriginal { get; set; }

        [MaxLength(500)]
        public string? RutaDocumento { get; set; }

        [MaxLength(100)]
        public string? ContentType { get; set; }

        public long? TamanoBytes { get; set; }

        [MaxLength(10)]
        public string? Extension { get; set; }

        public DateTime? AddRecord { get; set; }

        public DateTime? LastUpdate { get; set; }

        public int? CodiEmprNet { get; set; }

        public int? NumUsuar { get; set; }

        [ForeignKey(nameof(NumRepa))]
        public Repara? Repara { get; set; }
    }
}

/*
-- Script SQL para crear la tabla DOCANRE en SQL Server

CREATE TABLE [dbo].[DOCANRE] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [CodiEmpr] INT NOT NULL,
    [NumRepa] INT NOT NULL,
    [NomAsignado] NVARCHAR(256) NOT NULL,
    [NomOriginal] NVARCHAR(256) NULL,
    [RutaDocumento] NVARCHAR(500) NULL,
    [ContentType] NVARCHAR(100) NULL,
    [TamanoBytes] BIGINT NULL,
    [Extension] NVARCHAR(10) NULL,

    -- Campos de auditoría automática (asignados por AuditInterceptor)
    [AddRecord] DATETIME NULL,
    [LastUpdate] DATETIME NULL,
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,

    CONSTRAINT [PK_DOCANRE] PRIMARY KEY CLUSTERED ([Id] ASC)
);

-- Índice para búsquedas por NumRepa (consulta principal)
CREATE NONCLUSTERED INDEX [IX_DOCANRE_NumRepa]
    ON [dbo].[DOCANRE] ([NumRepa] ASC);

-- Índice para búsquedas por empresa
CREATE NONCLUSTERED INDEX [IX_DOCANRE_CodiEmpr]
    ON [dbo].[DOCANRE] ([CodiEmpr] ASC);

-- Índice para evitar duplicados de nombre asignado dentro de la misma reparación
CREATE UNIQUE NONCLUSTERED INDEX [IX_DOCANRE_NumRepa_NomAsignado]
    ON [dbo].[DOCANRE] ([NumRepa] ASC, [NomAsignado] ASC);

GO
*/
