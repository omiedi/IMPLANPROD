using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("NotificacionReglasSql")]
    public class NotificacionReglaSql : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Required]
        [MaxLength(200)]
        public string Titulo { get; set; } = string.Empty;

        [MaxLength(500)]
        public string? Resumen { get; set; }

        [Required]
        public bool Activa { get; set; }

        [Required]
        public int CreadorUsuarioId { get; set; }

        [Required]
        [Column(TypeName = "nvarchar(max)")]
        public string SqlQuery { get; set; } = string.Empty;

        [MaxLength(200)]
        public string? SqlConnectionName { get; set; }

        public int MaxFilasDetalle { get; set; }

        public int IntervaloMinutos { get; set; }

        public DateTime? UltimaEjecucion { get; set; }

        [MaxLength(64)]
        public string? UltimoResultadoHash { get; set; }

        [MaxLength(100)]
        public string? CampoSeguimiento { get; set; }

        [MaxLength(200)]
        public string? UltimoValorProcesado { get; set; }

        [Required]
        public int ModoSeguimiento { get; set; } = 0;

        public bool EsParaTodos { get; set; } = true;

        [MaxLength(2000)]
        public string? DestinatariosIds { get; set; }

        public DateTime? AddRecord { get; set; }

        public DateTime? LastUpdate { get; set; }

        public int? CodiEmprNet { get; set; }

        public int? NumUsuar { get; set; }
    }
}

/*
CREATE TABLE [dbo].[NotificacionReglasSql] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [Titulo] NVARCHAR(200) NOT NULL,
    [Resumen] NVARCHAR(500) NULL,
    [Activa] BIT NOT NULL CONSTRAINT [DF_NotificacionReglasSql_Activa] DEFAULT(1),
    [CreadorUsuarioId] INT NOT NULL,
    [SqlQuery] NVARCHAR(MAX) NOT NULL,
    [SqlConnectionName] NVARCHAR(200) NULL,
    [MaxFilasDetalle] INT NOT NULL CONSTRAINT [DF_NotificacionReglasSql_MaxFilasDetalle] DEFAULT(200),
    [IntervaloMinutos] INT NOT NULL CONSTRAINT [DF_NotificacionReglasSql_IntervaloMinutos] DEFAULT(15),
    [UltimaEjecucion] DATETIME2 NULL,
    [UltimoResultadoHash] NVARCHAR(64) NULL,
    [CampoSeguimiento] NVARCHAR(100) NULL,
    [UltimoValorProcesado] NVARCHAR(200) NULL,
    [ModoSeguimiento] INT NOT NULL CONSTRAINT [DF_NotificacionReglasSql_ModoSeguimiento] DEFAULT(0),
    [EsParaTodos] BIT NOT NULL CONSTRAINT [DF_NotificacionReglasSql_EsParaTodos] DEFAULT(1),
    [DestinatariosIds] NVARCHAR(2000) NULL,

    [AddRecord] DATETIME2 NULL,
    [LastUpdate] DATETIME2 NULL,
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,

    CONSTRAINT [PK_NotificacionReglasSql] PRIMARY KEY CLUSTERED ([Id] ASC)
);

CREATE NONCLUSTERED INDEX [IX_NotificacionReglasSql_Activa] ON [dbo].[NotificacionReglasSql]([Activa] ASC);
CREATE NONCLUSTERED INDEX [IX_NotificacionReglasSql_UltimaEjecucion] ON [dbo].[NotificacionReglasSql]([UltimaEjecucion] DESC);

-- ALTER TABLE para agregar nuevos campos a tablas existentes
ALTER TABLE [dbo].[NotificacionReglasSql] ADD [CampoSeguimiento] NVARCHAR(100) NULL;
ALTER TABLE [dbo].[NotificacionReglasSql] ADD [UltimoValorProcesado] NVARCHAR(200) NULL;
ALTER TABLE [dbo].[NotificacionReglasSql] ADD [ModoSeguimiento] INT NOT NULL DEFAULT(0);
ALTER TABLE [dbo].[NotificacionReglasSql] ADD [EsParaTodos] BIT NOT NULL DEFAULT(1);
ALTER TABLE [dbo].[NotificacionReglasSql] ADD [DestinatariosIds] NVARCHAR(2000) NULL;
*/
