using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("Notificaciones")]
    public class Notificacion : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Required]
        [MaxLength(200)]
        public string Titulo { get; set; } = string.Empty;

        [MaxLength(500)]
        public string? Resumen { get; set; }

        public bool EsParaTodos { get; set; }

        [Required]
        public int CreadorUsuarioId { get; set; }

        [Column(TypeName = "nvarchar(max)")]
        public string? DetalleHtml { get; set; }

        [MaxLength(64)]
        public string? DetalleHash { get; set; }

        public DateTime? FechaExpiracion { get; set; }

        public bool EsGeneradaPorRegla { get; set; }

        public int? ReglaSqlId { get; set; }

        [ForeignKey(nameof(ReglaSqlId))]
        public NotificacionReglaSql? ReglaSql { get; set; }

        public DateTime? AddRecord { get; set; }

        public DateTime? LastUpdate { get; set; }

        public int? CodiEmprNet { get; set; }

        public int? NumUsuar { get; set; }
    }
}

/*
CREATE TABLE [dbo].[Notificaciones] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [Titulo] NVARCHAR(200) NOT NULL,
    [Resumen] NVARCHAR(500) NULL,
    [EsParaTodos] BIT NOT NULL CONSTRAINT [DF_Notificaciones_EsParaTodos] DEFAULT(0),
    [CreadorUsuarioId] INT NOT NULL,
    [DetalleHtml] NVARCHAR(MAX) NULL,
    [DetalleHash] NVARCHAR(64) NULL,
    [FechaExpiracion] DATETIME2 NULL,
    [EsGeneradaPorRegla] BIT NOT NULL CONSTRAINT [DF_Notificaciones_EsGeneradaPorRegla] DEFAULT(0),
    [ReglaSqlId] INT NULL,

    [AddRecord] DATETIME2 NULL,
    [LastUpdate] DATETIME2 NULL,
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,

    CONSTRAINT [PK_Notificaciones] PRIMARY KEY CLUSTERED ([Id] ASC)
);

CREATE NONCLUSTERED INDEX [IX_Notificaciones_AddRecord] ON [dbo].[Notificaciones]([AddRecord] DESC);
CREATE NONCLUSTERED INDEX [IX_Notificaciones_CreadorUsuarioId] ON [dbo].[Notificaciones]([CreadorUsuarioId] ASC);
CREATE NONCLUSTERED INDEX [IX_Notificaciones_ReglaSqlId] ON [dbo].[Notificaciones]([ReglaSqlId] ASC);
*/
