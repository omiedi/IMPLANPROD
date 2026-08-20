using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("NotificacionOcultas")]
    public class NotificacionOculta : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public int NotificacionId { get; set; }

        [Required]
        public int UsuarioId { get; set; }

        [Required]
        public DateTime FechaOculta { get; set; }

        public DateTime? AddRecord { get; set; }

        public DateTime? LastUpdate { get; set; }

        public int? CodiEmprNet { get; set; }

        public int? NumUsuar { get; set; }

        [ForeignKey(nameof(NotificacionId))]
        public Notificacion? Notificacion { get; set; }

        [ForeignKey(nameof(UsuarioId))]
        public Usuario? Usuario { get; set; }
    }
}

/*-- =========================================================
-- Tabla: NotificacionOcultas
-- Objetivo: "Borrar para mí" / ocultar notificaciones por usuario
-- =========================================================

IF NOT EXISTS (
    SELECT *
    FROM sys.objects
    WHERE object_id = OBJECT_ID(N'[dbo].[NotificacionOcultas]')
      AND type IN (N'U')
)
BEGIN
    CREATE TABLE [dbo].[NotificacionOcultas] (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [NotificacionId] INT NOT NULL,
        [UsuarioId] INT NOT NULL,
        [FechaOculta] DATETIME2 NOT NULL,

        [AddRecord] DATETIME2 NULL,
        [LastUpdate] DATETIME2 NULL,
        [CodiEmprNet] INT NULL,
        [NumUsuar] INT NULL,

        CONSTRAINT [PK_NotificacionOcultas] PRIMARY KEY CLUSTERED ([Id] ASC),

        CONSTRAINT [FK_NotificacionOcultas_Notificaciones]
            FOREIGN KEY ([NotificacionId]) REFERENCES [dbo].[Notificaciones]([Id])
            ON DELETE CASCADE,

        CONSTRAINT [FK_NotificacionOcultas_Usuarios]
            FOREIGN KEY ([UsuarioId]) REFERENCES [dbo].[Usuarios]([Id])
            ON DELETE NO ACTION
    );

    PRINT 'Tabla NotificacionOcultas creada exitosamente';
END
ELSE
BEGIN
    PRINT 'La tabla NotificacionOcultas ya existe';
END
GO

-- Índice único: una notificación solo puede estar oculta 1 vez por usuario
IF NOT EXISTS (
    SELECT *
    FROM sys.indexes
    WHERE name = 'UX_NotificacionOcultas_NotificacionId_UsuarioId'
      AND object_id = OBJECT_ID(N'[dbo].[NotificacionOcultas]')
)
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX [UX_NotificacionOcultas_NotificacionId_UsuarioId]
    ON [dbo].[NotificacionOcultas]([NotificacionId] ASC, [UsuarioId] ASC);

    PRINT 'Índice UX_NotificacionOcultas_NotificacionId_UsuarioId creado';
END
GO

-- Índice para listar ocultas por usuario (y ordenar/filtrar por fecha)
IF NOT EXISTS (
    SELECT *
    FROM sys.indexes
    WHERE name = 'IX_NotificacionOcultas_UsuarioId_FechaOculta'
      AND object_id = OBJECT_ID(N'[dbo].[NotificacionOcultas]')
)
BEGIN
    CREATE NONCLUSTERED INDEX [IX_NotificacionOcultas_UsuarioId_FechaOculta]
    ON [dbo].[NotificacionOcultas]([UsuarioId] ASC, [FechaOculta] DESC)
    INCLUDE ([NotificacionId]);

    PRINT 'Índice IX_NotificacionOcultas_UsuarioId_FechaOculta creado';
END
GO 
*/