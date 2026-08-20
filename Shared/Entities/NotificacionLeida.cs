using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("NotificacionLeidas")]
    public class NotificacionLeida : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public int NotificacionId { get; set; }

        [Required]
        public int UsuarioId { get; set; }

        [Required]
        public DateTime FechaLeida { get; set; }

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

/*
CREATE TABLE [dbo].[NotificacionLeidas] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [NotificacionId] INT NOT NULL,
    [UsuarioId] INT NOT NULL,
    [FechaLeida] DATETIME2 NOT NULL,

    [AddRecord] DATETIME2 NULL,
    [LastUpdate] DATETIME2 NULL,
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,

    CONSTRAINT [PK_NotificacionLeidas] PRIMARY KEY CLUSTERED ([Id] ASC),
    CONSTRAINT [FK_NotificacionLeidas_Notificaciones] FOREIGN KEY ([NotificacionId]) REFERENCES [dbo].[Notificaciones]([Id]),
    CONSTRAINT [FK_NotificacionLeidas_Usuarios] FOREIGN KEY ([UsuarioId]) REFERENCES [dbo].[Usuarios]([Id])
);

CREATE UNIQUE NONCLUSTERED INDEX [IX_NotificacionLeidas_NotificacionId_UsuarioId]
ON [dbo].[NotificacionLeidas]([NotificacionId] ASC, [UsuarioId] ASC);

CREATE NONCLUSTERED INDEX [IX_NotificacionLeidas_UsuarioId_FechaLeida]
ON [dbo].[NotificacionLeidas]([UsuarioId] ASC, [FechaLeida] DESC)
INCLUDE ([NotificacionId]);
*/
