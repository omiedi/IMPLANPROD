using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    [Table("NotificacionDestinatarios")]
    public class NotificacionDestinatario : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        [Required]
        public int NotificacionId { get; set; }

        [Required]
        public int UsuarioId { get; set; }

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
CREATE TABLE [dbo].[NotificacionDestinatarios] (
    [Id] INT IDENTITY(1,1) NOT NULL,
    [NotificacionId] INT NOT NULL,
    [UsuarioId] INT NOT NULL,

    [AddRecord] DATETIME2 NULL,
    [LastUpdate] DATETIME2 NULL,
    [CodiEmprNet] INT NULL,
    [NumUsuar] INT NULL,

    CONSTRAINT [PK_NotificacionDestinatarios] PRIMARY KEY CLUSTERED ([Id] ASC),
    CONSTRAINT [FK_NotificacionDestinatarios_Notificaciones] FOREIGN KEY ([NotificacionId]) REFERENCES [dbo].[Notificaciones]([Id]),
    CONSTRAINT [FK_NotificacionDestinatarios_Usuarios] FOREIGN KEY ([UsuarioId]) REFERENCES [dbo].[Usuarios]([Id])
);

CREATE UNIQUE NONCLUSTERED INDEX [IX_NotificacionDestinatarios_NotificacionId_UsuarioId]
ON [dbo].[NotificacionDestinatarios]([NotificacionId] ASC, [UsuarioId] ASC);

CREATE NONCLUSTERED INDEX [IX_NotificacionDestinatarios_UsuarioId]
ON [dbo].[NotificacionDestinatarios]([UsuarioId] ASC)
INCLUDE ([NotificacionId]);
*/
