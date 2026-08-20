using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using IMPLANPROD.Shared.Interfaces;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Tabla que almacena los tokens temporales para recuperación de contraseña.
    /// Cada token es de un solo uso y expira a los 30 minutos de generado.
    /// Flujo: usuario pide reset → se genera token → se envía por email → usuario usa link → token se marca como Usado.
    /// </summary>
    [Table("PasswordResetToken")]
    public class PasswordResetToken : ITimestampEntity, IAuditableEntity
    {
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// ID del usuario que solicitó el reset
        /// </summary>
        public int UsuarioId { get; set; }

        /// <summary>
        /// Token único (hex 64 chars) enviado en el link del email
        /// </summary>
        [Required]
        [MaxLength(128)]
        public string Token { get; set; } = string.Empty;

        /// <summary>
        /// Fecha/hora de expiración del token (UTC). Por defecto: 30 minutos desde la creación.
        /// </summary>
        public DateTime FechaExpiracion { get; set; }

        /// <summary>
        /// Indica si el token ya fue utilizado. Una vez usado no puede reutilizarse.
        /// </summary>
        public bool Usado { get; set; } = false;

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }

    }
}

/*
-- =====================================================================
-- SCRIPT DE CREACIÓN: PasswordResetToken
-- Sistema: IMPLANPROD - Recuperación de contraseña por email
-- =====================================================================

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[PasswordResetToken]') AND type IN (N'U'))
BEGIN
    CREATE TABLE [dbo].[PasswordResetToken]
    (
        [Id]              INT           IDENTITY(1,1) NOT NULL,
        [UsuarioId]       INT           NOT NULL,
        [Token]           NVARCHAR(128) NOT NULL,
        [FechaExpiracion] DATETIME2     NOT NULL,
        [Usado]           BIT           NOT NULL DEFAULT 0,

        -- Campos de auditoría automática (AuditInterceptor)
        [AddRecord]       DATETIME2     NULL,
        [LastUpdate]      DATETIME2     NULL,
        [CodiEmprNet]     INT           NULL DEFAULT 0,
        [NumUsuar]        INT           NULL,

        CONSTRAINT [PK_PasswordResetToken] PRIMARY KEY CLUSTERED ([Id] ASC)
    );

    PRINT 'Tabla PasswordResetToken creada exitosamente.';
END
ELSE
BEGIN
    PRINT 'La tabla PasswordResetToken ya existe.';
END
GO

-- Índice sobre Token (búsqueda directa al validar el link del email)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_PasswordResetToken_Token' AND object_id = OBJECT_ID('PasswordResetToken'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_PasswordResetToken_Token]
        ON [dbo].[PasswordResetToken] ([Token] ASC);
    PRINT 'Índice IX_PasswordResetToken_Token creado.';
END
GO

-- Índice sobre UsuarioId (para invalidar tokens anteriores del mismo usuario)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_PasswordResetToken_UsuarioId' AND object_id = OBJECT_ID('PasswordResetToken'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_PasswordResetToken_UsuarioId]
        ON [dbo].[PasswordResetToken] ([UsuarioId] ASC, [Usado] ASC, [FechaExpiracion] ASC);
    PRINT 'Índice IX_PasswordResetToken_UsuarioId creado.';
END
GO

PRINT 'Script PasswordResetToken finalizado.';
GO
*/
