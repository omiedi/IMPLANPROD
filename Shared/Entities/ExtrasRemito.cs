using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad EXTRAS_REMITO - Datos adicionales asociados a un remito.
    /// </summary>
    [Table("EXTRAS_REMITO")]
    public class ExtrasRemito : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador unico autoincremental
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Numero de cliente asociado al remito
        /// </summary>
        [Column("NUME_CLIE")]
        public int? NumeClie { get; set; }

        /// <summary>
        /// Fecha del movimiento
        /// </summary>
        [Column("FECHMOV", TypeName = "date")]
        public DateTime? Fechmov { get; set; }

        /// <summary>
        /// DoPriLar del remito (numero largo)
        /// </summary>
        [Column("DOPRILAR")]
        [MaxLength(24)]
        public string? Doprilar { get; set; }

        /// <summary>
        /// Observaciones del remito (columna OBSERVACIONNES segun requerimiento)
        /// </summary>
        [Column("OBSERVACIONNES")]
        [MaxLength(256)]
        public string? Observaciones { get; set; }

        /// <summary>
        /// Cantidad de bultos
        /// </summary>
        [Column("BULTOS")]
        public int? Bultos { get; set; }

        /// <summary>
        /// Valor declarado del remito (4 decimales)
        /// </summary>
        [Column("VALOR_DECLARADO", TypeName = "decimal(18,4)")]
        public decimal? ValorDeclarado { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMATICOS - Implementacion de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creacion del registro
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de ultima actualizacion del registro
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Codigo de empresa - Implementacion de IAuditableEntity.CodiEmprNet
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Numero de usuario - Implementacion de IAuditableEntity.NumUsuar
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}
/*
-- =============================================
-- Script SQL Server: Crear tabla EXTRAS_REMITO
-- Descripcion: Datos adicionales de remitos
-- Fecha: 2026-06-01
-- =============================================

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'EXTRAS_REMITO' AND type = 'U')
BEGIN
    CREATE TABLE [dbo].[EXTRAS_REMITO]
    (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [NUME_CLIE] INT NULL,
        [FECHMOV] DATE NULL,
        [DOPRILAR] VARCHAR(24) NULL,
        [OBSERVACIONNES] VARCHAR(256) NULL,
        [BULTOS] INT NULL,
        [VALOR_DECLARADO] DECIMAL(18,4) NULL,

        -- Auditoria (ITimestampEntity + IAuditableEntity)
        [AddRecord] DATETIME2(7) NULL,
        [LastUpdate] DATETIME2(7) NULL,
        [CodiEmprNet] INT NULL,
        [NumUsuar] INT NULL,

        CONSTRAINT [PK_EXTRAS_REMITO] PRIMARY KEY CLUSTERED ([Id] ASC)
    );

    PRINT 'Tabla EXTRAS_REMITO creada exitosamente';
END
ELSE
BEGIN
    PRINT 'La tabla EXTRAS_REMITO ya existe';
END
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_EXTRAS_REMITO_NUME_CLIE_DOPRILAR' AND object_id = OBJECT_ID('EXTRAS_REMITO'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_EXTRAS_REMITO_NUME_CLIE_DOPRILAR]
    ON [dbo].[EXTRAS_REMITO] ([NUME_CLIE] ASC, [DOPRILAR] ASC);
    PRINT 'Indice IX_EXTRAS_REMITO_NUME_CLIE_DOPRILAR creado';
END
GO
*/
