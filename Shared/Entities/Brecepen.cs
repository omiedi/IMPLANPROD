using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;




namespace IMPLANPROD.Shared.Entities
{
    [Table("BRECEPEN")]
    public class Brecepen : IAuditableEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        [Column("FECHMOV")]
        public DateTime? Fechmov { get; set; }

        [Column("FEREMOVI")]
        public DateTime? Feremovi { get; set; }

        [Column("COD_INTE")]
        public int COD_INTE { get; set; }

        [Column("IDENLOTE")]
        [MaxLength(24)]
        public string? Idenlote { get; set; }

        [Column("CODIMOVI")]
        [MaxLength(6)]
        public string? Codimovi { get; set; }

        [Column("DOPRILAR")]
        [MaxLength(24)]
        public string? Doprilar { get; set; }

        [Column("DOSECLAR")]
        [MaxLength(24)]
        public string? Doseclar { get; set; }

        [Column("DOPRICOR")]
        [MaxLength(24)]
        public string? Dopricor { get; set; }

        [Column("DOSECCOR")]
        [MaxLength(24)]
        public string? Doseccor { get; set; }

        [Column("REFERCOR")]
        [MaxLength(64)]
        public string? Refercor { get; set; }

        [Column("VALOUNT", TypeName = "decimal(18,4)")]
        public decimal? Valount { get; set; }

        [Column("MONEVAL")]
        public short? Moneval { get; set; }

        [Column("DEPOMOVI")]
        public short? Depomovi { get; set; }

        [Column("NUME_CLIE")]
        public int? NumeClie { get; set; }

        [Column("NUME_PROV")]
        public int? NumeProv { get; set; }

        [Column("CANTINGRE", TypeName = "decimal(18,4)")]
        public decimal? Cantingre { get; set; }

        [Column("CANTSALID", TypeName = "decimal(18,4)")]
        public decimal? Cantsalid { get; set; }

        [Column("Clave_HP")]
        public int? ClaveHp { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// Se establece automáticamente cuando se crea el movimiento de stock
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// Se actualiza automáticamente cada vez que se modifica el movimiento
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// Se actualiza automáticamente con la empresa del usuario actual
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// Se actualiza automáticamente con el ID del usuario actual
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }

    }
}
/*
-- =============================================
--Script SQL Server: Crear tabla BRECEPEN
-- Descripción: Recepciones pendientes de aprobación (estructura solicitada) + auditoría
-- Fecha: 2026 - 03 - 02
-- =============================================

BEGIN
    CREATE TABLE[dbo].[BRECEPEN]
(
        [Id] INT IDENTITY(1,1) NOT NULL,

        [FECHMOV] DATE NULL,
        [FEREMOVI] DATETIME2(7) NULL,

        [COD_INTE] INT NOT NULL,
        [IDENLOTE] VARCHAR(24) NULL,
        [CODIMOVI] VARCHAR(6) NULL,

        [DOPRILAR] VARCHAR(24) NULL,
        [DOSECLAR] VARCHAR(24) NULL,
        [DOPRICOR] VARCHAR(24) NULL,
        [DOSECCOR] VARCHAR(24) NULL,

        [REFERCOR] VARCHAR(64) NULL,

        [VALOUNT] DECIMAL(18,4) NULL,
        [MONEVAL] SMALLINT NULL,
        [DEPOMOVI] SMALLINT NULL,

        [NUME_CLIE] INT NULL,
        [NUME_PROV] INT NULL,

        [CANTINGRE] DECIMAL(18,4) NULL,
        [CANTSALID] DECIMAL(18,4) NULL,

        [Clave_HP] INT NULL,

        -- =============================================================
        -- Auditoría (ITimestampEntity + IAuditableEntity)
        -- =============================================================
        [AddRecord] DATETIME2(7) NULL,
        [LastUpdate] DATETIME2(7) NULL,
        [CodiEmprNet] INT NULL,
        [NumUsuar] INT NULL,

        CONSTRAINT[PK_BRECEPEN] PRIMARY KEY CLUSTERED ([Id] ASC)
    );

PRINT 'Tabla BRECEPEN creada exitosamente';
END
ELSE
BEGIN
    PRINT 'La tabla BRECEPEN ya existe';
END
GO

-- Índices sugeridos (opcionales). Ajustar según consultas reales.
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_BRECEPEN_COD_INTE' AND object_id = OBJECT_ID('BRECEPEN'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BRECEPEN_COD_INTE] ON[dbo].[BRECEPEN]([COD_INTE] ASC);
PRINT 'Índice IX_BRECEPEN_COD_INTE creado';
END
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_BRECEPEN_FECHMOV' AND object_id = OBJECT_ID('BRECEPEN'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_BRECEPEN_FECHMOV] ON[dbo].[BRECEPEN]([FECHMOV] ASC);
PRINT 'Índice IX_BRECEPEN_FECHMOV creado';
END
GO
*/

