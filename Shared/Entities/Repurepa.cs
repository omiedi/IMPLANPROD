using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad REPUREPA - Repuestos por Reparación (Presupuesto/Recepción)
    /// 
    /// Objetivo:
    /// - Guardar repuestos asociados a una reparación (NUMREPA)
    /// - Debe poder administrarse de forma independiente del resto de la ficha
    /// - Incluye campos de auditoría (AuditInterceptor)
    /// </summary>
    [Table("REPUREPA")]
    public class Repurepa : IAuditableEntity
    {
        /// <summary>
        /// Identidad autoincremental
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Número de reparación / presupuesto activo
        /// </summary>
        public int NumRepa { get; set; }

        /// <summary>
        /// Código interno del producto (MASTER.CODINT)
        /// </summary>
        public int Cod_Inte { get; set; }

        /// <summary>
        /// Cantidad solicitada
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? Cant_Sol { get; set; }

        /// <summary>
        /// Precio unitario neto
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PrUn_Neto { get; set; }

        /// <summary>
        /// Marca de borrado (borrado lógico)
        /// </summary>
        public short? MarBorra { get; set; }

        /// <summary>
        /// Importe neto del ítem
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? Impo_Neto { get; set; }

        /// <summary>
        /// Número de orden del ítem (secuencia)
        /// </summary>
        public short? Nuord_Item { get; set; }

        /// <summary>
        /// Identificación de lote
        /// </summary>
        [StringLength(32)]
        public string? IdenLote { get; set; }

        /// <summary>
        /// Origen del registro (según VB6: se filtra ORIGEN = 0)
        /// </summary>
        public int? Origen { get; set; }

        /// <summary>
        /// Porcentaje de descuento aplicado al ítem
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PORDESC { get; set; }

        /// <summary>
        /// Marca para identificar ítems agregados manualmente desde la UI
        /// </summary>
        public short? MarcaItemAgregado { get; set; }

        // Campos de auditoría (IAuditableEntity)
        public DateTime? AddRecord { get; set; }
        public DateTime? LastUpdate { get; set; }
        public int? CodiEmprNet { get; set; }
        public int? NumUsuar { get; set; }

        /// <summary>
        /// Campo histórico usado en otras tablas del sistema (se mantiene por consistencia)
        /// </summary>
        public int? CodiEmpr { get; set; }
    }
}

/*
-- Script SQL para crear la tabla REPUREPA en SQL Server
-- IMPORTANTE: Ajustar el esquema (dbo) si su base usa otro.

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[REPUREPA]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[REPUREPA]
    (
        [Id] INT IDENTITY(1,1) NOT NULL,
        [NumRepa] INT NOT NULL,
        [Cod_Inte] INT NOT NULL,
        [Cant_Sol] DECIMAL(18,4) NULL,
        [PrUn_Neto] DECIMAL(18,4) NULL,
        [MarBorra] SMALLINT NULL,
        [Impo_Neto] DECIMAL(18,4) NULL,
        [Nuord_Item] SMALLINT NULL,
        [IdenLote] NVARCHAR(32) NULL,
        [Origen] INT NULL,
        [PORDESC] DECIMAL(18,4) NULL,
        [MarcaItemAgregado] SMALLINT NULL,

        -- Campos de auditoría automática (AuditInterceptor)
        [AddRecord] DATETIME NULL,
        [LastUpdate] DATETIME NULL,
        [CodiEmprNet] INT NULL,
        [NumUsuar] INT NULL,
        [CodiEmpr] INT NULL,

        CONSTRAINT [PK_REPUREPA] PRIMARY KEY CLUSTERED ([Id] ASC)
    );
END
GO

-- Índice principal de búsqueda: NumRepa
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_REPUREPA_NumRepa' AND object_id = OBJECT_ID('REPUREPA'))
    CREATE NONCLUSTERED INDEX [IX_REPUREPA_NumRepa] ON [dbo].[REPUREPA] ([NumRepa] ASC);
GO

-- Índice útil para búsquedas por (NumRepa + Origen)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_REPUREPA_NumRepa_Origen' AND object_id = OBJECT_ID('REPUREPA'))
    CREATE NONCLUSTERED INDEX [IX_REPUREPA_NumRepa_Origen] ON [dbo].[REPUREPA] ([NumRepa] ASC, [Origen] ASC);
GO

-- Índice útil para búsquedas por NumRepa + Cod_Inte
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_REPUREPA_NumRepa_CodInte' AND object_id = OBJECT_ID('REPUREPA'))
    CREATE NONCLUSTERED INDEX [IX_REPUREPA_NumRepa_CodInte] ON [dbo].[REPUREPA] ([NumRepa] ASC, [Cod_Inte] ASC);
GO
*/
