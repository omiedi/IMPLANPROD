using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Trazabilidad de Recepciones
    /// </summary>
    [Table("TRAZAREC")]
    public class Trazarec : ITimestampEntity, IAuditableEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("ID")]
        public int Id { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Column("CodiEmpr")]
        public int CodiEmpr { get; set; } = 0;

        /// <summary>
        /// Número de remito
        /// </summary>
        [Column("NROREMI")]
        public int? NroRemi { get; set; }

        /// <summary>
        /// Fecha de recepción
        /// </summary>
        [Column("Fech_Recep")]
        public DateTime? FechRecep { get; set; }

        /// <summary>
        /// Número de orden de compra
        /// </summary>
        [Column("Nume_Ocom")]
        public int? NumeOcom { get; set; }

        /// <summary>
        /// Cantidad remitida
        /// </summary>
        [Column("Cant_Remi")]
        public decimal? CantRemi { get; set; }

        /// <summary>
        /// Código interno del producto
        /// </summary>
        [Column("Cod_Inte")]
        public int? CodInte { get; set; }

        /// <summary>
        /// Número de pedido
        /// </summary>
        [Column("NroPed")]
        public int? NroPed { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}
/*
-- =============================================
-- Creación de tabla TRAZAREC
-- Trazabilidad de Recepciones
-- =============================================

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TRAZAREC]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[TRAZAREC](
        [ID] [int] IDENTITY(1,1) NOT NULL,
        [CodiEmpr] [int] NOT NULL DEFAULT 0,
        [NROREMI] [int] NULL,
        [Fech_Recep] [smalldatetime] NULL,
        [Nume_Ocom] [int] NULL,
        [Cant_Remi] [decimal](18, 4) NULL,
        [Cod_Inte] [int] NULL,
        [NroPed] [int] NULL,
        [AddRecord] [datetime] NULL,
        [LastUpdate] [datetime] NULL,
        [CodiEmprNet] [int] NULL,
        [NumUsuar] [int] NULL,
     CONSTRAINT [PK_TRAZAREC] PRIMARY KEY CLUSTERED 
    (
        [ID] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
    ) ON [PRIMARY]
END
GO

-- =============================================
-- Índices para TRAZAREC
-- =============================================

-- Índice en NROREMI (Número de Remito)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_NROREMI' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_NROREMI] ON [dbo].[TRAZAREC]
    (
        [NROREMI] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en Cod_Inte (Código Interno del Producto)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_Cod_Inte' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_Cod_Inte] ON [dbo].[TRAZAREC]
    (
        [Cod_Inte] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en Nume_Ocom (Número de Orden de Compra)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_Nume_Ocom' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_Nume_Ocom] ON [dbo].[TRAZAREC]
    (
        [Nume_Ocom] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en NroPed (Número de Pedido)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_NroPed' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_NroPed] ON [dbo].[TRAZAREC]
    (
        [NroPed] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice compuesto en NROREMI + Cod_Inte (para búsquedas específicas de producto en un remito)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_NROREMI_Cod_Inte' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_NROREMI_Cod_Inte] ON [dbo].[TRAZAREC]
    (
        [NROREMI] ASC,
        [Cod_Inte] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en Fech_Recep (Fecha de Recepción)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_Fech_Recep' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_Fech_Recep] ON [dbo].[TRAZAREC]
    (
        [Fech_Recep] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en CodiEmpr (Código de Empresa)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_CodiEmpr' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_CodiEmpr] ON [dbo].[TRAZAREC]
    (
        [CodiEmpr] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO
-- =============================================
-- Creación de tabla TRAZAREC
-- Trazabilidad de Recepciones
-- =============================================

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[TRAZAREC]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[TRAZAREC](
        [ID] [int] IDENTITY(1,1) NOT NULL,
        [CodiEmpr] [int] NOT NULL DEFAULT 0,
        [NROREMI] [int] NULL,
        [Fech_Recep] [smalldatetime] NULL,
        [Nume_Ocom] [int] NULL,
        [Cant_Remi] [decimal](18, 4) NULL,
        [Cod_Inte] [int] NULL,
        [NroPed] [int] NULL,
        [AddRecord] [datetime] NULL,
        [LastUpdate] [datetime] NULL,
        [CodiEmprNet] [int] NULL,
        [NumUsuar] [int] NULL,
     CONSTRAINT [PK_TRAZAREC] PRIMARY KEY CLUSTERED 
    (
        [ID] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
    ) ON [PRIMARY]
END
GO
/*
-- =============================================
-- Índices para TRAZAREC
-- =============================================

-- Índice en NROREMI (Número de Remito)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_NROREMI' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_NROREMI] ON [dbo].[TRAZAREC]
    (
        [NROREMI] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en Cod_Inte (Código Interno del Producto)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_Cod_Inte' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_Cod_Inte] ON [dbo].[TRAZAREC]
    (
        [Cod_Inte] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en Nume_Ocom (Número de Orden de Compra)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_Nume_Ocom' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_Nume_Ocom] ON [dbo].[TRAZAREC]
    (
        [Nume_Ocom] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en NroPed (Número de Pedido)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_NroPed' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_NroPed] ON [dbo].[TRAZAREC]
    (
        [NroPed] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice compuesto en NROREMI + Cod_Inte (para búsquedas específicas de producto en un remito)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_NROREMI_Cod_Inte' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_NROREMI_Cod_Inte] ON [dbo].[TRAZAREC]
    (
        [NROREMI] ASC,
        [Cod_Inte] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en Fech_Recep (Fecha de Recepción)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_Fech_Recep' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_Fech_Recep] ON [dbo].[TRAZAREC]
    (
        [Fech_Recep] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO

-- Índice en CodiEmpr (Código de Empresa)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_TRAZAREC_CodiEmpr' AND object_id = OBJECT_ID('[dbo].[TRAZAREC]'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_TRAZAREC_CodiEmpr] ON [dbo].[TRAZAREC]
    (
        [CodiEmpr] ASC
    )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
END
GO
*/