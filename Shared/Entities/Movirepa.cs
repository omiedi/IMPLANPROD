using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad Movirepa (Movimientos de Reparación) con auditoría completa automática
    /// Tiene la misma estructura que Movisto
    /// Implementa IAuditableEntity para timestamps y seguimiento de usuario/empresa
    /// </summary>
    [Table("movirepa")]
    public class Movirepa : IAuditableEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("Id")]
        public int Id { get; set; }

        [Column("FeReMovi")]
        public DateTime? FeReMovi { get; set; }

        /// <summary>
        /// Código de empresa 
        /// </summary>
        [Column("CodiEmpr")]
        public int? CodiEmpr { get; set; }

        /// <summary>
        /// Número de usuario - Implementación de IAuditableEntity.NumUsuar
        /// Se actualiza automáticamente con el ID del usuario actual
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }

        [Column("IdenMovi")]
        [StringLength(32)]
        public string? IdenMovi { get; set; }

        [Column("DeLaRemi")]
        public string? DeLaRemi { get; set; }

        [Column("Descri_Lar")]
        public string? Descri_Lar { get; set; }

        [Column("FechMov")]
        public DateTime? FechMov { get; set; }

        [Column("Cod_Inte")]
        public int? Cod_Inte { get; set; }

        [Column("Cod_Exte")]
        [StringLength(24)]
        public string? Cod_Exte { get; set; }

        [Column("Descrip")]
        [StringLength(64)]
        public string? Descrip { get; set; }

        [Column("IdenLote")]
        [StringLength(16)]
        public string? IdenLote { get; set; }

        [Column("LotConsum")]
        public string? LotConsum { get; set; }

        [Column("PNL_Des")]
        [StringLength(32)]
        public string? PNL_Des { get; set; }

        [Column("PLC_Has")]
        [StringLength(32)]
        public string? PLC_Has { get; set; }

        [Column("CodiMovi")]
        [StringLength(2)]
        public string? CodiMovi { get; set; }

        [Column("DoPriCor")]
        [StringLength(12)]
        public string? DoPriCor { get; set; }

        [Column("DoPriLar")]
        [StringLength(24)]
        public string? DoPriLar { get; set; }

        [Column("DoSecCor")]
        [StringLength(12)]
        public string? DoSecCor { get; set; }

        [Column("DoSecLar")]
        [StringLength(24)]
        public string? DoSecLar { get; set; }

        [Column("DocuOrig")]
        [StringLength(24)]
        public string? DocuOrig { get; set; }

        [Column("NumRem_Cli")]
        [StringLength(16)]
        public string? NumRem_Cli { get; set; }

        [Column("NuorItem")]
        public int? NuorItem { get; set; }

        [Column("NuPedCli")]
        public int? NuPedCli { get; set; }

        [Column("NumeOrdPed")]
        public short? NumeOrdPed { get; set; }

        [Column("ReferCor")]
        [StringLength(64)]
        public string? ReferCor { get; set; }

        [Column("CostoRep")]
        public decimal? CostoRep { get; set; }

        [Column("ValoUnit")]
        public decimal? ValoUnit { get; set; }

        [Column("MoneVal")]
        public short? MoneVal { get; set; }

        [Column("TipoCam")]
        public decimal? TipoCam { get; set; }

        [Column("DepoMovi")]
        public int? DepoMovi { get; set; }

        [Column("Nume_Clie")]
        public int? Nume_Clie { get; set; }

        [Column("Sucu_Clie")]
        public short? Sucu_Clie { get; set; }

        [Column("Nume_Prov")]
        public int? Nume_Prov { get; set; }

        [Column("CantIngre")]
        public decimal? CantIngre { get; set; }

        [Column("CantSalid")]
        public decimal? CantSalid { get; set; }

        [Column("CantSaldo")]
        public decimal? CantSaldo { get; set; }

        [Column("Verificado")]
        public short? Verificado { get; set; }

        [Column("Causa")]
        public short? Causa { get; set; }

        [Column("User_Tran")]
        [StringLength(32)]
        public string? User_Tran { get; set; }

        [Column("CodCarpImp")]
        [StringLength(12)]
        public string? CodCarpImp { get; set; }

        [Column("ValoUnitIva")]
        public decimal? ValoUnitIva { get; set; }

        [Column("VerificadoPor")]
        public short? VerificadoPor { get; set; }

        [Column("FHor_Verif")]
        public DateTime? FHor_Verif { get; set; }

        [Column("ValTotIt")]
        public decimal? ValTotIt { get; set; }

        [Column("ValTotItIva")]
        public decimal? ValTotItIva { get; set; }

        [Column("TipoTran")]
        public short? TipoTran { get; set; }

        [Column("VALOBADA_1")]
        [StringLength(256)]
        public string? VALOBADA_1 { get; set; }

        [Column("VALOBADA_2")]
        [StringLength(256)]
        public string? VALOBADA_2 { get; set; }

        [Column("OCOMPRA_RTOABIERTO")]
        [StringLength(64)]
        public string? OCOMPRA_RTOABIERTO { get; set; }

        [Column("NroEnvio")]
        public int? NroEnvio { get; set; }

        [Column("NUMPVTA")]
        public short? NUMPVTA { get; set; }

        [Column("DOPRINUM")]
        public int? DOPRINUM { get; set; }

        [Column("IdCalidad")]
        public int? IdCalidad { get; set; }

        [Column("ID_OPERACION")]
        public short? ID_OPERACION { get; set; }

        [Column("PESOKG")]
        public decimal? PESOKG { get; set; }

        [Column("METROS")]
        public decimal? METROS { get; set; }

        [Column("CierreId")]
        public int? CierreId { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================
        
        /// <summary>
        /// Fecha de creación del registro
        /// Se establece automáticamente cuando se crea el movimiento de reparación
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
    }
}
/* CREA LA TABLA Y LOS INDICES 
 * -- =============================================
-- Script para crear tabla MOVIREPA
-- Estructura idéntica a MOVISTO
-- =============================================

-- Crear la tabla MOVIREPA
CREATE TABLE [dbo].[movirepa](
    [Id] [int] IDENTITY(1,1) NOT NULL,
    [FeReMovi] [datetime] NULL,
    [CodiEmpr] [int] NULL,
    [NumUsuar] [int] NULL,
    [IdenMovi] [nvarchar](32) NULL,
    [DeLaRemi] [nvarchar](max) NULL,
    [Descri_Lar] [nvarchar](max) NULL,
    [FechMov] [datetime] NULL,
    [Cod_Inte] [int] NULL,
    [Cod_Exte] [nvarchar](24) NULL,
    [Descrip] [nvarchar](64) NULL,
    [IdenLote] [nvarchar](16) NULL,
    [LotConsum] [nvarchar](max) NULL,
    [PNL_Des] [nvarchar](32) NULL,
    [PLC_Has] [nvarchar](32) NULL,
    [CodiMovi] [nvarchar](2) NULL,
    [DoPriCor] [nvarchar](12) NULL,
    [DoPriLar] [nvarchar](24) NULL,
    [DoSecCor] [nvarchar](12) NULL,
    [DoSecLar] [nvarchar](24) NULL,
    [DocuOrig] [nvarchar](24) NULL,
    [NumRem_Cli] [nvarchar](16) NULL,
    [NuorItem] [int] NULL,
    [NuPedCli] [int] NULL,
    [NumeOrdPed] [smallint] NULL,
    [ReferCor] [nvarchar](64) NULL,
    [CostoRep] [decimal](18, 2) NULL,
    [ValoUnit] [decimal](18, 2) NULL,
    [MoneVal] [smallint] NULL,
    [TipoCam] [decimal](18, 2) NULL,
    [DepoMovi] [int] NULL,
    [Nume_Clie] [int] NULL,
    [Sucu_Clie] [smallint] NULL,
    [Nume_Prov] [int] NULL,
    [CantIngre] [decimal](18, 2) NULL,
    [CantSalid] [decimal](18, 2) NULL,
    [CantSaldo] [decimal](18, 2) NULL,
    [Verificado] [smallint] NULL,
    [Causa] [smallint] NULL,
    [User_Tran] [nvarchar](32) NULL,
    [CodCarpImp] [nvarchar](12) NULL,
    [ValoUnitIva] [decimal](18, 2) NULL,
    [VerificadoPor] [smallint] NULL,
    [FHor_Verif] [datetime] NULL,
    [ValTotIt] [decimal](18, 2) NULL,
    [ValTotItIva] [decimal](18, 2) NULL,
    [TipoTran] [smallint] NULL,
    [VALOBADA_1] [nvarchar](256) NULL,
    [VALOBADA_2] [nvarchar](256) NULL,
    [OCOMPRA_RTOABIERTO] [nvarchar](64) NULL,
    [NroEnvio] [int] NULL,
    [NUMPVTA] [smallint] NULL,
    [DOPRINUM] [int] NULL,
    [IdCalidad] [int] NULL,
    [ID_OPERACION] [smallint] NULL,
    [PESOKG] [decimal](18, 2) NULL,
    [METROS] [decimal](18, 2) NULL,
    [CierreId] [int] NULL,
    [AddRecord] [datetime] NULL,
    [LastUpdate] [datetime] NULL,
    [CodiEmprNet] [int] NULL,
 CONSTRAINT [PK_movirepa] PRIMARY KEY CLUSTERED 
(
    [Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

-- =============================================
-- CREAR ÍNDICES NO AGRUPADOS
-- =============================================

-- Índice para CantIngre
CREATE NONCLUSTERED INDEX [ICantIngre] ON [dbo].[movirepa]
(
    [CantIngre] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para CantSalid
CREATE NONCLUSTERED INDEX [ICantSalid] ON [dbo].[movirepa]
(
    [CantSalid] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para Cod_Inte
CREATE NONCLUSTERED INDEX [ICod_Inte] ON [dbo].[movirepa]
(
    [Cod_Inte] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para CodiEmpr
CREATE NONCLUSTERED INDEX [ICodiEmpr] ON [dbo].[movirepa]
(
    [CodiEmpr] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para CodiMovi
SET ANSI_PADDING ON
GO
CREATE NONCLUSTERED INDEX [ICodiMovi] ON [dbo].[movirepa]
(
    [CodiMovi] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para DepoMovi
CREATE NONCLUSTERED INDEX [IDepoMovi] ON [dbo].[movirepa]
(
    [DepoMovi] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para DoPriCor
SET ANSI_PADDING ON
GO
CREATE NONCLUSTERED INDEX [IDoPriCor] ON [dbo].[movirepa]
(
    [DoPriCor] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para DoPriLar
SET ANSI_PADDING ON
GO
CREATE NONCLUSTERED INDEX [IDoPriLar] ON [dbo].[movirepa]
(
    [DoPriLar] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para DOPRINUM
CREATE NONCLUSTERED INDEX [IDOPRINUM] ON [dbo].[movirepa]
(
    [DOPRINUM] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para DoSecCor
SET ANSI_PADDING ON
GO
CREATE NONCLUSTERED INDEX [IDoSecCor] ON [dbo].[movirepa]
(
    [DoSecCor] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para DoSecLar
SET ANSI_PADDING ON
GO
CREATE NONCLUSTERED INDEX [IDoSecLar] ON [dbo].[movirepa]
(
    [DoSecLar] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para FechMov
CREATE NONCLUSTERED INDEX [IFechMov] ON [dbo].[movirepa]
(
    [FechMov] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para IdenLote
SET ANSI_PADDING ON
GO
CREATE NONCLUSTERED INDEX [IIdenLote] ON [dbo].[movirepa]
(
    [IdenLote] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para Nume_Clie
CREATE NONCLUSTERED INDEX [INume_Clie] ON [dbo].[movirepa]
(
    [Nume_Clie] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para NUMPVTA
CREATE NONCLUSTERED INDEX [INUMPVTA] ON [dbo].[movirepa]
(
    [NUMPVTA] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para NumUsuar
CREATE NONCLUSTERED INDEX [INumUsuar] ON [dbo].[movirepa]
(
    [NumUsuar] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice para NuPedCli
CREATE NONCLUSTERED INDEX [INuPedCli] ON [dbo].[movirepa]
(
    [NuPedCli] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- =============================================
-- ÍNDICES COMPUESTOS RECOMENDADOS
-- =============================================

-- Índice compuesto para búsquedas por fecha y código de movimiento
CREATE NONCLUSTERED INDEX [IX_FechMov_CodiMovi] ON [dbo].[movirepa]
(
    [FechMov] ASC,
    [CodiMovi] ASC
)
INCLUDE ([Cod_Inte], [CantIngre], [CantSalid], [DepoMovi])
WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice compuesto para búsquedas por cliente y fecha
CREATE NONCLUSTERED INDEX [IX_Nume_Clie_FechMov] ON [dbo].[movirepa]
(
    [Nume_Clie] ASC,
    [FechMov] ASC
)
INCLUDE ([CodiMovi], [DoPriCor], [CantSalid])
WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice compuesto para búsquedas por producto y depósito
CREATE NONCLUSTERED INDEX [IX_Cod_Inte_DepoMovi] ON [dbo].[movirepa]
(
    [Cod_Inte] ASC,
    [DepoMovi] ASC
)
INCLUDE ([FechMov], [CantIngre], [CantSalid], [IdenLote])
WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- Índice compuesto para búsquedas por remito (NUMPVTA + DOPRINUM)
CREATE NONCLUSTERED INDEX [IX_NUMPVTA_DOPRINUM] ON [dbo].[movirepa]
(
    [NUMPVTA] ASC,
    [DOPRINUM] ASC
)
INCLUDE ([FechMov], [Nume_Clie], [CodiMovi])
WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, FILLFACTOR = 90) ON [PRIMARY]
GO

-- =============================================
-- VALORES POR DEFECTO Y RESTRICCIONES
-- =============================================

-- Agregar valor por defecto para AddRecord
ALTER TABLE [dbo].[movirepa] ADD CONSTRAINT [DF_movirepa_AddRecord] DEFAULT (getdate()) FOR [AddRecord]
GO

-- Agregar valor por defecto para LastUpdate
ALTER TABLE [dbo].[movirepa] ADD CONSTRAINT [DF_movirepa_LastUpdate] DEFAULT (getdate()) FOR [LastUpdate]
GO

-- =============================================
-- SCRIPT COMPLETADO
-- =============================================

PRINT 'Tabla MOVIREPA creada exitosamente con todos los índices'
PRINT 'Total de índices creados: 22 (18 simples + 4 compuestos)'
GO */
