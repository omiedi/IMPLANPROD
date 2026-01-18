using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad para Recepción de Bombas para Evaluar Reparación
    /// Tabla: REPARA
    /// </summary>
    [Table("REPARA")]
    public class Repara : IAuditableEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; } // ID autoincremental

        public int NumRepa { get; set; } // Número de reparación

        [StringLength(128)]
        public string? Referen { get; set; } // Referencia

        public DateTime? FechRecep { get; set; } // Fecha de recepción

        public DateTime? FechEnSol { get; set; } // Fecha entrega solicitada

        public DateTime? FechaCierreRepa { get; set; } // Fecha cierre reparación

        public int? Nume_Cli { get; set; } // Número de cliente (FK a DEUDOR12)

        [StringLength(128)]
        public string? Raso_Cli { get; set; } // Razón social cliente

        [StringLength(64)]
        public string? Domi_Cli { get; set; } // Domicilio cliente

        [StringLength(64)]
        public string? Loca_Cli { get; set; } // Localidad cliente

        [StringLength(64)]
        public string? Pais_Cli { get; set; } // País cliente

        [StringLength(24)]
        public string? Tele_Cli { get; set; } // Teléfono cliente

        [StringLength(64)]
        public string? Mail_Cli { get; set; } // Email cliente

        public int? Cod_Inte { get; set; } // Código interno producto (FK a MASTER)

        [StringLength(128)]
        public string? Descripcion { get; set; } // Descripción del equipo

        [StringLength(32)]
        public string? NumeroSerie { get; set; } // Número de serie

        [StringLength(2048)]
        public string? Accesorios { get; set; } // Accesorios

        [StringLength(1024)]
        public string? Defec_Infor { get; set; } // Defectos informados

        [Column(TypeName = "nvarchar(max)")]
        public string? Observa { get; set; } // Observaciones

        [StringLength(1024)]
        public string? Trab_Efect { get; set; } // Trabajos efectuados

        [Column(TypeName = "decimal(18,4)")]
        public decimal? PresuEsti { get; set; } // Presupuesto estimado

        public short? Status { get; set; } // Estado

        [StringLength(24)]
        public string? Usuario_Repu { get; set; } // Usuario reparación

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Estima_Horas { get; set; } // Estimación de horas

        [StringLength(24)]
        public string? Or_Compra { get; set; } // Orden de compra

        public DateTime? FOr_Compra { get; set; } // Fecha orden de compra

        [StringLength(24)]
        public string? NumRemito { get; set; } // Número de remito

        public DateTime? FechEnt { get; set; } // Fecha de entrega

        [StringLength(48)]
        public string? Nom_Tra { get; set; } // Nombre transporte

        [StringLength(24)]
        public string? Cuit_Cli { get; set; } // CUIT cliente

        [StringLength(48)]
        public string? Piva_Cli { get; set; } // Posición IVA cliente

        [StringLength(48)]
        public string? Conductor { get; set; } // Conductor

        [StringLength(16)]
        public string? Patente { get; set; } // Patente

        [StringLength(128)]
        public string? Lentre_Cli { get; set; } // Lugar de entrega cliente

        public short? Status_Gar { get; set; } // Estado garantía

        public short? CausaFalla { get; set; } // Causa de falla

        [StringLength(128)]
        public string? DeCauFalla { get; set; } // Descripción causa falla

        public short? TipoFalla { get; set; } // Tipo de falla

        [StringLength(128)]
        public string? DeTipFalla { get; set; } // Descripción tipo falla

        [StringLength(64)]
        public string? Referen2 { get; set; } // Referencia 2

        [Column(TypeName = "decimal(18,4)")]
        public decimal? HoraMObra { get; set; } // Horas mano de obra

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Imp_Mobra { get; set; } // Importe mano de obra

        [Column(TypeName = "decimal(18,4)")]
        public decimal? ValMObra { get; set; } // Valor mano de obra

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Imp_Repu { get; set; } // Importe repuestos

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Imp_Tot { get; set; } // Importe total

        public DateTime? FechPres { get; set; } // Fecha presupuesto

        public DateTime? FechApro { get; set; } // Fecha aprobación

        public int? NroPed { get; set; } // Número de pedido

        [Column(TypeName = "nvarchar(max)")]
        public string? AnotaRep { get; set; } // Anotaciones reparación

        [StringLength(16)]
        public string? CodRepa { get; set; } // Código reparación (FK a INDIREP)

        public int? NuOrdRep { get; set; } // Número orden reparación

        [StringLength(128)]
        public string? RefOrdRep { get; set; } // Referencia orden reparación

        public DateTime? FechOrdRep { get; set; } // Fecha orden reparación

        [Column(TypeName = "nvarchar(max)")]
        public string? Inspeccion { get; set; } // Inspección

        [Column(TypeName = "nvarchar(max)")]
        public string? Reparacion { get; set; } // Reparación

        [StringLength(12)]
        public string? IdSubOr { get; set; } // ID sub orden

        [Column(TypeName = "nvarchar(max)")]
        public string? Sitio { get; set; } // Sitio

        [StringLength(24)]
        public string? Factura { get; set; } // Factura

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Importe { get; set; } // Importe

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Tot_Repuestos { get; set; } // Total repuestos

        [Column(TypeName = "decimal(18,4)")]
        public decimal? Tot_Mano_Obra { get; set; } // Total mano de obra

        [StringLength(32)]
        public string? UsRealizaRepa { get; set; } // Usuario que realiza reparación

        [Column(TypeName = "decimal(18,4)")]
        public decimal? DESCTO1 { get; set; } // Descuento 1 del presupuesto

        [Column(TypeName = "decimal(18,4)")]
        public decimal? DESCTO2 { get; set; } // Descuento 2 por medio de cobranza

        public short? MED_COBRANZA { get; set; } // Medio de cobranza (FK a MEDITEA)

        [StringLength(16)]
        public string? NUIDASIEN { get; set; } // Número ID asiento

        public short? TEC_EXT { get; set; } // Técnico externo

        [StringLength(24)]
        public string? TALON_TEC { get; set; } // Talón técnico

        [Column(TypeName = "decimal(18,4)")]
        public decimal? MOBRA_EN_GTIA { get; set; } // Mano de obra en garantía

        [StringLength(128)]
        public string? ObsLibre { get; set; } // Observaciones libres

        public short? MoneEmis { get; set; } // Moneda emisión

        [Column(TypeName = "decimal(10,4)")]
        public decimal? TipoCam { get; set; } // Tipo de cambio

        public short? STAT_APROBACION { get; set; } // Estado aprobación

        public short? CondPag { get; set; } // Condición de pago

        // Campos de auditoría (IAuditableEntity)
        public DateTime? AddRecord { get; set; }
        public DateTime? LastUpdate { get; set; }
        public int? CodiEmprNet { get; set; }
        public int? NumUsuar { get; set; }
        public int? CodiEmpr { get; set; }
    }
}

/*
-- Script SQL para crear la tabla REPARA
-- Recepción de Bombas para Evaluar Reparación

IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[REPARA]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[REPARA](
        [Id] INT IDENTITY(1,1) NOT NULL,
        [NumRepa] INT NOT NULL,
        [Referen] NVARCHAR(128) NULL,
        [FechRecep] DATETIME NULL,
        [FechEnSol] DATETIME NULL,
        [FechaCierreRepa] DATETIME NULL,
        [Nume_Cli] INT NULL,
        [Raso_Cli] NVARCHAR(128) NULL,
        [Domi_Cli] NVARCHAR(64) NULL,
        [Loca_Cli] NVARCHAR(64) NULL,
        [Pais_Cli] NVARCHAR(64) NULL,
        [Tele_Cli] NVARCHAR(24) NULL,
        [Mail_Cli] NVARCHAR(64) NULL,
        [Cod_Inte] INT NULL,
        [Descripcion] NVARCHAR(128) NULL,
        [NumeroSerie] NVARCHAR(32) NULL,
        [Accesorios] NVARCHAR(2048) NULL,
        [Defec_Infor] NVARCHAR(1024) NULL,
        [Observa] NVARCHAR(MAX) NULL,
        [Trab_Efect] NVARCHAR(1024) NULL,
        [PresuEsti] DECIMAL(18,4) NULL,
        [Status] SMALLINT NULL,
        [Usuario_Repu] NVARCHAR(24) NULL,
        [Estima_Horas] DECIMAL(18,4) NULL,
        [Or_Compra] NVARCHAR(24) NULL,
        [FOr_Compra] DATETIME NULL,
        [NumRemito] NVARCHAR(24) NULL,
        [FechEnt] DATETIME NULL,
        [Nom_Tra] NVARCHAR(48) NULL,
        [Cuit_Cli] NVARCHAR(24) NULL,
        [Piva_Cli] NVARCHAR(48) NULL,
        [Conductor] NVARCHAR(48) NULL,
        [Patente] NVARCHAR(16) NULL,
        [Lentre_Cli] NVARCHAR(128) NULL,
        [Status_Gar] SMALLINT NULL,
        [CausaFalla] SMALLINT NULL,
        [DeCauFalla] NVARCHAR(128) NULL,
        [TipoFalla] SMALLINT NULL,
        [DeTipFalla] NVARCHAR(128) NULL,
        [Referen2] NVARCHAR(64) NULL,
        [HoraMObra] DECIMAL(18,4) NULL,
        [Imp_Mobra] DECIMAL(18,4) NULL,
        [ValMObra] DECIMAL(18,4) NULL,
        [Imp_Repu] DECIMAL(18,4) NULL,
        [Imp_Tot] DECIMAL(18,4) NULL,
        [FechPres] DATETIME NULL,
        [FechApro] DATETIME NULL,
        [NroPed] INT NULL,
        [AnotaRep] NVARCHAR(MAX) NULL,
        [CodRepa] NVARCHAR(16) NULL,
        [NuOrdRep] INT NULL,
        [RefOrdRep] NVARCHAR(128) NULL,
        [FechOrdRep] DATETIME NULL,
        [Inspeccion] NVARCHAR(MAX) NULL,
        [Reparacion] NVARCHAR(MAX) NULL,
        [IdSubOr] NVARCHAR(12) NULL,
        [Sitio] NVARCHAR(MAX) NULL,
        [Factura] NVARCHAR(24) NULL,
        [Importe] DECIMAL(18,4) NULL,
        [Tot_Repuestos] DECIMAL(18,4) NULL,
        [Tot_Mano_Obra] DECIMAL(18,4) NULL,
        [UsRealizaRepa] NVARCHAR(32) NULL,
        [DESCTO1] DECIMAL(18,4) NULL,
        [DESCTO2] DECIMAL(18,4) NULL,
        [MED_COBRANZA] SMALLINT NULL,
        [NUIDASIEN] NVARCHAR(16) NULL,
        [TEC_EXT] SMALLINT NULL,
        [TALON_TEC] NVARCHAR(24) NULL,
        [MOBRA_EN_GTIA] DECIMAL(18,4) NULL,
        [ObsLibre] NVARCHAR(128) NULL,
        [MoneEmis] SMALLINT NULL,
        [TipoCam] DECIMAL(10,4) NULL,
        [STAT_APROBACION] SMALLINT NULL,
        [CondPag] SMALLINT NULL,
        -- Campos de auditoría
        [AddRecord] DATETIME NULL,
        [LastUpdate] DATETIME NULL,
        [CodiEmprNet] INT NULL,
        [NumUsuar] INT NULL,
        [CodiEmpr] INT NULL,
        CONSTRAINT [PK_REPARA] PRIMARY KEY CLUSTERED ([Id] ASC)
    )
END
GO

-- Índices para mejorar el rendimiento
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_REPARA_NumRepa' AND object_id = OBJECT_ID('REPARA'))
    CREATE NONCLUSTERED INDEX [IX_REPARA_NumRepa] ON [dbo].[REPARA] ([NumRepa])
GO
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_REPARA_Nume_Cli' AND object_id = OBJECT_ID('REPARA'))
    CREATE NONCLUSTERED INDEX [IX_REPARA_Nume_Cli] ON [dbo].[REPARA] ([Nume_Cli])
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_REPARA_CodRepa' AND object_id = OBJECT_ID('REPARA'))
    CREATE NONCLUSTERED INDEX [IX_REPARA_CodRepa] ON [dbo].[REPARA] ([CodRepa])
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_REPARA_FechRecep' AND object_id = OBJECT_ID('REPARA'))
    CREATE NONCLUSTERED INDEX [IX_REPARA_FechRecep] ON [dbo].[REPARA] ([FechRecep])
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_REPARA_Status' AND object_id = OBJECT_ID('REPARA'))
    CREATE NONCLUSTERED INDEX [IX_REPARA_Status] ON [dbo].[REPARA] ([Status])
GO

IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_REPARA_NumeroSerie' AND object_id = OBJECT_ID('REPARA'))
    CREATE NONCLUSTERED INDEX [IX_REPARA_NumeroSerie] ON [dbo].[REPARA] ([NumeroSerie])
GO

-- Comentarios de la tabla
EXEC sys.sp_addextendedproperty 
    @name=N'MS_Description', 
    @value=N'Tabla para Recepción de Bombas para Evaluar Reparación - Migrado desde VB6 RECEBO09.frm' , 
    @level0type=N'SCHEMA',
    @level0name=N'dbo', 
    @level1type=N'TABLE',
    @level1name=N'REPARA'
GO
*/
