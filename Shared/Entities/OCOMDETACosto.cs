using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que almacena los costos de los items de órdenes de compra
    /// Permite gestionar precios/costos de forma separada con control de acceso
    /// </summary>
    [Table("OCOMDETACosto")]
    public class OCOMDETACosto : ITimestampEntity, IAuditableEntity
    {
        // ===================================================================
        // CLAVE PRIMARIA
        // ===================================================================
        
        /// <summary>
        /// ID único del registro de costo
        /// </summary>
        [Key]
        [Column("OCOMDETA_ID")]
        public int OCOMDETA_ID { get; set; }

        // ===================================================================
        // CAMPOS DE IDENTIFICACIÓN DE LA ORDEN
        // ===================================================================

        /// <summary>
        /// Número de orden de compra
        /// Relacionado con OCOMENCA.NumeOcom y OCOMDETA.NumeOcom
        /// </summary>
        [Column("Nume_Ocom")]
        public int? NumeOcom { get; set; }

        /// <summary>
        /// Código de empresa (siempre 0 según regla global)
        /// </summary>
        [Column("CodiEmpr")]
        public int CodiEmpr { get; set; } = 0;

        /// <summary>
        /// Número de item dentro de la orden
        /// Relacionado con OCOMDETA.NuOrdItem
        /// </summary>
        [Column("NuOrd_Item")]
        public short? NuOrdItem { get; set; }

        // ===================================================================
        // INFORMACIÓN DEL ENCABEZADO (DUPLICADA PARA CONSULTAS)
        // ===================================================================

        /// <summary>
        /// Fecha de emisión de la orden
        /// </summary>
        [Column("Femi_Ocom")]
        public DateTime? FemiOcom { get; set; }

        /// <summary>
        /// Número de proveedor
        /// </summary>
        [Column("Npro_Ocom")]
        public short? NproOcom { get; set; }

        /// <summary>
        /// Razón social del proveedor
        /// </summary>
        [Column("RaSo_Prov")]
        [MaxLength(64)]
        public string? RaSoProv { get; set; }

        // ===================================================================
        // INFORMACIÓN DEL PRODUCTO
        // ===================================================================

        /// <summary>
        /// Código interno del producto
        /// Relacionado con OCOMDETA.CodInte
        /// </summary>
        [Column("Cod_Inte")]
        public short? CodInte { get; set; }

        /// <summary>
        /// Código externo del producto
        /// Relacionado con OCOMDETA.CodExte
        /// </summary>
        [Column("Cod_Exte")]
        [MaxLength(24)]
        public string? CodExte { get; set; }

        /// <summary>
        /// Descripción del item
        /// Relacionado con OCOMDETA.Descritem
        /// </summary>
        [Column("Descritem")]
        [MaxLength(64)]
        public string? Descritem { get; set; }

        // ===================================================================
        // UNIDADES DE MEDIDA Y CANTIDADES
        // ===================================================================

        /// <summary>
        /// Unidad de medida de almacén
        /// </summary>
        [Column("Uni_Mas")]
        [MaxLength(4)]
        public string? UniMas { get; set; }

        /// <summary>
        /// Unidades por bulto/caja
        /// </summary>
        [Column("Unids_Ocom", TypeName = "decimal(18,4)")]
        public decimal? UnidsOcom { get; set; }

        /// <summary>
        /// Cantidad solicitada en la orden
        /// Relacionado con OCOMDETA.CantOcom
        /// </summary>
        [Column("Cant_Ocom", TypeName = "decimal(18,4)")]
        public decimal? CantOcom { get; set; }

        /// <summary>
        /// Unidad de medida de compra
        /// Relacionado con OCOMDETA.UniCom
        /// </summary>
        [Column("Uni_Com")]
        [MaxLength(4)]
        public string? UniCom { get; set; }

        // ===================================================================
        // COSTO/PRECIO (CAMPO PRINCIPAL EDITABLE)
        // ===================================================================

        /// <summary>
        /// Precio unitario neto (COSTO)
        /// Este es el campo principal que se edita en el modal de costos
        /// Relacionado con OCOMDETA.PreUnit pero almacenado de forma separada
        /// </summary>
        [Column("PrUn_Neto", TypeName = "decimal(18,4)")]
        public decimal? PrUnNeto { get; set; }

        /// <summary>
        /// Observaciones del costo del item
        /// Máximo 1024 caracteres
        /// </summary>
        [Column("Observaciones", TypeName = "varchar(1024)")]
        [MaxLength(1024)]
        public string? Observaciones { get; set; }

        // ===================================================================
        // FECHAS Y MONEDAS
        // ===================================================================

        /// <summary>
        /// Fecha de entrega solicitada
        /// Relacionado con OCOMDETA.FentreSol
        /// </summary>
        [Column("Fentre_Sol")]
        public DateTime? FentreSol { get; set; }

        /// <summary>
        /// Moneda del costo
        /// </summary>
        [Column("Mone_Cos")]
        public short? MoneCos { get; set; }

        /// <summary>
        /// Moneda de emisión de la orden
        /// Relacionado con TAMONED.CodigoMoneda para obtener descripción
        /// </summary>
        [Column("Mone_Emis")]
        public short? MoneEmis { get; set; }

        // ===================================================================
        // ESTADO
        // ===================================================================

        /// <summary>
        /// Estado del item en la orden
        /// Relacionado con OCOMDETA.StatOcom
        /// Valores: 0=Activo, 9=Anulado
        /// Solo se muestran items con StatOcom < 9
        /// </summary>
        [Column("Stat_Ocom")]
        public short? StatOcom { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// Se asigna automáticamente al crear
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// Se actualiza automáticamente al modificar
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        // ===================================================================
        // CAMPOS DE AUDITORÍA - Implementación de IAuditableEntity
        // ===================================================================

        /// <summary>
        /// Código de empresa para auditoría (siempre 0)
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario que creó/modificó el registro
        /// Relacionado con usuarios.idflexoft para validación de derechos
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
SCRIPT SQL PARA CREAR LA TABLA EN SQL SERVER MANAGEMENT STUDIO:
================================================================

-- =============================================
-- Script: Crear tabla OCOMDETACosto
-- Descripción: Tabla para almacenar costos de items de órdenes de compra con control de acceso separado
-- Fecha: 2026-04-13
-- =============================================

USE [IMPLANPROD]
GO

-- Verificar si la tabla ya existe
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[OCOMDETACosto]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[OCOMDETACosto](
        -- ===================================================================
        -- CLAVE PRIMARIA
        -- ===================================================================
        [OCOMDETA_ID] [int] IDENTITY(1,1) NOT NULL,
        
        -- ===================================================================
        -- CAMPOS DE IDENTIFICACIÓN DE LA ORDEN
        -- ===================================================================
        [Nume_Ocom] [int] NULL,
        [CodiEmpr] [int] NULL,
        [NuOrd_Item] [smallint] NULL,
        
        -- ===================================================================
        -- INFORMACIÓN DEL ENCABEZADO (DUPLICADA PARA CONSULTAS)
        -- ===================================================================
        [Femi_Ocom] [smalldatetime] NULL,
        [Npro_Ocom] [smallint] NULL,
        [RaSo_Prov] [nvarchar](64) NULL,
        
        -- ===================================================================
        -- INFORMACIÓN DEL PRODUCTO
        -- ===================================================================
        [Cod_Inte] [smallint] NULL,
        [Cod_Exte] [nvarchar](24) NULL,
        [Descritem] [nvarchar](64) NULL,
        
        -- ===================================================================
        -- UNIDADES DE MEDIDA Y CANTIDADES
        -- ===================================================================
        [Uni_Mas] [nvarchar](4) NULL,
        [Unids_Ocom] [decimal](18, 4) NULL,
        [Cant_Ocom] [decimal](18, 4) NULL,
        [Uni_Com] [nvarchar](4) NULL,
        
        -- ===================================================================
        -- COSTO/PRECIO (CAMPO PRINCIPAL EDITABLE)
        -- ===================================================================
        [PrUn_Neto] [decimal](18, 4) NULL,
        
        -- ===================================================================
        -- FECHAS Y MONEDAS
        -- ===================================================================
        [Fentre_Sol] [smalldatetime] NULL,
        [Mone_Cos] [smallint] NULL,
        [Mone_Emis] [smallint] NULL,
        
        -- ===================================================================
        -- ESTADO
        -- ===================================================================
        [Stat_Ocom] [smallint] NULL,
        
        -- ===================================================================
        -- CAMPOS DE TIMESTAMP AUTOMÁTICOS
        -- ===================================================================
        [AddRecord] [datetime2](7) NULL,
        [LastUpdate] [datetime2](7) NULL,
        
        -- ===================================================================
        -- CAMPOS DE AUDITORÍA
        -- ===================================================================
        [CodiEmprNet] [int] NULL,
        [NumUsuar] [int] NULL,
        
        -- ===================================================================
        -- CONSTRAINT DE CLAVE PRIMARIA
        -- ===================================================================
        CONSTRAINT [PK_OCOMDETACosto] PRIMARY KEY CLUSTERED ([OCOMDETA_ID] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
    ) ON [PRIMARY]
    
    PRINT 'Tabla OCOMDETACosto creada exitosamente'
END
ELSE
BEGIN
    PRINT 'La tabla OCOMDETACosto ya existe'
END
GO

-- ===================================================================
-- ÍNDICES RECOMENDADOS PARA MEJORAR RENDIMIENTO
-- ===================================================================

-- Índice para búsqueda por orden de compra
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_OCOMDETACosto_NumeOcom' AND object_id = OBJECT_ID('OCOMDETACosto'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_OCOMDETACosto_NumeOcom] ON [dbo].[OCOMDETACosto]
    (
        [Nume_Ocom] ASC
    )
    INCLUDE ([NuOrd_Item], [Cod_Inte], [PrUn_Neto])
    WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
    
    PRINT 'Índice IX_OCOMDETACosto_NumeOcom creado'
END
GO

-- Índice compuesto para búsqueda única (NumeOcom + NuOrdItem + CodInte)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_OCOMDETACosto_Unique' AND object_id = OBJECT_ID('OCOMDETACosto'))
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX [IX_OCOMDETACosto_Unique] ON [dbo].[OCOMDETACosto]
    (
        [Nume_Ocom] ASC,
        [NuOrd_Item] ASC,
        [Cod_Inte] ASC
    )
    WHERE ([Nume_Ocom] IS NOT NULL AND [NuOrd_Item] IS NOT NULL AND [Cod_Inte] IS NOT NULL)
    WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, IGNORE_DUP_KEY = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
    
    PRINT 'Índice IX_OCOMDETACosto_Unique creado'
END
GO

-- ===================================================================
-- VERIFICACIÓN FINAL
-- ===================================================================
SELECT 
    'Tabla creada' AS Estado,
    COUNT(*) AS TotalRegistros
FROM [dbo].[OCOMDETACosto]
GO

PRINT '============================================='
PRINT 'Script completado exitosamente'
PRINT 'Tabla: OCOMDETACosto'
PRINT 'Descripción: Costos de items de órdenes de compra'
PRINT 'Índices: 2 (NumeOcom, Unique)'
PRINT '============================================='
GO

NOTAS DE MAPEO:
===============
- smallint (SQL) → short (C#) para NuOrdItem, NproOcom, CodInte, MoneCos, MoneEmis, StatOcom
- decimal(18,4) (SQL) → decimal con [Precision(18,4)] (C#)
- nvarchar → string con [MaxLength]
- smalldatetime → DateTime?
- CodiEmpr siempre = 0 (regla global del sistema)

RELACIONES CLAVE:
=================
- NumeOcom + NuOrdItem + CodInte: Identifican únicamente un item de orden
- Se relaciona con OCOMDETA mediante estos 3 campos
- MoneEmis se relaciona con TAMONED.CodigoMoneda para obtener descripción
- NumUsuar se relaciona con usuarios.idflexoft para validación de derechos DER_ACCE

FLUJO DE DATOS:
===============
1. Usuario con derecho "VerCostosOC" (lectura/escritura) ve botón con ícono $
2. Al hacer clic, se abre modal mostrando items de OCOMDETA (StatOcom < 9)
3. Se hace JOIN con OCOMDETACosto para mostrar costos existentes
4. Usuario edita campo PrUnNeto (costo)
5. Al guardar, se hace INSERT o UPDATE en OCOMDETACosto
6. Se guardan: NumeOcom, NuOrdItem, CodInte, CodExte, Descritem, PrUnNeto
*/
