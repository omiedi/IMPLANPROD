using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que almacena las novedades de cambios de precios de productos
    /// Registra el historial de modificaciones de precios con información del proveedor y moneda
    /// </summary>
    [Table("NOVEDADES")]
    public class Novedades : ITimestampEntity, IAuditableEntity
    {
        // ===================================================================
        // CLAVE PRIMARIA (AGREGADA)
        // ===================================================================

        /// <summary>
        /// ID único de la novedad (campo agregado para identificación única)
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("Id")]
        public int Id { get; set; }

        // ===================================================================
        // CAMPOS ORIGINALES DE LA TABLA
        // ===================================================================

        /// <summary>
        /// Código interno del producto
        /// Relacionado con MASTER.CodInte
        /// Conversión: smallint (SQL) → int (C#)
        /// </summary>
        [Column("Codint")]
        public int? Codint { get; set; }

        /// <summary>
        /// Código de empresa
        /// Por defecto siempre 0 según reglas globales del sistema
        /// </summary>
        [Column("CodiEmpr")]
        public int CodiEmpr { get; set; } = 0;

        /// <summary>
        /// Fecha de modificación del precio
        /// Registra cuándo se realizó el cambio de precio
        /// </summary>
        [Column("Fech_Mod")]
        public DateTime? FechMod { get; set; }

        /// <summary>
        /// Precio anterior (SIEMPRE EN PESOS)
        /// Almacena el precio previo antes de la modificación convertido a pesos
        /// Si es la primera vez, se guarda igual a PrecNew
        /// Conversión: float (SQL) → decimal (C#)
        /// Ejemplo: Si era USD 100 con cotiz 1050 = $105,000 pesos
        /// </summary>
        [Column("Prec_Ant", TypeName = "decimal(18,4)")]
        public decimal? PrecAnt { get; set; }

        /// <summary>
        /// Precio nuevo (SIEMPRE EN PESOS)
        /// Almacena el nuevo precio después de la modificación convertido a pesos
        /// Se calcula multiplicando el precio en moneda original por la cotización
        /// Conversión: float (SQL) → decimal (C#)
        /// Ejemplo: USD 110 * cotiz 1050 = $115,500 pesos
        /// </summary>
        [Column("Prec_New", TypeName = "decimal(18,4)")]
        public decimal? PrecNew { get; set; }

        /// <summary>
        /// Número de proveedor
        /// Relacionado con Proved12.NumProv
        /// </summary>
        [Column("NumProv")]
        public int? NumProv { get; set; }

        /// <summary>
        /// Observaciones sobre el cambio de precio
        /// Permite registrar comentarios o justificaciones del cambio
        /// </summary>
        [Column("Observacion")]
        [MaxLength(512)]
        public string? Observacion { get; set; }

        /// <summary>
        /// Código de moneda
        /// Relacionado con TAMONED.CodMoneda
        /// </summary>
        [Column("Moneda")]
        public short? Moneda { get; set; }

        /// <summary>
        /// Lista de reposición
        /// Identificador de la lista de precios aplicable
        /// </summary>
        [Column("Lista_Repos")]
        [MaxLength(12)]
        public string? ListaRepos { get; set; }

        /// <summary>
        /// Cotización de la moneda al momento del cambio
        /// Se usa para convertir el precio en moneda original a pesos
        /// IMPORTANTE: Permite detectar cambios de precio por variación de cotización
        /// Ejemplo: Si el precio en USD no cambió pero la cotización sí,
        ///          el precio en pesos será diferente y debe registrarse
        /// Conversión: float (SQL) → decimal (C#)
        /// Valores típicos: 1 (pesos), 1050 (USD), 1150 (EUR)
        /// </summary>
        [Column("COTIZA_DOLAR", TypeName = "decimal(18,4)")]
        public decimal? CotizaDolar { get; set; }

        /// <summary>
        /// Número de orden de compra asociada
        /// Relacionado con OCOMENCA.Nume_Ocom
        /// Permite vincular el cambio de precio con una orden de compra específica
        /// </summary>
        [Column("Nume_Ocom")]
        public int? NumeOcom { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// Se asigna automáticamente al crear la novedad
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// Se actualiza automáticamente al modificar la novedad
        /// </summary>
        [Column("LastUpdate")]
        public DateTime? LastUpdate { get; set; }

        /// <summary>
        /// Código de empresa - Implementación de IAuditableEntity.CodiEmprNet
        /// Por defecto siempre 0 según reglas globales
        /// </summary>
        [Column("CodiEmprNet")]
        public int? CodiEmprNet { get; set; }

        /// <summary>
        /// Número de usuario que realizó la operación
        /// Relacionado con usuarios.idflexoft para validación de derechos
        /// Implementación de IAuditableEntity.NumUsuar
        /// </summary>
        [Column("NumUsuar")]
        public int? NumUsuar { get; set; }
    }
}

/*
===============================================================================
SCRIPT SQL PARA CREAR LA TABLA EN SQL SERVER MANAGEMENT STUDIO
===============================================================================

-- =============================================
-- Script: Crear tabla NOVEDADES
-- Descripción: Tabla para almacenar historial de cambios de precios de productos
-- Fecha: 2026-04-14
-- =============================================

USE [IMPLANPROD]
GO

-- Verificar si la tabla ya existe
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[NOVEDADES]') AND type in (N'U'))
BEGIN
    CREATE TABLE [dbo].[NOVEDADES](
        -- ===================================================================
        -- CLAVE PRIMARIA (AGREGADA)
        -- ===================================================================
        [Id] [int] IDENTITY(1,1) NOT NULL,
        
        -- ===================================================================
        -- CAMPOS ORIGINALES
        -- ===================================================================
        [Codint] [int] NULL,
        [CodiEmpr] [int] NULL DEFAULT (0),
        [Fech_Mod] [smalldatetime] NULL,
        [Prec_Ant] [decimal](18,4) NULL,
        [Prec_New] [decimal](18,4) NULL,
        [NumProv] [int] NULL,
        [Observacion] [nvarchar](512) NULL,
        [Moneda] [smallint] NULL,
        [Lista_Repos] [nvarchar](12) NULL,
        [COTIZA_DOLAR] [decimal](18,4) NULL,
        [Nume_Ocom] [int] NULL,
        
        -- ===================================================================
        -- CAMPOS DE AUDITORÍA (AGREGADOS)
        -- ===================================================================
        [AddRecord] [datetime2](7) NULL,
        [LastUpdate] [datetime2](7) NULL,
        [CodiEmprNet] [int] NULL DEFAULT (0),
        [NumUsuar] [int] NULL,
        
        -- ===================================================================
        -- CONSTRAINT DE CLAVE PRIMARIA
        -- ===================================================================
        CONSTRAINT [PK_NOVEDADES] PRIMARY KEY CLUSTERED ([Id] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
    ) ON [PRIMARY]
    
    PRINT 'Tabla NOVEDADES creada exitosamente'
END
ELSE
BEGIN
    PRINT 'La tabla NOVEDADES ya existe'
END
GO

-- ===================================================================
-- ÍNDICES RECOMENDADOS PARA MEJORAR RENDIMIENTO
-- ===================================================================

-- Índice para búsqueda por código interno de producto
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_NOVEDADES_Codint' AND object_id = OBJECT_ID('NOVEDADES'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_NOVEDADES_Codint] ON [dbo].[NOVEDADES]
    (
        [Codint] ASC
    )
    INCLUDE ([Fech_Mod], [Prec_Ant], [Prec_New], [NumProv])
    WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
    
    PRINT 'Índice IX_NOVEDADES_Codint creado'
END
GO

-- Índice para búsqueda por proveedor
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_NOVEDADES_NumProv' AND object_id = OBJECT_ID('NOVEDADES'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_NOVEDADES_NumProv] ON [dbo].[NOVEDADES]
    (
        [NumProv] ASC
    )
    INCLUDE ([Codint], [Fech_Mod], [Prec_New])
    WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
    
    PRINT 'Índice IX_NOVEDADES_NumProv creado'
END
GO

-- Índice para búsqueda por fecha de modificación
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = 'IX_NOVEDADES_FechMod' AND object_id = OBJECT_ID('NOVEDADES'))
BEGIN
    CREATE NONCLUSTERED INDEX [IX_NOVEDADES_FechMod] ON [dbo].[NOVEDADES]
    (
        [Fech_Mod] DESC
    )
    INCLUDE ([Codint], [NumProv], [Prec_Ant], [Prec_New])
    WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON)
    
    PRINT 'Índice IX_NOVEDADES_FechMod creado'
END
GO

-- ===================================================================
-- VERIFICACIÓN FINAL
-- ===================================================================
SELECT 
    'Tabla creada' AS Estado,
    COUNT(*) AS TotalRegistros
FROM [dbo].[NOVEDADES]
GO

PRINT '============================================='
PRINT 'Script completado exitosamente'
PRINT 'Tabla: NOVEDADES'
PRINT 'Descripción: Historial de cambios de precios'
PRINT 'Índices: 3 (Codint, NumProv, FechMod)'
PRINT '============================================='
GO

-- ===================================================================
-- MODIFICACIONES PARA BASES DE DATOS EXISTENTES
-- ===================================================================

-- Modificar columna Codint de smallint a int
-- IMPORTANTE: Ejecutar este script si la tabla ya existe y necesita actualización
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[NOVEDADES]') AND type in (N'U'))
BEGIN
    -- Verificar si la columna es smallint
    IF EXISTS (SELECT * FROM sys.columns 
               WHERE object_id = OBJECT_ID(N'[dbo].[NOVEDADES]') 
               AND name = 'Codint' 
               AND system_type_id = 52) -- 52 = smallint
    BEGIN
        -- Modificar el tipo de dato de smallint a int
        ALTER TABLE [dbo].[NOVEDADES]
        ALTER COLUMN [Codint] [int] NULL
        
        PRINT '✅ Columna Codint modificada de smallint a int exitosamente'
    END
    ELSE
    BEGIN
        PRINT 'ℹ️ La columna Codint ya es de tipo int o no existe'
    END
END
GO

===============================================================================
NOTAS DE MAPEO Y CONVERSIONES:
===============================================================================

CONVERSIONES APLICADAS:
- float (SQL) → decimal (C#) para Prec_Ant, Prec_New, COTIZA_DOLAR
- smallint (SQL) → int (C#) para Codint (modificado para soportar códigos más grandes)
- smallint (SQL) → short? (C#) para Moneda
- CodiEmpr: smallint (SQL) → int (C#) con valor por defecto 0
- nvarchar → string con [MaxLength]
- smalldatetime → DateTime?

CAMPOS AGREGADOS:
- Id: int IDENTITY(1,1) - Clave primaria única
- AddRecord: datetime2 - Fecha de creación del registro
- LastUpdate: datetime2 - Fecha de última actualización
- CodiEmprNet: int - Código de empresa para auditoría (siempre 0)
- NumUsuar: int - Usuario que realizó la operación

FLUJO DE DATOS:
1. FRONTEND → BACKEND:
   - Usuario registra cambio de precio en la UI
   - Se envía DTO con datos de la novedad
   - Service asigna automáticamente AddRecord, NumUsuar, CodiEmprNet

2. BACKEND → FRONTEND:
   - Se consulta historial de novedades por producto o proveedor
   - Se retorna lista de cambios con fechas y precios
   - UI muestra evolución de precios en el tiempo

3. RELACIONES:
   - Codint → MASTER.CodInte (producto)
   - NumProv → Proved12.NumProv (proveedor)
   - Moneda → TAMONED.CodMoneda (tipo de moneda)
   - NumeOcom → OCOMENCA.Nume_Ocom (orden de compra)

REGLAS DE NEGOCIO:
- CodiEmpr siempre = 0 (regla global del sistema)
- NumUsuar se obtiene de usuarios.idflexoft del usuario activo
- Prec_Ant y Prec_New permiten calcular variación porcentual
- CotizaDolar permite conversiones de moneda al momento del cambio
- FechMod registra la fecha exacta del cambio para auditoría

*/
