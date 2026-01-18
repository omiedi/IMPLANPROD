using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad BOLRECEP - Boletas de Recepción
    /// Representa las recepciones de materiales/productos de órdenes de compra
    /// </summary>
    [Table("BOLRECEP")]
    public class Bolrecep : IAuditableEntity
    {
        /// <summary>
        /// Identificador interno autoincremental (IDENTITY) y clave primaria
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }
        /// <summary>
        /// Número de Boleta de Recepción
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? NumeBoRe { get; set; }
        /// <summary>
        /// Código de Empresa
        /// </summary>
        public int? CodiEmpr { get; set; }
        /// <summary>
        /// Descripción
        /// </summary>
        [MaxLength(64)]
        public string? Descrip { get; set; }
        /// <summary>
        /// Identificador de Lote
        /// </summary>
        [MaxLength(12)]
        public string? IdenLote { get; set; }
        /// <summary>
        /// Fecha de Recepción
        /// </summary>
        public DateTime? FechRecep { get; set; }
        /// <summary>
        /// Código Interno del Producto
        /// </summary>
        [Column("Cod_Inte")]
        public int? CodInte { get; set; }
        /// <summary>
        /// Código Externo del Producto
        /// </summary>
        [Column("Cod_Exte")]
        [MaxLength(24)]
        public string? CodExte { get; set; }
        /// <summary>
        /// Descripción del Item
        /// </summary>
        [MaxLength(64)]
        public string? DescripItem { get; set; }
        /// <summary>
        /// Unidad de Medida
        /// </summary>
        [MaxLength(4)]
        public string? Unidad { get; set; }
        /// <summary>
        /// Cantidad Nominal
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CantNom { get; set; }
        /// <summary>
        /// Medida Largo
        /// </summary>
        [Column("M_Largo", TypeName = "decimal(18,4)")]
        public decimal? MLargo { get; set; }
        /// <summary>
        /// Medida Ancho
        /// </summary>
        [Column("M_Ancho", TypeName = "decimal(18,4)")]
        public decimal? MAncho { get; set; }
        /// <summary>
        /// Depósito de Entrada
        /// </summary>
        public short? DepoEnt { get; set; }
        /// <summary>
        /// Ubicación Física
        /// </summary>
        [MaxLength(32)]
        public string? UbicFis { get; set; }
        /// <summary>
        /// Número de Proveedor
        /// </summary>
        public int? NumProv { get; set; }
        /// <summary>
        /// Razón Social del Proveedor
        /// </summary>
        [Column("RaSo_Prov")]
        [MaxLength(64)]
        public string? RaSoProv { get; set; }
        /// <summary>
        /// Lote del Proveedor
        /// </summary>
        [Column("LoteCol_Prov")]
        [MaxLength(48)]
        public string? LoteColProv { get; set; }
        /// <summary>
        /// Certificado de Calidad del Proveedor
        /// </summary>
        [Column("CerCal_Prov")]
        [MaxLength(24)]
        public string? CerCalProv { get; set; }
        /// <summary>
        /// Número de Remito del Proveedor
        /// </summary>
        [Column("NumRem_Prov")]
        [MaxLength(18)]
        public string? NumRemProv { get; set; }
        /// <summary>
        /// Número de Factura del Proveedor
        /// </summary>
        [Column("NumFac_Prov")]
        [MaxLength(16)]
        public string? NumFacProv { get; set; }
        /// <summary>
        /// Número de Convenio Vigente
        /// </summary>
        [MaxLength(24)]
        public string? NumConVig { get; set; }
        /// <summary>
        /// Número de Orden de Compra
        /// </summary>
        public int? NumOC { get; set; }
        /// <summary>
        /// Número de Atención
        /// </summary>
        [Column("Nume_Aten")]
        public int? NumeAten { get; set; }
        /// <summary>
        /// Indicador Contable
        /// </summary>
        public short? IContab { get; set; }
        /// <summary>
        /// Moneda del Valor Unitario
        /// </summary>
        [MaxLength(4)]
        public string? MoneVaUni { get; set; }
        /// <summary>
        /// Valor Unitario
        /// </summary>
        [Column("Valo_Unit", TypeName = "decimal(18,4)")]
        public decimal? ValoUnit { get; set; }
        /// <summary>
        /// Cantidad Aprobada
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CanToApro { get; set; }
        /// <summary>
        /// Cantidad Rechazada
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CanToRech { get; set; }
        /// <summary>
        /// Cantidad Faltante
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? CanToFalta { get; set; }
        /// <summary>
        /// Fecha de Liberación
        /// </summary>
        public DateTime? FechLibe { get; set; }
        /// <summary>
        /// Número de Orden de Servicio
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? NumeOrdSer { get; set; }
        /// <summary>
        /// Observaciones
        /// </summary>
        public string? Observa { get; set; }
        /// <summary>
        /// Estado
        /// </summary>
        public short? Status { get; set; }
        /// <summary>
        /// Usuario que Aprobó
        /// </summary>
        [MaxLength(32)]
        public string? UsuApro { get; set; }
        /// <summary>
        /// Fecha y Hora de Aprobación
        /// </summary>
        public DateTime? FeHoApro { get; set; }
        /// <summary>
        /// Marca del Producto
        /// </summary>
        [MaxLength(32)]
        public string? Marca { get; set; }
        /// <summary>
        /// Número de Parte
        /// </summary>
        [MaxLength(32)]
        public string? NumeParte { get; set; }
        /// <summary>
        /// Fecha de Fabricación
        /// </summary>
        public DateTime? FechFab { get; set; }
        /// <summary>
        /// Plazo de Uso
        /// </summary>
        [Column("Plazo_Uso")]
        public DateTime? PlazoUso { get; set; }
        /// <summary>
        /// Trazabilidad del Cliente
        /// </summary>
        [Column("Traz_Cli")]
        [MaxLength(32)]
        public string? TrazCli { get; set; }
        /// <summary>
        /// Código del Cliente
        /// </summary>
        [Column("Cod_Cli")]
        [MaxLength(32)]
        public string? CodCli { get; set; }
        /// <summary>
        /// Tipo
        /// </summary>
        [MaxLength(32)]
        public string? Tipo { get; set; }
        /// <summary>
        /// Características
        /// </summary>
        [MaxLength(32)]
        public string? Carac { get; set; }
        /// <summary>
        /// Depósito Ubicación
        /// </summary>
        public short? DepoUbic { get; set; }
        /// <summary>
        /// Anotaciones de Orden de Fabricación
        /// </summary>
        public string? AnotaOrf { get; set; }
        /// <summary>
        /// Motivo de Rechazo
        /// </summary>
        public short? MotiRecha { get; set; }
        /// <summary>
        /// Observaciones de Rechazo
        /// </summary>
        public string? ObseRecha { get; set; }
        /// <summary>
        /// Destino de Rechazo
        /// </summary>
        public short? Destirecha { get; set; }
        /// <summary>
        /// GUID único del registro
        /// </summary>
        [Column("rowguid")]
        public Guid Rowguid { get; set; } = Guid.NewGuid();

        /// <summary>
        /// Vía de Entrega
        /// </summary>
        [Column("Via_Entrega")]
        public short? ViaEntrega { get; set; }
        /// <summary>
        /// Valor OBADA 1
        /// </summary>
        [Column("VALOBADA_1")]
        [MaxLength(256)]
        public string? Valobada1 { get; set; }

        /// <summary>
        /// Valor OBADA 2
        /// </summary>
        [Column("VALOBADA_2")]
        [MaxLength(256)]
        public string? Valobada2 { get; set; }

        /// <summary>
        /// Número de Item de Orden de Compra
        /// </summary>
        public short? NroItemOC { get; set; }

        /// <summary>
        /// Fecha y Hora de Recepción
        /// </summary>
        public DateTime? FEYHORecep { get; set; }

        /// <summary>
        /// Peso
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? Peso { get; set; }

        /// <summary>
        /// Proveedor Exterior
        /// </summary>
        public int? ProvExterior { get; set; }
        /// <summary>
        /// Porcentaje de Nacionalización
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PjeNacionaliz { get; set; }
        /// <summary>
        /// Vía Exterior
        /// </summary>
        public short? ViaExterior { get; set; }
        /// <summary>
        /// Patente del Transporte
        /// </summary>
        [MaxLength(16)]
        public string? PatenteTransporte { get; set; }

        /// <summary>
        /// Kilogramos de Balanza Pública
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? KgBalanzaPublica { get; set; }
        /// <summary>
        /// Cantidad de Bultos
        /// </summary>
        public short? CantidadBultos { get; set; }
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
/** 
 -- Agregar campo NumUsuar (Número de Usuario)
ALTER TABLE BOLRECEP
ADD NumUsuar INT NULL;



ALTER TABLE BOLRECEP ALTER COLUMN CodiEmpr INT;


-- 1.2. Cambiar Cod_Inte  a INT
ALTER TABLE BOLRECEP ALTER COLUMN Cod_Inte INT;

-- 1.3. Cambiar NumProv  a INT
ALTER TABLE BOLRECEP ALTER COLUMN NumProv INT;

-- 2.1. Cambiar NumeBoRe de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN NumeBoRe DECIMAL(18,4);

-- 2.2. Cambiar CantNom de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN CantNom DECIMAL(18,4);

-- 2.3. Cambiar M_Largo de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN M_Largo DECIMAL(18,4);

-- 2.4. Cambiar M_Ancho de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN M_Ancho DECIMAL(18,4);

-- 2.5. Cambiar Valo_Unit de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN Valo_Unit DECIMAL(18,4);

-- 2.6. Cambiar CanToApro de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN CanToApro DECIMAL(18,4);

-- 2.7. Cambiar CanToRech de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN CanToRech DECIMAL(18,4);

-- 2.8. Cambiar CanToFalta de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN CanToFalta DECIMAL(18,4);

-- 2.9. Cambiar NumeOrdSer de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN NumeOrdSer DECIMAL(18,4);

-- 2.10. Cambiar Peso de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN Peso DECIMAL(18,4);

-- 2.11. Cambiar PjeNacionaliz de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN PjeNacionaliz DECIMAL(18,4);

-- 2.12. Cambiar KgBalanzaPublica de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN KgBalanzaPublica DECIMAL(18,4);

ALTER TABLE BOLRECEP ADD Id INT IDENTITY(1,1) NOT NULL;

-- 0.2. Establecer Id como clave primaria
ALTER TABLE BOLRECEP ADD CONSTRAINT PK_BOLRECEP_Id PRIMARY KEY (Id);


ALTER TABLE BOLRECEP
ADD NumUsuar INT NULL;

-- 1. Agregar los campos permitiendo NULL temporalmente
ALTER TABLE BOLRECEP
ADD 
    CodiEmprNet INT NULL,
    LastUpdate DATETIME2(7) NULL,
    AddRecord DATETIME2(7) NULL;

-- 2. Actualizar los registros existentes con la fecha actual
UPDATE BOLRECEP
SET 
    LastUpdate = GETDATE(),
    AddRecord = GETDATE()
WHERE LastUpdate IS NULL OR AddRecord IS NULL;

-- 3. Cambiar los campos a NOT NULL con valor por defecto para futuros registros
ALTER TABLE BOLRECEP
ALTER COLUMN LastUpdate DATETIME2(7) NOT NULL;

ALTER TABLE BOLRECEP
ALTER COLUMN AddRecord DATETIME2(7) NOT NULL;

-- 4. Agregar constraint de valor por defecto para nuevos registros
ALTER TABLE BOLRECEP
ADD CONSTRAINT DF_BOLRECEP_LastUpdate DEFAULT GETDATE() FOR LastUpdate;

ALTER TABLE BOLRECEP
ADD CONSTRAINT DF_BOLRECEP_AddRecord DEFAULT GETDATE() FOR AddRecord;

SELECT 
    COLUMN_NAME,
    DATA_TYPE,
    CHARACTER_MAXIMUM_LENGTH,
    NUMERIC_PRECISION,
    NUMERIC_SCALE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'BOLRECEP'
    AND COLUMN_NAME IN (
        'CodiEmpr', 'Cod_Inte', 'NumProv',
        'NumeBoRe', 'CantNom', 'M_Largo', 'M_Ancho', 'Valo_Unit',
        'CanToApro', 'CanToRech', 'CanToFalta', 'NumeOrdSer',
        'Peso', 'PjeNacionaliz', 'KgBalanzaPublica'
    )
ORDER BY COLUMN_NAME;

-- =====================================================================================
-- SECCIÓN 3: CREAR ÍNDICES PARA OPTIMIZAR CONSULTAS
-- =====================================================================================
-- Se crean índices en campos frecuentemente usados en búsquedas y joins
-- Esto mejora significativamente el rendimiento de las consultas

-- 3.1. Índice en NumProv (Número de Proveedor)
-- Útil para consultas por proveedor
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_NumProv')
    CREATE INDEX IX_BOLRECEP_NumProv ON [dbo].[BOLRECEP] ([NumProv] ASC);

-- 3.2. Índice en Cod_Inte (Código Interno del Producto)
-- Útil para consultas por producto
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_CodInte')
    CREATE INDEX IX_BOLRECEP_CodInte ON [dbo].[BOLRECEP] ([Cod_Inte] ASC);

-- 3.3. Índice en IdenLote (Identificador de Lote)
-- Útil para trazabilidad y consultas por lote
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_IdenLote')
    CREATE INDEX IX_BOLRECEP_IdenLote ON [dbo].[BOLRECEP] ([IdenLote] ASC);

-- 3.4. Índice en NumOC (Número de Orden de Compra)
-- Útil para relacionar con órdenes de compra
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_NumOC')
    CREATE INDEX IX_BOLRECEP_NumOC ON [dbo].[BOLRECEP] ([NumOC] ASC);

-- 3.5. Índice en NumeBoRe (Número de Boleta de Recepción)
-- Útil para búsquedas por número de boleta
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_NumeBoRe')
    CREATE INDEX IX_BOLRECEP_NumeBoRe ON [dbo].[BOLRECEP] ([NumeBoRe] ASC);

-- 3.6. Índice en FechRecep (Fecha de Recepción)
-- Útil para consultas por rango de fechas
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_FechRecep')
    CREATE INDEX IX_BOLRECEP_FechRecep ON [dbo].[BOLRECEP] ([FechRecep] ASC);

-- 3.7. Índice compuesto en NumOC y Cod_Inte
-- Útil para consultas que filtran por orden de compra y producto
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_NumOC_CodInte')
    CREATE INDEX IX_BOLRECEP_NumOC_CodInte ON [dbo].[BOLRECEP] ([NumOC] ASC, [Cod_Inte] ASC);

-- 3.8. Índice compuesto en NumProv y FechRecep
-- Útil para consultas de recepciones por proveedor en un período
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_NumProv_FechRecep')
    CREATE INDEX IX_BOLRECEP_NumProv_FechRecep ON [dbo].[BOLRECEP] ([NumProv] ASC, [FechRecep] ASC);

-- =====================================================================================
-- FIN DE LAS INSTRUCCIONES
-- =====================================================================================
-- VERIFICACIÓN: Ejecutar después de los cambios para confirmar los nuevos tipos e índices
-- =====================================================================================

-- Verificar tipos de datos
SELECT 
    COLUMN_NAME,
    DATA_TYPE,
    CHARACTER_MAXIMUM_LENGTH,
    NUMERIC_PRECISION,
    NUMERIC_SCALE,
    IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'BOLRECEP'
    AND COLUMN_NAME IN (
        'Id', 'CodiEmpr', 'Cod_Inte', 'NumProv',
        'NumeBoRe', 'CantNom', 'M_Largo', 'M_Ancho', 'Valo_Unit',
        'CanToApro', 'CanToRech', 'CanToFalta', 'NumeOrdSer',
        'Peso', 'PjeNacionaliz', 'KgBalanzaPublica',
        'IdenLote', 'NumOC', 'FechRecep'
    )
ORDER BY COLUMN_NAME;

-- Verificar índices creados
SELECT 
    i.name AS IndexName,
    COL_NAME(ic.object_id, ic.column_id) AS ColumnName,
    i.type_desc AS IndexType,
    i.is_unique AS IsUnique
FROM sys.indexes i
INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
WHERE i.object_id = OBJECT_ID('BOLRECEP')
    AND i.name LIKE 'IX_BOLRECEP_%'
ORDER BY i.name, ic.key_ordinal;ALTER TABLE BOLRECEP ALTER COLUMN CodiEmpr INT;

-- 1.2. Cambiar Cod_Inte  a INT
ALTER TABLE BOLRECEP ALTER COLUMN Cod_Inte INT;

-- 1.3. Cambiar NumProv  a INT
ALTER TABLE BOLRECEP ALTER COLUMN NumProv INT;

-- 2.1. Cambiar NumeBoRe de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN NumeBoRe DECIMAL(18,4);

-- 2.2. Cambiar CantNom de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN CantNom DECIMAL(18,4);

-- 2.3. Cambiar M_Largo de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN M_Largo DECIMAL(18,4);

-- 2.4. Cambiar M_Ancho de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN M_Ancho DECIMAL(18,4);

-- 2.5. Cambiar Valo_Unit de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN Valo_Unit DECIMAL(18,4);

-- 2.6. Cambiar CanToApro de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN CanToApro DECIMAL(18,4);

-- 2.7. Cambiar CanToRech de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN CanToRech DECIMAL(18,4);

-- 2.8. Cambiar CanToFalta de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN CanToFalta DECIMAL(18,4);

-- 2.9. Cambiar NumeOrdSer de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN NumeOrdSer DECIMAL(18,4);

-- 2.10. Cambiar Peso de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN Peso DECIMAL(18,4);

-- 2.11. Cambiar PjeNacionaliz de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN PjeNacionaliz DECIMAL(18,4);

-- 2.12. Cambiar KgBalanzaPublica de FLOAT a DECIMAL(18,4)
ALTER TABLE BOLRECEP ALTER COLUMN KgBalanzaPublica DECIMAL(18,4);

ALTER TABLE BOLRECEP ADD Id INT IDENTITY(1,1) NOT NULL;

-- 0.2. Establecer Id como clave primaria
ALTER TABLE BOLRECEP ADD CONSTRAINT PK_BOLRECEP_Id PRIMARY KEY (Id);


SELECT 
    COLUMN_NAME,
    DATA_TYPE,
    CHARACTER_MAXIMUM_LENGTH,
    NUMERIC_PRECISION,
    NUMERIC_SCALE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'BOLRECEP'
    AND COLUMN_NAME IN (
        'CodiEmpr', 'Cod_Inte', 'NumProv',
        'NumeBoRe', 'CantNom', 'M_Largo', 'M_Ancho', 'Valo_Unit',
        'CanToApro', 'CanToRech', 'CanToFalta', 'NumeOrdSer',
        'Peso', 'PjeNacionaliz', 'KgBalanzaPublica'
    )
ORDER BY COLUMN_NAME;

-- =====================================================================================
-- SECCIÓN 3: CREAR ÍNDICES PARA OPTIMIZAR CONSULTAS
-- =====================================================================================
-- Se crean índices en campos frecuentemente usados en búsquedas y joins
-- Esto mejora significativamente el rendimiento de las consultas

-- 3.1. Índice en NumProv (Número de Proveedor)
-- Útil para consultas por proveedor
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_NumProv')
    CREATE INDEX IX_BOLRECEP_NumProv ON [dbo].[BOLRECEP] ([NumProv] ASC);

-- 3.2. Índice en Cod_Inte (Código Interno del Producto)
-- Útil para consultas por producto
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_CodInte')
    CREATE INDEX IX_BOLRECEP_CodInte ON [dbo].[BOLRECEP] ([Cod_Inte] ASC);

-- 3.3. Índice en IdenLote (Identificador de Lote)
-- Útil para trazabilidad y consultas por lote
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_IdenLote')
    CREATE INDEX IX_BOLRECEP_IdenLote ON [dbo].[BOLRECEP] ([IdenLote] ASC);

-- 3.4. Índice en NumOC (Número de Orden de Compra)
-- Útil para relacionar con órdenes de compra
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_NumOC')
    CREATE INDEX IX_BOLRECEP_NumOC ON [dbo].[BOLRECEP] ([NumOC] ASC);

-- 3.5. Índice en NumeBoRe (Número de Boleta de Recepción)
-- Útil para búsquedas por número de boleta
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_NumeBoRe')
    CREATE INDEX IX_BOLRECEP_NumeBoRe ON [dbo].[BOLRECEP] ([NumeBoRe] ASC);

-- 3.6. Índice en FechRecep (Fecha de Recepción)
-- Útil para consultas por rango de fechas
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_FechRecep')
    CREATE INDEX IX_BOLRECEP_FechRecep ON [dbo].[BOLRECEP] ([FechRecep] ASC);

-- 3.7. Índice compuesto en NumOC y Cod_Inte
-- Útil para consultas que filtran por orden de compra y producto
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_NumOC_CodInte')
    CREATE INDEX IX_BOLRECEP_NumOC_CodInte ON [dbo].[BOLRECEP] ([NumOC] ASC, [Cod_Inte] ASC);

-- 3.8. Índice compuesto en NumProv y FechRecep
-- Útil para consultas de recepciones por proveedor en un período
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID(N'[dbo].[BOLRECEP]') AND name = N'IX_BOLRECEP_NumProv_FechRecep')
    CREATE INDEX IX_BOLRECEP_NumProv_FechRecep ON [dbo].[BOLRECEP] ([NumProv] ASC, [FechRecep] ASC);

-- =====================================================================================
-- FIN DE LAS INSTRUCCIONES
-- =====================================================================================
-- VERIFICACIÓN: Ejecutar después de los cambios para confirmar los nuevos tipos e índices
-- =====================================================================================
/*
-- Verificar tipos de datos
SELECT 
    COLUMN_NAME,
    DATA_TYPE,
    CHARACTER_MAXIMUM_LENGTH,
    NUMERIC_PRECISION,
    NUMERIC_SCALE,
    IS_NULLABLE
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'BOLRECEP'
    AND COLUMN_NAME IN (
        'Id', 'CodiEmpr', 'Cod_Inte', 'NumProv',
        'NumeBoRe', 'CantNom', 'M_Largo', 'M_Ancho', 'Valo_Unit',
        'CanToApro', 'CanToRech', 'CanToFalta', 'NumeOrdSer',
        'Peso', 'PjeNacionaliz', 'KgBalanzaPublica',
        'IdenLote', 'NumOC', 'FechRecep'
    )
ORDER BY COLUMN_NAME;

-- Verificar índices creados
SELECT 
    i.name AS IndexName,
    COL_NAME(ic.object_id, ic.column_id) AS ColumnName,
    i.type_desc AS IndexType,
    i.is_unique AS IsUnique
FROM sys.indexes i
INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
WHERE i.object_id = OBJECT_ID('BOLRECEP')
    AND i.name LIKE 'IX_BOLRECEP_%'
ORDER BY i.name, ic.key_ordinal;
*/