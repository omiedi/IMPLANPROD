using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad que representa un pedido en el sistema.
    /// Mapea a la tabla PEDENCA en la base de datos.
    /// </summary>
    [Table("PEDENCA")]
    public class Pedenca : IAuditableEntity
    {
        /// <summary>
        /// Identificador interno autoincremental (IDENTITY) y clave primaria.
        /// Se agrega para usar una PK numérica por rendimiento y mantener NroPed como dato de negocio.
        /// </summary>
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Id { get; set; }

        /// <summary>
        /// Número de pedido
        /// </summary>
        public int? NroPed { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        public int? CodiEmpr { get; set; } = 0;

        /// <summary>
        /// Referencia del pedido
        /// </summary>
        [StringLength(128)]
        public string? Referencia { get; set; }

        /// <summary>
        /// Tipo de pedido
        /// </summary>
        [StringLength(8)]
        public string? TipoPed { get; set; }

        /// <summary>
        /// Fecha de emisión del pedido
        /// </summary>
        public DateTime? FechEmis { get; set; }

        /// <summary>
        /// Número de cliente
        /// </summary>
        public int? NroClie { get; set; }

        /// <summary>
        /// Razón social del cliente
        /// </summary>
        [StringLength(64)]
        public string? RasoClie { get; set; }

        /// <summary>
        /// Número de orden de compra
        /// </summary>
        [StringLength(32)]
        public string? NroOcom { get; set; }

        /// <summary>
        /// Domicilio de entrega
        /// </summary>
        [StringLength(48)]
        public string? DomiEnt { get; set; }

        /// <summary>
        /// Localidad de entrega
        /// </summary>
        [StringLength(48)]
        public string? LocaEnt { get; set; }

        /// <summary>
        /// Transporte de entrega
        /// </summary>
        [StringLength(48)]
        public string? TranspEnt { get; set; }

        /// <summary>
        /// Número de transporte
        /// </summary>
        public short? NTranspo { get; set; }

        /// <summary>
        /// Número de sucursal de entrega
        /// </summary>
        public short? NSucEnt { get; set; }

        /// <summary>
        /// Observaciones del pedido
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? Observa { get; set; }

        /// <summary>
        /// Fecha solicitada de entrega
        /// </summary>
        public DateTime? FeSolEnt { get; set; }

        /// <summary>
        /// Fin de vigencia del pedido
        /// </summary>
        public DateTime? FinVigen { get; set; }

        /// <summary>
        /// Estado del pedido
        /// </summary>
        public short? Status { get; set; }

        /// <summary>
        /// Vendedor representante
        /// </summary>
        [StringLength(8)]
        public string? VenRep { get; set; }

        /// <summary>
        /// Lista de precios
        /// </summary>
        public short? ListPre { get; set; }

        /// <summary>
        /// Condición de pago
        /// </summary>
        public short? CondPag { get; set; }

        /// <summary>
        /// Vendedor asignado
        /// </summary>
        public short? Vended { get; set; }

        /// <summary>
        /// Porcentaje de comisión del vendedor
        /// </summary>
        [Column(TypeName = "decimal(18,4)")]
        public decimal? PorComVend { get; set; }

        /// <summary>
        /// Tipo de descuento aplicado
        /// </summary>
        [StringLength(14)]
        public string? TDescApli { get; set; }

        /// <summary>
        /// Anotaciones del pedido
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? AnotaPed { get; set; }

        /// <summary>
        /// Fecha de aprobación del pedido
        /// </summary>
        public DateTime? FApro_Nope { get; set; }

        /// <summary>
        /// Usuario que aprobó el pedido
        /// </summary>
        [Column("Usua_Apro ")]
        public int? Usua_Apro { get; set; }

        /// <summary>
        /// Número de usuario
        /// </summary>
        public int? NumUsuar { get; set; }

        /// <summary>
        /// Recargo con pago
        /// </summary>
        [StringLength(14)]
        public string? Rec_ConPa { get; set; }

        /// <summary>
        /// Motivo de orden de venta
        /// </summary>
        public short? OV_Motivo { get; set; }

        /// <summary>
        /// Destino de orden de venta
        /// </summary>
        [StringLength(128)]
        public string? OV_Destino { get; set; }

        /// <summary>
        /// Vehículo de orden de venta
        /// </summary>
        [StringLength(128)]
        public string? OV_Vehiculo { get; set; }

        /// <summary>
        /// Fecha de salida de orden de venta
        /// </summary>
        public DateTime? OV_Fesalida { get; set; }

        /// <summary>
        /// Fecha de regreso de orden de venta
        /// </summary>
        public DateTime? OV_FeRegreso { get; set; }

        /// <summary>
        /// Condiciones de despacho de orden de venta
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? OV_CondDespa { get; set; }

        /// <summary>
        /// Primera autorización de orden de venta
        /// </summary>
        [StringLength(64)]
        public string? OV_Autoriza1 { get; set; }

        /// <summary>
        /// Segunda autorización de orden de venta
        /// </summary>
        [StringLength(64)]
        public string? OV_Autoriza2 { get; set; }

        /// <summary>
        /// Observaciones de orden de venta
        /// </summary>
        [Column(TypeName = "ntext")]
        public string? OV_Observaciones { get; set; }

        /// <summary>
        /// Número de serie de devolución de orden de venta
        /// </summary>
        [StringLength(24)]
        public string? OV_NSerieDev { get; set; }

        /// <summary>
        /// Código de carpeta de impresión
        /// </summary>
        [StringLength(12)]
        public string? CodCarpImp { get; set; }

        /// <summary>
        /// Tipo de ingreso
        /// </summary>
        public short? TipoIngre { get; set; }

        /// <summary>
        /// Descuento de retención del cliente
        /// </summary>
        [StringLength(14)]
        public string? Dto_Ret_Cte { get; set; }

        /// <summary>
        /// Descuento por cantidad
        /// </summary>
        [StringLength(14)]
        public string? DtoxCant { get; set; }

        /// <summary>
        /// Texto de flete
        /// </summary>
        [StringLength(64)]
        public string? FLETE_TEX { get; set; }

        /// <summary>
        /// Fecha de anulación del pedido
        /// </summary>
        public DateTime? FechAnul { get; set; }

        /// <summary>
        /// Número de revisión
        /// </summary>
        public short? Revision { get; set; }

        /// <summary>
        /// Fecha del último cambio
        /// </summary>
        public DateTime? FechUltCambio { get; set; }

        /// <summary>
        /// Atributo comercial
        /// </summary>
        public short? AtributoComercial { get; set; }

        /// <summary>
        /// Moneda de emisión (por defecto 1 = pesos)
        /// </summary>
        [NotMapped] // Esta propiedad no se mapeará a la base de datos
        public short? MoneEmis { get; set; } = 1;

        /// <summary>
        /// Tipo de cambio
        /// </summary>
        [NotMapped] // Esta propiedad no se mapeará a la base de datos
        [Column(TypeName = "decimal(18,4)")]
        public decimal? TipoCam { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        public int? CodiEmprNet { get; set; } = 0;

        [Display(Name = "Fecha de Creación")]
        public DateTime? AddRecord { get; set; }

        [Display(Name = "Última Actualización")]
        public DateTime? LastUpdate { get; set; }
        
        /// <summary>
        /// Identificador único global de la fila
        /// </summary>
        [Column("rowguid")]
        public Guid Rowguid { get; set; } = Guid.NewGuid();
}
/*
-- =====================================================================================
-- INSTRUCCIONES SQL PARA MODIFICAR LA TABLA PEDENCA (AGREGAR ID IDENTITY Y MOVER PK)
-- Ejecutar en SQL Server Management Studio para sincronizar con la entidad C#
-- =====================================================================================

-- PASO 0: Agregar columna Id como IDENTITY y establecerla como clave primaria
-- ---------------------------------------------------------------------------------
IF COL_LENGTH('dbo.PEDENCA', 'Id') IS NULL
BEGIN
    ALTER TABLE [dbo].[PEDENCA]
    ADD [Id] INT IDENTITY(1,1) NOT NULL;
END

-- Eliminar PK existente sobre NroPed (si aplica) y crear PK en Id
DECLARE @pkName sysname;
SELECT @pkName = kc.name
FROM sys.key_constraints kc
JOIN sys.tables t ON kc.parent_object_id = t.object_id
JOIN sys.index_columns ic ON kc.parent_object_id = ic.object_id AND kc.unique_index_id = ic.index_id
JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id
WHERE kc.type = 'PK' AND t.name = 'PEDENCA' AND c.name = 'NroPed';

IF @pkName IS NOT NULL
BEGIN
    PRINT 'Eliminando PK existente ' + @pkName + ' sobre NroPed...';
    DECLARE @sql nvarchar(4000);
    SET @sql = N'ALTER TABLE [dbo].[PEDENCA] DROP CONSTRAINT [' + @pkName + N']';
    EXEC sp_executesql @sql;
END

-- Si no existe PK en Id, crearla
IF NOT EXISTS (
    SELECT 1
    FROM sys.key_constraints kc
    JOIN sys.index_columns ic ON kc.parent_object_id = ic.object_id AND kc.unique_index_id = ic.index_id
    JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id
    WHERE kc.type = 'PK' AND OBJECT_NAME(kc.parent_object_id) = 'PEDENCA' AND c.name = 'Id'
)
BEGIN
    ALTER TABLE [dbo].[PEDENCA] ADD CONSTRAINT [PK_PEDENCA_Id] PRIMARY KEY CLUSTERED ([Id]);
END

-- (Opcional) Crear índice único en NroPed si debe mantenerse único como dato de negocio
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID('PEDENCA') AND name = 'UX_PEDENCA_NroPed')
BEGIN
    CREATE UNIQUE NONCLUSTERED INDEX [UX_PEDENCA_NroPed] ON [dbo].[PEDENCA]([NroPed]) WHERE [NroPed] IS NOT NULL;
END

-- =====================================================================================
-- INSTRUCCIONES ADICIONALES (ya existentes)
-- =====================================================================================

-- PASO 1: Cambiar tipos de datos de campos existentes
-- ----------------------------------------------------

-- 1.1. Cambiar CodiEmpr de smallint a int
ALTER TABLE [dbo].[PEDENCA] ALTER COLUMN [CodiEmpr] int;

-- 1.2. Cambiar Usua_Apro de smallint a int  
ALTER TABLE [dbo].[PEDENCA] ALTER COLUMN [Usua_Apro ] int;

-- 1.3. Cambiar PorComVend de float a decimal(18,4)
ALTER TABLE [dbo].[PEDENCA] ALTER COLUMN [PorComVend] decimal(18,4);

-- PASO 2: Agregar nuevos campos
-- -----------------------------

-- 2.1. Agregar campo NumUsuar como int
ALTER TABLE [dbo].[PEDENCA] ADD [NumUsuar] int;

-- PASO 3: Verificar y mantener campo rowguid
-- -----------------------------------------

-- 3.1. Verificar si existe el campo rowguid
IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS 
               WHERE TABLE_NAME = 'PEDENCA' AND COLUMN_NAME = 'rowguid')
BEGIN
    -- 3.2. Agregar campo rowguid si no existe
    ALTER TABLE [dbo].[PEDENCA] ADD [rowguid] UNIQUEIDENTIFIER ROWGUIDCOL NOT NULL DEFAULT NEWID();
END

-- PASO 4: Crear índices recomendados para mejorar rendimiento
-- ----------------------------------------------------------

-- 4.1. Índice en NroPed (si no se usa único, puede ser no clustered estándar)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID('PEDENCA') AND name = 'IX_PEDENCA_NroPed')
CREATE INDEX IX_PEDENCA_NroPed ON [dbo].[PEDENCA] ([NroPed]);

-- 4.2. Índice en NroClie (para consultas por cliente)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID('PEDENCA') AND name = 'IX_PEDENCA_NroClie')
CREATE INDEX IX_PEDENCA_NroClie ON [dbo].[PEDENCA] ([NroClie]);

-- 4.3. Índice compuesto en CodiEmpr y Status (para consultas empresariales)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID('PEDENCA') AND name = 'IX_PEDENCA_Empresa_Status')
CREATE INDEX IX_PEDENCA_Empresa_Status ON [dbo].[PEDENCA] ([CodiEmpr], [Status]);

-- 4.4. Índice en FechEmis (para consultas por fecha)
IF NOT EXISTS (SELECT * FROM sys.indexes WHERE object_id = OBJECT_ID('PEDENCA') AND name = 'IX_PEDENCA_FechEmis')
CREATE INDEX IX_PEDENCA_FechEmis ON [dbo].[PEDENCA] ([FechEmis]);

-- PASO 5: Verificación de cambios aplicados
-- -----------------------------------------

-- 5.1. Verificar estructura de la tabla actualizada
SELECT 
    COLUMN_NAME,
    DATA_TYPE,
    IS_NULLABLE,
    CHARACTER_MAXIMUM_LENGTH,
    NUMERIC_PRECISION,
    NUMERIC_SCALE
FROM INFORMATION_SCHEMA.COLUMNS 
WHERE TABLE_NAME = 'PEDENCA'
ORDER BY ORDINAL_POSITION;

-- 5.2. Verificar índices creados
SELECT 
    i.name AS IndexName,
    i.type_desc AS IndexType,
    c.name AS ColumnName
FROM sys.indexes i
INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
INNER JOIN sys.columns c ON ic.object_id = c.object_id AND ic.column_id = c.column_id
WHERE i.object_id = OBJECT_ID('PEDENCA')
ORDER BY i.name, ic.key_ordinal;

-- =====================================================================================
-- NOTAS IMPORTANTES:
-- =====================================================================================
-- 1. Ejecutar los comandos en orden secuencial
-- 2. Realizar backup de la base de datos antes de ejecutar cambios estructurales
-- 3. Verificar que no hay transacciones activas en la tabla antes de modificar estructura
-- 4. Los índices son opcionales pero recomendados para mejorar el rendimiento
-- 5. El campo rowguid solo debe eliminarse si no existen dependencias
-- =====================================================================================
*/


////---------------------------------------------------
////--SI QUISIERA ELMINAR LOS TRIGGERS Eliminar trigger de PEDENCA **** SIN BORRAR LA TABLA DE AUDITORIA
////---------------------------------------------------
////IF OBJECT_ID('trg_Auditoria_PEDENCA', 'TR') IS NOT NULL
////BEGIN
////    DROP TRIGGER trg_Auditoria_PEDENCA;
////PRINT 'Trigger trg_Auditoria_PEDENCA eliminado.';
////END
////ELSE
////BEGIN
////    PRINT 'Trigger trg_Auditoria_PEDENCA no existe.';
////END
////GO

////---------------------------------------------------
////-- Eliminar trigger de PEDDETA
////---------------------------------------------------
////IF OBJECT_ID('trg_Auditoria_PEDDETA', 'TR') IS NOT NULL
////BEGIN
////    DROP TRIGGER trg_Auditoria_PEDDETA;
////PRINT 'Trigger trg_Auditoria_PEDDETA eliminado.';
////END
////ELSE
////BEGIN
////    PRINT 'Trigger trg_Auditoria_PEDDETA no existe.';
////END
////GO
///


//ALTER TABLE PEDENCA
//ADD     AddRecord DATETIME2 NOT NULL DEFAULT GETDATE(),
//LastUpdate DATETIME2 NOT NULL DEFAULT GETDATE(),
//CodiEmprNet INT NOT NULL DEFAULT 0;
}