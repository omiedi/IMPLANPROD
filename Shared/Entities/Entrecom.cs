using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace IMPLANPROD.Shared.Entities
{
    /// <summary>
    /// Entidad ENTRECOM - Entregas de compras
    /// Representa las entregas de productos comprados
    /// </summary>
    [Table("ENTRECOM")]
    public class Entrecom : ITimestampEntity, IAuditableEntity
    {
        /// <summary>
        /// Identificador único del registro
        /// </summary>
        [Key]
        public int Id { get; set; }

        /// <summary>
        /// Fecha y hora de la entrega
        /// </summary>
        [Display(Name = "Fecha")]
        public DateTime? Fecha { get; set; }

        /// <summary>
        /// Código de empresa
        /// </summary>
        [Display(Name = "Código Empresa")]
        public short? CodiEmpr { get; set; }

        /// <summary>
        /// Número de documento
        /// </summary>
        [Display(Name = "Documento")]
        [MaxLength(14)]
        public string? Documento { get; set; }

        /// <summary>
        /// Tipo de documento
        /// </summary>
        [Display(Name = "Tipo Documento")]
        [MaxLength(16)]
        public string? TipoDoc { get; set; }

        /// <summary>
        /// Número de cliente
        /// </summary>
        [Display(Name = "Número Cliente")]
        public int? Nume_Cli { get; set; }

        /// <summary>
        /// ID del documento
        /// </summary>
        [Display(Name = "ID Documento")]
        public int? IDocum { get; set; }

        /// <summary>
        /// Código interno del producto
        /// </summary>
        [Display(Name = "Código Interno")]
        public int? Cod_Inte { get; set; }

        /// <summary>
        /// Código externo del producto
        /// </summary>
        [Display(Name = "Código Externo")]
        [MaxLength(24)]
        public string? Cod_Exte { get; set; }

        /// <summary>
        /// Cantidad entregada
        /// </summary>
        [Display(Name = "Cantidad")]
        [Column(TypeName = "decimal(18,4)")]
        public decimal? Cantidad { get; set; }

        /// <summary>
        /// Estado de la entrega
        /// </summary>
        [Display(Name = "Estado")]
        public short? Status { get; set; }

        /// <summary>
        /// Cantidad de veces que se envió
        /// </summary>
        [Display(Name = "Veces Enviado")]
        public short? VecesEnviado { get; set; }

        // ===================================================================
        // CAMPOS DE TIMESTAMP AUTOMÁTICOS - Implementación de ITimestampEntity
        // ===================================================================

        /// <summary>
        /// Fecha de creación del registro
        /// Se establece automáticamente cuando se crea la entrega
        /// </summary>
        [Column("AddRecord")]
        public DateTime? AddRecord { get; set; }

        /// <summary>
        /// Fecha de última actualización del registro
        /// Se actualiza automáticamente cada vez que se modifica la entrega
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
===============================================================================
INSTRUCCIONES SQL PARA MODIFICAR LA TABLA ENTRECOM
===============================================================================

-- 1. Modificar el tipo de dato de Fecha de smalldatetime a datetime2 (soporta día y hora con mayor precisión)
ALTER TABLE [dbo].[ENTRECOM]
ALTER COLUMN [Fecha] datetime2 NULL;

-- 2. Modificar Nume_Cli de smallint a int
ALTER TABLE [dbo].[ENTRECOM]
ALTER COLUMN [Nume_Cli] int NULL;

-- 3. Modificar Cod_Inte de smallint a int
ALTER TABLE [dbo].[ENTRECOM]
ALTER COLUMN [Cod_Inte] int NULL;

-- 4. Modificar Cod_Exte de nvarchar(16) a nvarchar(24)
ALTER TABLE [dbo].[ENTRECOM]
ALTER COLUMN [Cod_Exte] nvarchar(24) NULL;

-- 5. Modificar Cantidad de float a decimal(18,4)
ALTER TABLE [dbo].[ENTRECOM]
ALTER COLUMN [Cantidad] decimal(18,4) NULL;

-- 6. Agregar columna Id como clave primaria (si no existe)
-- NOTA: Solo ejecutar si la tabla no tiene clave primaria
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'Id')
BEGIN
    ALTER TABLE [dbo].[ENTRECOM]
    ADD [Id] int IDENTITY(1,1) NOT NULL;
    
    ALTER TABLE [dbo].[ENTRECOM]
    ADD CONSTRAINT PK_ENTRECOM PRIMARY KEY CLUSTERED ([Id]);
END
GO

===============================================================================
INSTRUCCIONES SQL PARA AGREGAR CAMPOS DE AUDITORÍA
===============================================================================

-- 7. Agregar campo AddRecord (fecha de creación)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'AddRecord')
BEGIN
    ALTER TABLE [dbo].[ENTRECOM]
    ADD [AddRecord] datetime2 NULL DEFAULT GETDATE();
END
GO

-- 8. Agregar campo LastUpdate (fecha de última actualización)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'LastUpdate')
BEGIN
    ALTER TABLE [dbo].[ENTRECOM]
    ADD [LastUpdate] datetime2 NULL;
END
GO

-- 9. Agregar campo CodiEmprNet (código de empresa del usuario)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'CodiEmprNet')
BEGIN
    ALTER TABLE [dbo].[ENTRECOM]
    ADD [CodiEmprNet] int NULL;
END
GO

-- 10. Agregar campo NumUsuar (número de usuario)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'NumUsuar')
BEGIN
    ALTER TABLE [dbo].[ENTRECOM]
    ADD [NumUsuar] int NULL;
END
GO

===============================================================================
SCRIPT COMPLETO DE MIGRACIÓN (EJECUTAR EN ORDEN)
===============================================================================

-- Ejecutar todo el script en una transacción para poder hacer rollback si hay errores
BEGIN TRANSACTION;

BEGIN TRY
    -- Modificar tipos de datos existentes
    ALTER TABLE [dbo].[ENTRECOM] ALTER COLUMN [Fecha] datetime2 NULL;
    ALTER TABLE [dbo].[ENTRECOM] ALTER COLUMN [Nume_Cli] int NULL;
    ALTER TABLE [dbo].[ENTRECOM] ALTER COLUMN [Cod_Inte] int NULL;
    ALTER TABLE [dbo].[ENTRECOM] ALTER COLUMN [Cod_Exte] nvarchar(24) NULL;
    ALTER TABLE [dbo].[ENTRECOM] ALTER COLUMN [Cantidad] decimal(18,4) NULL;
    
    -- Agregar Id si no existe
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'Id')
    BEGIN
        ALTER TABLE [dbo].[ENTRECOM] ADD [Id] int IDENTITY(1,1) NOT NULL;
        ALTER TABLE [dbo].[ENTRECOM] ADD CONSTRAINT PK_ENTRECOM PRIMARY KEY CLUSTERED ([Id]);
    END
    
    -- Agregar campos de auditoría
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'AddRecord')
        ALTER TABLE [dbo].[ENTRECOM] ADD [AddRecord] datetime2 NULL DEFAULT GETDATE();
    
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'LastUpdate')
        ALTER TABLE [dbo].[ENTRECOM] ADD [LastUpdate] datetime2 NULL;
    
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'CodiEmprNet')
        ALTER TABLE [dbo].[ENTRECOM] ADD [CodiEmprNet] int NULL;
    
    IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'NumUsuar')
        ALTER TABLE [dbo].[ENTRECOM] ADD [NumUsuar] int NULL;
    
    -- Si todo salió bien, confirmar los cambios
    COMMIT TRANSACTION;
    PRINT 'Migración de tabla ENTRECOM completada exitosamente.';
END TRY
BEGIN CATCH
    -- Si hubo algún error, deshacer todos los cambios
    ROLLBACK TRANSACTION;
    PRINT 'Error en la migración: ' + ERROR_MESSAGE();
    THROW;
END CATCH;
GO

===============================================================================
VERIFICACIÓN DE LA ESTRUCTURA FINAL
===============================================================================

-- Verificar la estructura de la tabla después de la migración
SELECT 
    COLUMN_NAME,
    DATA_TYPE,
    CHARACTER_MAXIMUM_LENGTH,
    IS_NULLABLE,
    COLUMN_DEFAULT
FROM INFORMATION_SCHEMA.COLUMNS
WHERE TABLE_NAME = 'ENTRECOM'
ORDER BY ORDINAL_POSITION;
GO

===============================================================================
INSTRUCCIONES SQL PARA AGREGAR CAMPO Cantingre
===============================================================================

-- 11. Agregar campo Cantingre (cantidad ingresada) de tipo decimal(18,4)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'Cantingre')
BEGIN
    ALTER TABLE [dbo].[ENTRECOM]
    ADD [Cantingre] decimal(18,4) NULL;
END
GO

-- 12. Agregar campo CantSalid (cantidad salida) de tipo decimal(18,4)
IF NOT EXISTS (SELECT * FROM sys.columns WHERE object_id = OBJECT_ID(N'[dbo].[ENTRECOM]') AND name = 'CantSalid')
BEGIN
    ALTER TABLE [dbo].[ENTRECOM]
    ADD [CantSalid] decimal(18,4) NULL;
END
GO


===============================================================================
NOTAS IMPORTANTES
===============================================================================

1. **Backup**: Hacer backup de la base de datos antes de ejecutar estos scripts
2. **Datos existentes**: Los scripts preservan todos los datos existentes
3. **Conversiones**: 
   - smalldatetime a datetime2: Conversión automática sin pérdida de datos
   - smallint a int: Conversión automática sin pérdida de datos
   - float a decimal(18,4): Puede haber pequeños ajustes de redondeo
   - nvarchar(16) a nvarchar(24): Sin pérdida de datos
4. **Clave primaria**: Si la tabla ya tiene una clave primaria, comentar esa sección
5. **Auditoría automática**: Los campos de auditoría se actualizarán automáticamente
   mediante el AuditInterceptor configurado en el DataContext

===============================================================================
*/
